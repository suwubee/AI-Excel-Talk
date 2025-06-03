"""
AI Excel 智能分析工具 - 用户会话管理器
处理多用户支持、文件隔离和隐私保护
"""

import os
import uuid
import json
import hashlib
import tempfile
import shutil
from datetime import datetime, timedelta
from typing import Dict, Any, Optional, List
import threading
import time
import atexit
import logging
from pathlib import Path

class UserSessionManager:
    """用户会话管理器"""
    
    def __init__(self, base_upload_dir: str = "user_uploads", 
                 session_timeout_hours: int = 24,
                 cleanup_interval_minutes: int = 60):
        """
        初始化用户会话管理器
        
        Args:
            base_upload_dir: 用户文件存储基础目录
            session_timeout_hours: 会话超时时间（小时）
            cleanup_interval_minutes: 清理任务执行间隔（分钟）
        """
        self.base_upload_dir = Path(base_upload_dir)
        self.session_timeout = timedelta(hours=session_timeout_hours)
        self.cleanup_interval = cleanup_interval_minutes * 60
        
        # 创建基础目录
        self.base_upload_dir.mkdir(exist_ok=True)
        
        # 会话信息存储
        self.sessions_file = self.base_upload_dir / "sessions.json"
        self.sessions_lock = threading.Lock()
        
        # 初始化日志
        self.logger = logging.getLogger(__name__)
        
        # 启动清理任务
        self._start_cleanup_task()
        
        # 注册程序退出时的清理
        atexit.register(self.cleanup_on_exit)
    
    def generate_user_session_id(self, identifier: str = None) -> str:
        """
        生成用户会话ID
        
        Args:
            identifier: 可选的用户标识符（如IP地址等）
        
        Returns:
            唯一的会话ID
        """
        if identifier:
            # 基于标识符生成更稳定的会话ID
            hash_object = hashlib.md5(f"{identifier}_{datetime.now().strftime('%Y%m%d')}".encode())
            base_id = hash_object.hexdigest()[:16]
        else:
            # 生成随机会话ID
            base_id = str(uuid.uuid4()).replace('-', '')[:16]
        
        # 确保唯一性
        timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
        return f"user_{base_id}_{timestamp}"
    
    def create_user_workspace(self, session_id: str) -> Path:
        """
        为用户创建独立的工作空间
        
        Args:
            session_id: 用户会话ID
        
        Returns:
            用户工作空间路径
        """
        user_dir = self.base_upload_dir / session_id
        user_dir.mkdir(exist_ok=True)
        
        # 创建子目录
        (user_dir / "uploads").mkdir(exist_ok=True)
        (user_dir / "exports").mkdir(exist_ok=True)
        (user_dir / "temp").mkdir(exist_ok=True)
        
        # 记录会话信息
        self._update_session_info(session_id, {
            'created_at': datetime.now().isoformat(),
            'last_access': datetime.now().isoformat(),
            'workspace_path': str(user_dir),
            'file_count': 0
        })
        
        return user_dir
    
    def get_user_workspace(self, session_id: str) -> Optional[Path]:
        """
        获取用户工作空间路径
        
        Args:
            session_id: 用户会话ID
        
        Returns:
            用户工作空间路径，如果不存在则返回None
        """
        user_dir = self.base_upload_dir / session_id
        if user_dir.exists():
            # 更新最后访问时间
            self._update_session_access(session_id)
            return user_dir
        return None
    
    def save_uploaded_file(self, session_id: str, uploaded_file, filename: str = None) -> Path:
        """
        保存用户上传的文件到隔离的工作空间
        
        Args:
            session_id: 用户会话ID
            uploaded_file: Streamlit上传的文件对象
            filename: 可选的文件名
        
        Returns:
            保存的文件路径
        """
        # 获取或创建用户工作空间
        workspace = self.get_user_workspace(session_id)
        if not workspace:
            workspace = self.create_user_workspace(session_id)
        
        # 生成安全的文件名
        if not filename:
            filename = uploaded_file.name
        
        safe_filename = self._sanitize_filename(filename)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        unique_filename = f"{timestamp}_{safe_filename}"
        
        file_path = workspace / "uploads" / unique_filename
        
        # 保存文件
        with open(file_path, "wb") as f:
            f.write(uploaded_file.read())
        
        # 更新会话信息
        self._increment_file_count(session_id)
        
        self.logger.info(f"文件已保存: {file_path} (会话: {session_id})")
        return file_path
    
    def get_export_path(self, session_id: str, filename: str) -> Path:
        """
        获取用户导出文件的路径
        
        Args:
            session_id: 用户会话ID
            filename: 导出文件名
        
        Returns:
            导出文件路径
        """
        workspace = self.get_user_workspace(session_id)
        if not workspace:
            workspace = self.create_user_workspace(session_id)
        
        safe_filename = self._sanitize_filename(filename)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        export_filename = f"{timestamp}_{safe_filename}"
        
        return workspace / "exports" / export_filename
    
    def get_temp_path(self, session_id: str, filename: str = None) -> Path:
        """
        获取用户临时文件路径
        
        Args:
            session_id: 用户会话ID
            filename: 可选的文件名
        
        Returns:
            临时文件路径
        """
        workspace = self.get_user_workspace(session_id)
        if not workspace:
            workspace = self.create_user_workspace(session_id)
        
        if filename:
            safe_filename = self._sanitize_filename(filename)
            return workspace / "temp" / safe_filename
        else:
            # 生成随机临时文件名
            temp_filename = f"temp_{uuid.uuid4().hex[:8]}.tmp"
            return workspace / "temp" / temp_filename
    
    def get_user_excel_files(self, session_id: str) -> List[Dict[str, Any]]:
        """
        获取用户已上传的Excel文件列表
        
        Args:
            session_id: 用户会话ID
        
        Returns:
            Excel文件信息列表，包含文件名、路径、上传时间等
        """
        workspace = self.get_user_workspace(session_id)
        if not workspace:
            return []
        
        uploads_dir = workspace / "uploads"
        if not uploads_dir.exists():
            return []
        
        excel_files = []
        excel_extensions = ['.xlsx', '.xls', '.xlsm', '.xlsb']
        
        try:
            for file_path in uploads_dir.iterdir():
                if file_path.is_file() and file_path.suffix.lower() in excel_extensions:
                    stat_info = file_path.stat()
                    
                    # 解析文件名中的时间戳（如果存在）
                    filename = file_path.name
                    display_name = filename
                    
                    # 如果文件名包含时间戳前缀，提取原始文件名
                    if '_' in filename and len(filename.split('_')[0]) == 15:
                        # 格式：20241201_123456_原始文件名.xlsx
                        parts = filename.split('_', 2)
                        if len(parts) >= 3:
                            display_name = parts[2]
                    
                    excel_files.append({
                        'filename': filename,  # 实际文件名
                        'display_name': display_name,  # 显示名称
                        'path': str(file_path),
                        'size': stat_info.st_size,
                        'modified_time': datetime.fromtimestamp(stat_info.st_mtime),
                        'size_mb': round(stat_info.st_size / (1024 * 1024), 2)
                    })
            
            # 按修改时间排序，最新的在前
            excel_files.sort(key=lambda x: x['modified_time'], reverse=True)
            
        except Exception as e:
            self.logger.error(f"获取Excel文件列表失败 {session_id}: {e}")
        
        return excel_files
    
    def get_user_file_by_name(self, session_id: str, filename: str) -> Optional[Path]:
        """
        根据文件名获取用户文件路径
        
        Args:
            session_id: 用户会话ID
            filename: 文件名（实际文件名或显示名称）
        
        Returns:
            文件路径，如果不存在则返回None
        """
        workspace = self.get_user_workspace(session_id)
        if not workspace:
            return None
        
        uploads_dir = workspace / "uploads"
        if not uploads_dir.exists():
            return None
        
        try:
            # 首先尝试直接匹配文件名
            direct_path = uploads_dir / filename
            if direct_path.exists():
                return direct_path
            
            # 如果直接匹配失败，搜索所有文件
            for file_path in uploads_dir.iterdir():
                if file_path.is_file():
                    # 检查是否匹配实际文件名
                    if file_path.name == filename:
                        return file_path
                    
                    # 检查是否匹配显示名称（去除时间戳前缀）
                    current_filename = file_path.name
                    if '_' in current_filename and len(current_filename.split('_')[0]) == 15:
                        parts = current_filename.split('_', 2)
                        if len(parts) >= 3 and parts[2] == filename:
                            return file_path
            
        except Exception as e:
            self.logger.error(f"查找用户文件失败 {session_id}: {e}")
        
        return None
    
    def get_all_user_files(self, session_id: str) -> List[Dict[str, Any]]:
        """
        获取用户所有已上传文件的列表
        
        Args:
            session_id: 用户会话ID
        
        Returns:
            文件信息列表
        """
        workspace = self.get_user_workspace(session_id)
        if not workspace:
            return []
        
        uploads_dir = workspace / "uploads"
        if not uploads_dir.exists():
            return []
        
        all_files = []
        
        try:
            for file_path in uploads_dir.iterdir():
                if file_path.is_file():
                    stat_info = file_path.stat()
                    
                    filename = file_path.name
                    display_name = filename
                    file_type = "其他"
                    
                    # 解析文件名中的时间戳（如果存在）
                    if '_' in filename and len(filename.split('_')[0]) == 15:
                        parts = filename.split('_', 2)
                        if len(parts) >= 3:
                            display_name = parts[2]
                    
                    # 确定文件类型
                    ext = file_path.suffix.lower()
                    if ext in ['.xlsx', '.xls', '.xlsm', '.xlsb']:
                        file_type = "Excel"
                    elif ext in ['.csv']:
                        file_type = "CSV"
                    elif ext in ['.txt']:
                        file_type = "文本"
                    elif ext in ['.pdf']:
                        file_type = "PDF"
                    elif ext in ['.doc', '.docx']:
                        file_type = "Word"
                    
                    all_files.append({
                        'filename': filename,
                        'display_name': display_name,
                        'path': str(file_path),
                        'type': file_type,
                        'size': stat_info.st_size,
                        'modified_time': datetime.fromtimestamp(stat_info.st_mtime),
                        'size_mb': round(stat_info.st_size / (1024 * 1024), 2)
                    })
            
            # 按修改时间排序，最新的在前
            all_files.sort(key=lambda x: x['modified_time'], reverse=True)
            
        except Exception as e:
            self.logger.error(f"获取用户文件列表失败 {session_id}: {e}")
        
        return all_files
    
    def cleanup_user_session(self, session_id: str) -> bool:
        """
        清理用户会话数据
        
        Args:
            session_id: 用户会话ID
        
        Returns:
            是否成功清理
        """
        try:
            user_dir = self.base_upload_dir / session_id
            if user_dir.exists():
                shutil.rmtree(user_dir)
                self.logger.info(f"已清理用户会话: {session_id}")
            
            # 从会话记录中移除
            self._remove_session_info(session_id)
            return True
            
        except Exception as e:
            self.logger.error(f"清理用户会话失败 {session_id}: {e}")
            return False
    
    def cleanup_expired_sessions(self) -> List[str]:
        """
        清理过期的用户会话
        
        Returns:
            已清理的会话ID列表
        """
        cleaned_sessions = []
        sessions = self._load_sessions()
        current_time = datetime.now()
        
        for session_id, session_info in sessions.items():
            try:
                last_access = datetime.fromisoformat(session_info['last_access'])
                if current_time - last_access > self.session_timeout:
                    if self.cleanup_user_session(session_id):
                        cleaned_sessions.append(session_id)
                        
            except Exception as e:
                self.logger.error(f"检查会话过期时出错 {session_id}: {e}")
        
        return cleaned_sessions
    
    def get_session_stats(self) -> Dict[str, Any]:
        """
        获取会话统计信息
        
        Returns:
            会话统计信息
        """
        sessions = self._load_sessions()
        total_sessions = len(sessions)
        
        # 计算活跃会话数（24小时内活跃）
        active_sessions = 0
        current_time = datetime.now()
        
        for session_info in sessions.values():
            try:
                last_access = datetime.fromisoformat(session_info['last_access'])
                if current_time - last_access <= timedelta(hours=24):
                    active_sessions += 1
            except:
                pass
        
        # 计算总文件数
        total_files = sum(session_info.get('file_count', 0) for session_info in sessions.values())
        
        # 计算磁盘使用量
        total_size = self._calculate_directory_size(self.base_upload_dir)
        
        return {
            'total_sessions': total_sessions,
            'active_sessions': active_sessions,
            'total_files': total_files,
            'disk_usage_mb': round(total_size / (1024 * 1024), 2),
            'cleanup_interval_minutes': self.cleanup_interval // 60,
            'session_timeout_hours': self.session_timeout.total_seconds() / 3600
        }
    
    def _sanitize_filename(self, filename: str) -> str:
        """清理文件名，移除危险字符"""
        import re
        # 移除危险字符
        safe_name = re.sub(r'[^\w\s\-\._]', '', filename)
        # 限制长度
        if len(safe_name) > 100:
            name, ext = os.path.splitext(safe_name)
            safe_name = name[:95] + ext
        return safe_name
    
    def _load_sessions(self) -> Dict[str, Any]:
        """加载会话信息"""
        with self.sessions_lock:
            if self.sessions_file.exists():
                try:
                    with open(self.sessions_file, 'r', encoding='utf-8') as f:
                        return json.load(f)
                except Exception as e:
                    self.logger.error(f"加载会话信息失败: {e}")
            return {}
    
    def _save_sessions(self, sessions: Dict[str, Any]):
        """保存会话信息"""
        with self.sessions_lock:
            try:
                with open(self.sessions_file, 'w', encoding='utf-8') as f:
                    json.dump(sessions, f, ensure_ascii=False, indent=2)
            except Exception as e:
                self.logger.error(f"保存会话信息失败: {e}")
    
    def _update_session_info(self, session_id: str, info: Dict[str, Any]):
        """更新会话信息"""
        sessions = self._load_sessions()
        sessions[session_id] = info
        self._save_sessions(sessions)
    
    def _update_session_access(self, session_id: str):
        """更新会话最后访问时间"""
        sessions = self._load_sessions()
        if session_id in sessions:
            sessions[session_id]['last_access'] = datetime.now().isoformat()
            self._save_sessions(sessions)
    
    def _increment_file_count(self, session_id: str):
        """增加会话文件计数"""
        sessions = self._load_sessions()
        if session_id in sessions:
            sessions[session_id]['file_count'] = sessions[session_id].get('file_count', 0) + 1
            self._save_sessions(sessions)
    
    def _remove_session_info(self, session_id: str):
        """从会话记录中移除"""
        sessions = self._load_sessions()
        if session_id in sessions:
            del sessions[session_id]
            self._save_sessions(sessions)
    
    def _calculate_directory_size(self, directory: Path) -> int:
        """计算目录大小（字节）"""
        total_size = 0
        try:
            for file_path in directory.rglob('*'):
                if file_path.is_file():
                    total_size += file_path.stat().st_size
        except Exception as e:
            self.logger.error(f"计算目录大小失败: {e}")
        return total_size
    
    def _cleanup_task(self):
        """定期清理任务"""
        while True:
            try:
                time.sleep(self.cleanup_interval)
                cleaned = self.cleanup_expired_sessions()
                if cleaned:
                    self.logger.info(f"定期清理完成，清理了 {len(cleaned)} 个过期会话")
                    
            except Exception as e:
                self.logger.error(f"定期清理任务出错: {e}")
    
    def _start_cleanup_task(self):
        """启动后台清理任务"""
        cleanup_thread = threading.Thread(target=self._cleanup_task, daemon=True)
        cleanup_thread.start()
        self.logger.info("后台清理任务已启动")
    
    def cleanup_on_exit(self):
        """程序退出时的清理"""
        self.logger.info("程序退出，执行最终清理...")


class UserConfigManager:
    """用户配置管理器"""
    
    def __init__(self, session_manager: UserSessionManager):
        self.session_manager = session_manager
        self.logger = logging.getLogger(__name__)
    
    def save_user_config(self, session_id: str, config: Dict[str, Any]) -> bool:
        """
        保存用户配置到本地
        
        Args:
            session_id: 用户会话ID
            config: 配置信息
        
        Returns:
            是否保存成功
        """
        try:
            workspace = self.session_manager.get_user_workspace(session_id)
            if not workspace:
                workspace = self.session_manager.create_user_workspace(session_id)
            
            config_file = workspace / "user_config.json"
            
            # 添加时间戳
            config['last_updated'] = datetime.now().isoformat()
            
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            
            self.logger.info(f"用户配置已保存: {session_id}")
            return True
            
        except Exception as e:
            self.logger.error(f"保存用户配置失败 {session_id}: {e}")
            return False
    
    def load_user_config(self, session_id: str) -> Optional[Dict[str, Any]]:
        """
        加载用户配置
        
        Args:
            session_id: 用户会话ID
        
        Returns:
            用户配置信息，如果不存在则返回None
        """
        try:
            workspace = self.session_manager.get_user_workspace(session_id)
            if not workspace:
                return None
            
            config_file = workspace / "user_config.json"
            if config_file.exists():
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                return config
            
            return None
            
        except Exception as e:
            self.logger.error(f"加载用户配置失败 {session_id}: {e}")
            return None
    
    def get_config_for_browser_cache(self, config: Dict[str, Any]) -> Dict[str, Any]:
        """
        获取适合浏览器缓存的配置（移除敏感信息）
        
        Args:
            config: 完整配置信息
        
        Returns:
            适合缓存的配置信息
        """
        safe_config = config.copy()
        
        # 移除敏感信息
        sensitive_keys = ['api_key', 'password', 'secret', 'token']
        for key in sensitive_keys:
            if key in safe_config:
                # 只保留前几位和后几位，中间用*号替代
                value = str(safe_config[key])
                if len(value) > 8:
                    safe_config[key] = value[:4] + '*' * (len(value) - 8) + value[-4:]
                else:
                    safe_config[key] = '*' * len(value)
        
        # 添加缓存标识
        safe_config['cached_at'] = datetime.now().isoformat()
        safe_config['cache_type'] = 'browser_safe'
        
        return safe_config
    
    def save_browser_cache_config(self, session_id: str, config: Dict[str, Any]) -> bool:
        """
        保存浏览器缓存配置到本地文件（用于持久化）
        
        Args:
            session_id: 用户会话ID
            config: 适合浏览器缓存的配置
        
        Returns:
            是否保存成功
        """
        try:
            workspace = self.session_manager.get_user_workspace(session_id)
            if not workspace:
                workspace = self.session_manager.create_user_workspace(session_id)
            
            cache_file = workspace / "browser_cache.json"
            safe_config = self.get_config_for_browser_cache(config)
            
            with open(cache_file, 'w', encoding='utf-8') as f:
                json.dump(safe_config, f, ensure_ascii=False, indent=2)
            
            self.logger.info(f"浏览器缓存配置已保存: {session_id}")
            return True
            
        except Exception as e:
            self.logger.error(f"保存浏览器缓存配置失败 {session_id}: {e}")
            return False
    
    def load_browser_cache_config(self, session_id: str) -> Optional[Dict[str, Any]]:
        """
        加载浏览器缓存配置
        
        Args:
            session_id: 用户会话ID
        
        Returns:
            浏览器缓存配置，如果不存在则返回None
        """
        try:
            workspace = self.session_manager.get_user_workspace(session_id)
            if not workspace:
                return None
            
            cache_file = workspace / "browser_cache.json"
            if cache_file.exists():
                with open(cache_file, 'r', encoding='utf-8') as f:
                    cache_config = json.load(f)
                return cache_config
            
            return None
            
        except Exception as e:
            self.logger.error(f"加载浏览器缓存配置失败 {session_id}: {e}")
            return None 