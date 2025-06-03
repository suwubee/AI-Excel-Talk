"""
AI Excel 智能分析工具 - 多用户版配置文件
包含用户管理、隐私保护和多租户配置
"""

# OpenAI API 配置
OPENAI_CONFIG = {
    "api_key": "",  # 用户自己配置
    "base_url": "https://api.openai.com/v1",
    "default_model": "gpt-3.5-turbo",
    "temperature": 0.7,
    "max_tokens": 3000,
}

# 多用户管理配置
MULTIUSER_CONFIG = {
    "enable_multiuser": True,  # 启用多用户支持
    "base_upload_dir": "user_uploads",  # 用户文件存储基础目录
    "session_timeout_hours": 24,  # 会话超时时间（小时）
    "cleanup_interval_minutes": 60,  # 清理任务执行间隔（分钟）
    "max_users_concurrent": 100,  # 最大并发用户数
}

# 用户隐私和安全配置
PRIVACY_CONFIG = {
    "auto_cleanup_enabled": True,  # 启用自动清理
    "cleanup_expired_sessions": True,  # 清理过期会话
    "file_isolation": True,  # 文件隔离
    "config_encryption": False,  # 配置加密（可选功能）
    "anonymous_sessions": True,  # 匿名会话模式
    "log_user_activities": False,  # 是否记录用户活动（隐私考虑）
}

# 文件管理配置
FILE_CONFIG = {
    "max_upload_size_mb": 200,  # 单文件最大上传大小（MB）
    "max_files_per_user": 50,  # 每用户最大文件数
    "allowed_extensions": ["xlsx", "xls"],  # 允许的文件扩展名
    "quarantine_suspicious_files": True,  # 隔离可疑文件
    "scan_uploaded_files": False,  # 文件扫描（需要额外依赖）
}

# 存储限制配置
STORAGE_CONFIG = {
    "max_storage_per_user_mb": 1024,  # 每用户最大存储空间（MB）
    "max_total_storage_gb": 50,  # 系统总存储限制（GB）
    "storage_warning_threshold": 0.8,  # 存储警告阈值（80%）
    "auto_cleanup_when_full": True,  # 存储满时自动清理
}

# 浏览器缓存配置
BROWSER_CACHE_CONFIG = {
    "enable_browser_cache": True,  # 启用浏览器缓存
    "cache_non_sensitive_config": True,  # 缓存非敏感配置
    "cache_duration_days": 7,  # 缓存时长（天）
    "sensitive_keys": [  # 敏感信息键名
        "api_key", "password", "secret", "token", 
        "private_key", "access_token", "refresh_token"
    ],
}

# 会话管理配置
SESSION_CONFIG = {
    "session_id_length": 32,  # 会话ID长度
    "session_renewal_threshold_hours": 12,  # 会话续期阈值（小时）
    "max_inactive_time_minutes": 180,  # 最大非活跃时间（分钟）
    "enable_session_persistence": True,  # 启用会话持久化
    "session_cleanup_on_startup": True,  # 启动时清理过期会话
}

# AI分析配置
AI_CONFIG = {
    "enable_deep_analysis": True,
    "enable_code_generation": True,
    "enable_business_insights": True,
    "max_chat_history_per_user": 50,  # 每用户最大聊天记录数
    "auto_save_analysis": True,
    "analysis_timeout_seconds": 120,  # AI分析超时时间
}

# 代码执行安全配置
CODE_EXECUTION_CONFIG = {
    "enable_code_execution": True,
    "execution_timeout_seconds": 30,  # 代码执行超时
    "max_output_length": 10000,  # 最大输出长度
    "restricted_imports": [  # 禁用的导入模块
        "os", "sys", "subprocess", "shutil", "glob",
        "socket", "urllib", "requests", "ftplib",
        "pickle", "marshal", "shelve", "dbm"
    ],
    "allowed_imports": [  # 允许的导入模块
        "pandas", "numpy", "matplotlib", "plotly", 
        "seaborn", "scipy", "sklearn", "datetime",
        "json", "math", "statistics", "re"
    ],
    "sandbox_mode": True,  # 沙箱模式
}

# 日志和监控配置
LOGGING_CONFIG = {
    "enable_logging": True,
    "log_level": "INFO",  # DEBUG, INFO, WARNING, ERROR
    "log_file": "logs/multiuser_app.log",
    "max_log_size_mb": 100,
    "backup_count": 5,
    "log_rotation": True,
    "log_user_sessions": True,  # 记录用户会话
    "log_file_operations": True,  # 记录文件操作
    "log_cleanup_operations": True,  # 记录清理操作
}

# 性能配置
PERFORMANCE_CONFIG = {
    "enable_caching": True,
    "cache_analysis_results": True,
    "max_cache_entries": 1000,
    "cache_ttl_seconds": 3600,
    "async_file_operations": False,  # 异步文件操作（实验性）
    "optimize_memory_usage": True,
}

# 通知配置
NOTIFICATION_CONFIG = {
    "enable_notifications": False,  # 启用通知（未来功能）
    "storage_warning_notifications": True,
    "cleanup_notifications": True,
    "error_notifications": True,
    "notification_channels": ["log"],  # log, email, webhook
}

# 备份配置
BACKUP_CONFIG = {
    "enable_backups": False,  # 启用备份（可选功能）
    "backup_interval_hours": 24,
    "backup_retention_days": 7,
    "backup_location": "backups/",
    "compress_backups": True,
}

# 管理员配置
ADMIN_CONFIG = {
    "enable_admin_panel": False,  # 启用管理面板（未来功能）
    "admin_password": "",  # 管理员密码
    "admin_endpoints": ["/admin", "/stats", "/cleanup"],
    "require_admin_auth": True,
}

# 开发和调试配置
DEBUG_CONFIG = {
    "debug_mode": False,
    "verbose_logging": False,
    "show_session_ids": True,  # 在界面显示会话ID
    "show_storage_stats": True,  # 显示存储统计
    "enable_profiling": False,  # 性能分析
}

# 功能开关
FEATURES = {
    "multiuser_support": True,
    "file_isolation": True,
    "auto_cleanup": True,
    "config_persistence": True,
    "browser_caching": True,
    "session_management": True,
    "code_execution": True,
    "ai_analysis": True,
    "data_export": True,
    "privacy_protection": True,
}

def get_multiuser_config():
    """
    获取多用户版完整配置
    """
    return {
        "openai": OPENAI_CONFIG,
        "multiuser": MULTIUSER_CONFIG,
        "privacy": PRIVACY_CONFIG,
        "file": FILE_CONFIG,
        "storage": STORAGE_CONFIG,
        "browser_cache": BROWSER_CACHE_CONFIG,
        "session": SESSION_CONFIG,
        "ai": AI_CONFIG,
        "code_execution": CODE_EXECUTION_CONFIG,
        "logging": LOGGING_CONFIG,
        "performance": PERFORMANCE_CONFIG,
        "notification": NOTIFICATION_CONFIG,
        "backup": BACKUP_CONFIG,
        "admin": ADMIN_CONFIG,
        "debug": DEBUG_CONFIG,
        "features": FEATURES,
    }

def validate_config():
    """
    验证配置有效性
    """
    config = get_multiuser_config()
    issues = []
    
    # 检查存储配置
    if config['storage']['max_storage_per_user_mb'] > config['storage']['max_total_storage_gb'] * 1024:
        issues.append("单用户存储限制超过系统总限制")
    
    # 检查会话配置
    if config['session']['session_renewal_threshold_hours'] > config['multiuser']['session_timeout_hours']:
        issues.append("会话续期阈值超过会话超时时间")
    
    # 检查清理配置
    if config['multiuser']['cleanup_interval_minutes'] > config['multiuser']['session_timeout_hours'] * 60:
        issues.append("清理间隔时间过长，可能导致过期会话堆积")
    
    return issues

def get_recommended_config_for_deployment():
    """
    获取推荐的部署配置
    """
    return {
        "multiuser": {
            "session_timeout_hours": 24,
            "cleanup_interval_minutes": 60,
            "max_users_concurrent": 50,
        },
        "storage": {
            "max_storage_per_user_mb": 512,
            "max_total_storage_gb": 20,
        },
        "privacy": {
            "auto_cleanup_enabled": True,
            "cleanup_expired_sessions": True,
            "anonymous_sessions": True,
        },
        "security": {
            "sandbox_mode": True,
            "file_isolation": True,
        }
    } 