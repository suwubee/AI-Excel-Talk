#!/usr/bin/env python3
"""
文档分析依赖安装脚本
自动检测并安装缺失的文档分析依赖
支持服务器环境和开发环境，包括系统级依赖（如ffmpeg）
"""

import subprocess
import sys
import os
import warnings
import shutil
from typing import List, Dict, Tuple

# 抑制一些服务器环境下的警告
warnings.filterwarnings("ignore", category=RuntimeWarning)
os.environ['PYTHONWARNINGS'] = 'ignore::RuntimeWarning'

def check_ffmpeg() -> bool:
    """检查ffmpeg是否已安装"""
    return shutil.which('ffmpeg') is not None

def install_ffmpeg_debian():
    """在Debian/Ubuntu系统上安装ffmpeg"""
    print("\n🎵 检测到需要安装 ffmpeg...")
    
    if not check_ffmpeg():
        print("📦 ffmpeg 未安装，尝试自动安装...")
        
        try:
            # 检查是否有sudo权限
            result = subprocess.run(['sudo', '-n', 'true'], capture_output=True)
            has_sudo = result.returncode == 0
        except:
            has_sudo = False
        
        if has_sudo:
            print("🔧 使用 sudo 安装 ffmpeg...")
            try:
                # 更新包列表
                subprocess.run(['sudo', 'apt', 'update'], check=True, capture_output=True)
                # 安装ffmpeg
                subprocess.run(['sudo', 'apt', 'install', '-y', 'ffmpeg'], check=True, capture_output=True)
                print("✅ ffmpeg 安装成功！")
                return True
            except subprocess.CalledProcessError as e:
                print(f"❌ ffmpeg 安装失败: {e}")
                return False
        else:
            print("⚠️ 需要管理员权限安装 ffmpeg")
            print("💡 请手动执行以下命令:")
            print("   sudo apt update")
            print("   sudo apt install -y ffmpeg")
            return False
    else:
        print("✅ ffmpeg 已安装")
        return True

def check_system_dependencies():
    """检查系统级依赖"""
    print("\n🔍 检查系统依赖...")
    
    deps_status = {}
    
    # 检查ffmpeg
    ffmpeg_available = check_ffmpeg()
    deps_status['ffmpeg'] = ffmpeg_available
    
    if ffmpeg_available:
        print("✅ ffmpeg 可用")
    else:
        print("❌ ffmpeg 缺失（可能影响音频处理功能）")
    
    return deps_status

def run_pip_command(packages: List[str]) -> Tuple[int, int]:
    """批量安装包，返回(成功数, 总数)"""
    success_count = 0
    total_count = len(packages)
    
    for package in packages:
        try:
            print(f"📦 安装 {package}...")
            # 使用静默安装，减少输出噪音
            result = subprocess.run([
                sys.executable, "-m", "pip", "install", package, 
                "--quiet", "--no-warn-script-location"
            ], check=True, capture_output=True, text=True)
            print(f"✅ {package} 安装成功")
            success_count += 1
        except subprocess.CalledProcessError as e:
            print(f"❌ {package} 安装失败")
            if e.stderr and "error" in e.stderr.lower():
                # 只显示真正的错误，过滤警告
                error_lines = [line for line in e.stderr.split('\n') 
                              if 'error' in line.lower() and 'warning' not in line.lower()]
                if error_lines:
                    print(f"   错误: {error_lines[0].strip()}")
        except Exception as e:
            print(f"❌ {package} 安装异常: {str(e)}")
    
    return success_count, total_count

def check_imports() -> Dict[str, bool]:
    """检查关键模块导入状态"""
    modules = {
        'markitdown': 'markitdown',
        'python-docx': 'docx', 
        'pymupdf': 'fitz',
        'PyPDF2': 'PyPDF2',
        'pdfplumber': 'pdfplumber',
        'Pillow': 'PIL'
    }
    
    status = {}
    print("\n🔍 检查Python依赖状态...")
    
    for package_name, import_name in modules.items():
        try:
            # 临时捕获导入时的警告
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                __import__(import_name)
            status[package_name] = True
            print(f"✅ {package_name}")
        except ImportError:
            status[package_name] = False
            print(f"❌ {package_name} 缺失")
        except Exception as e:
            # 处理其他可能的导入异常
            status[package_name] = False
            print(f"⚠️ {package_name} 导入异常: {str(e)}")
    
    return status

def test_document_analyzer():
    """测试文档分析器功能"""
    print("\n🧪 测试文档分析器...")
    try:
        # 抑制测试时的警告
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            from document_analyzer import DocumentAnalyzer
            analyzer = DocumentAnalyzer()
            is_ready, missing = analyzer.is_ready_for_analysis()
        
        if is_ready:
            print("✅ 文档分析器可以正常工作")
            return True
        else:
            print(f"⚠️ 缺少关键依赖: {missing}")
            return False
    except Exception as e:
        print(f"❌ 文档分析器测试失败: {str(e)}")
        return False

def check_server_environment():
    """检查服务器环境"""
    env_info = {
        'python_version': f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}",
        'platform': sys.platform,
        'is_venv': hasattr(sys, 'real_prefix') or (hasattr(sys, 'base_prefix') and sys.base_prefix != sys.prefix),
        'pip_version': None,
        'os_release': None
    }
    
    try:
        result = subprocess.run([sys.executable, "-m", "pip", "--version"], 
                              capture_output=True, text=True)
        if result.returncode == 0:
            env_info['pip_version'] = result.stdout.strip().split()[1]
    except:
        pass
    
    # 检查Linux发行版
    if sys.platform == 'linux':
        try:
            with open('/etc/os-release', 'r') as f:
                for line in f:
                    if line.startswith('PRETTY_NAME='):
                        env_info['os_release'] = line.split('=')[1].strip().strip('"')
                        break
        except:
            pass
    
    return env_info

def main():
    """主函数"""
    print("🚀 AI Excel 文档分析依赖管理工具")
    print("=" * 50)
    
    # 检查环境信息
    env = check_server_environment()
    print(f"🐍 Python {env['python_version']} ({env['platform']})")
    if env['os_release']:
        print(f"🐧 系统: {env['os_release']}")
    if env['is_venv']:
        print("📦 虚拟环境: 已激活")
    if env['pip_version']:
        print(f"📦 pip版本: {env['pip_version']}")
    
    # 检查系统依赖
    sys_deps = check_system_dependencies()
    
    # 如果是Linux且缺少ffmpeg，询问是否安装
    if sys.platform == 'linux' and not sys_deps.get('ffmpeg', False):
        try:
            response = input("\n是否安装 ffmpeg? (推荐，用于音频处理) (y/n): ").strip().lower()
            if response in ['y', 'yes', '是']:
                install_ffmpeg_debian()
        except (EOFError, KeyboardInterrupt):
            print("\n⚠️ 非交互环境，跳过 ffmpeg 安装")
    
    # 检查Python依赖状态
    status = check_imports()
    missing = [pkg for pkg, available in status.items() if not available]
    
    if not missing:
        print(f"\n🎉 所有Python依赖已安装 ({len(status)}/{len(status)})")
        test_document_analyzer()
        print("\n✅ 系统就绪，可以使用文档分析功能！")
        return
    
    print(f"\n❌ 缺少 {len(missing)} 个Python依赖: {', '.join(missing)}")
    
    # 在服务器环境下，提供非交互式选项
    if len(sys.argv) > 1 and sys.argv[1] in ['--auto', '-a']:
        print("🤖 自动安装模式...")
        install_missing = True
    else:
        # 交互式询问
        try:
            response = input("\n是否安装缺失的Python依赖? (y/n): ").strip().lower()
            install_missing = response in ['y', 'yes', '是']
        except (EOFError, KeyboardInterrupt):
            print("\n⚠️ 检测到非交互环境，使用自动安装模式")
            install_missing = True
    
    if not install_missing:
        print("💡 手动安装命令:")
        print("pip install markitdown[all] python-docx pymupdf PyPDF2 pdfplumber Pillow")
        return
    
    # 准备安装包列表（使用最新版本）
    packages_to_install = []
    if not status['markitdown']:
        packages_to_install.append("markitdown[all]>=0.1.0")
    if not status['python-docx']:
        packages_to_install.append("python-docx>=0.8.11")
    if not status['pymupdf']:
        packages_to_install.append("pymupdf>=1.23.0")
    if not status['PyPDF2']:
        packages_to_install.append("PyPDF2>=3.0.0")
    if not status['pdfplumber']:
        packages_to_install.append("pdfplumber>=0.9.0")
    if not status['Pillow']:
        packages_to_install.append("Pillow>=9.0.0")
    
    # 批量安装
    print(f"\n🔧 开始安装 {len(packages_to_install)} 个Python包...")
    success, total = run_pip_command(packages_to_install)
    
    print(f"\n📊 安装结果: {success}/{total} 成功")
    
    if success == total:
        print("🎉 所有依赖安装完成！")
        test_document_analyzer()
    elif success > 0:
        print("⚠️ 部分依赖安装成功，核心功能可用")
        test_document_analyzer()
    else:
        print("❌ 安装失败，请检查网络连接或权限")
        print("💡 可能的解决方案:")
        print("   1. sudo python install_document_dependencies.py")
        print("   2. pip install --user markitdown[all] python-docx pymupdf")
        print("   3. 检查网络连接和防火墙设置")
    
    print("\n📚 功能说明:")
    print("- markitdown[all]: 文档内容提取和转换")
    print("- python-docx: DOCX文档结构分析") 
    print("- pymupdf: PDF处理和图片提取")
    print("- pdfplumber: PDF表格和布局分析")
    print("- Pillow: 图像处理和水印检测")
    
    # 服务器环境下的额外提示
    if env['platform'] == 'linux':
        print("\n🐧 Linux服务器提示:")
        print("- 如遇权限问题，请使用 sudo 或 --user 参数")
        if not sys_deps.get('ffmpeg', False):
            print("- 建议安装 ffmpeg: sudo apt install ffmpeg")
        print("- 建议在虚拟环境中运行以避免依赖冲突")

if __name__ == "__main__":
    main() 