#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AI Excel 智能分析工具 - 多用户版启动脚本

使用方法:
    python run_multiuser.py
    
或者指定端口:
    python run_multiuser.py --port 8502
"""

import subprocess
import sys
import argparse
import os
from pathlib import Path

def check_dependencies():
    """检查依赖包是否已安装"""
    try:
        import streamlit
        import pandas
        import openai
        import plotly
        import numpy
        print("✅ 所有依赖包已正确安装")
        return True
    except ImportError as e:
        print(f"❌ 缺少依赖包: {e}")
        print("📦 请运行以下命令安装依赖:")
        print("   pip install -r requirements.txt")
        return False

def check_files():
    """检查必要文件是否存在"""
    required_files = [
        "app_enhanced_multiuser.py",
        "user_session_manager.py", 
        "excel_utils.py",
        "config_multiuser.py",
        "requirements.txt"
    ]
    
    missing_files = []
    for file in required_files:
        if not Path(file).exists():
            missing_files.append(file)
    
    if missing_files:
        print(f"❌ 缺少必要文件: {', '.join(missing_files)}")
        return False
    
    print("✅ 所有必要文件检查通过")
    return True

def main():
    parser = argparse.ArgumentParser(description="AI Excel 智能分析工具 - 多用户版")
    parser.add_argument("--port", type=int, default=8501, help="指定端口号（默认: 8501）")
    parser.add_argument("--debug", action="store_true", help="启用调试模式")
    parser.add_argument("--host", default="localhost", help="指定主机地址（默认: localhost）")
    
    args = parser.parse_args()
    
    print("🚀 AI Excel 智能分析工具 - 多用户版")
    print("=" * 50)
    
    # 检查依赖
    if not check_dependencies():
        sys.exit(1)
    
    # 检查文件
    if not check_files():
        sys.exit(1)
    
    # 构建启动命令
    cmd = [
        sys.executable, "-m", "streamlit", "run",
        "app_enhanced_multiuser.py",
        "--server.port", str(args.port),
        "--server.address", args.host,
        "--server.headless", "false"
    ]
    
    if args.debug:
        cmd.extend(["--logger.level", "debug"])
    
    print(f"🌐 启动地址: http://{args.host}:{args.port}")
    print(f"🔧 调试模式: {'开启' if args.debug else '关闭'}")
    print("📝 使用说明:")
    print("   1. 在侧边栏配置OpenAI API Key")
    print("   2. 选择模型（推荐 deepseek-v3）")
    print("   3. 上传Excel文件开始分析")
    print("=" * 50)
    print("⚡ 正在启动应用...")
    
    try:
        # 启动Streamlit应用
        subprocess.run(cmd)
    except KeyboardInterrupt:
        print("\n👋 应用已停止")
    except Exception as e:
        print(f"❌ 启动失败: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main() 