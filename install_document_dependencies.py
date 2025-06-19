#!/usr/bin/env python3
"""
文档分析依赖安装脚本
自动检测并安装缺失的文档分析依赖
"""

import subprocess
import sys
from typing import List, Dict

def run_pip_command(command: List[str]) -> bool:
    """运行pip命令"""
    try:
        result = subprocess.run(command, check=True, capture_output=True, text=True)
        print(f"✅ 成功: {' '.join(command)}")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ 失败: {' '.join(command)}")
        print(f"错误信息: {e.stderr}")
        return False

def check_package_availability() -> Dict[str, bool]:
    """检查包的可用性"""
    packages = {}
    
    # 检查markitdown
    try:
        from markitdown import MarkItDown
        packages['markitdown'] = True
        print("✅ markitdown 已安装")
    except ImportError:
        packages['markitdown'] = False
        print("❌ markitdown 未安装或缺少依赖")
    
    # 检查python-docx
    try:
        from docx import Document
        packages['python-docx'] = True
        print("✅ python-docx 已安装")
    except ImportError:
        packages['python-docx'] = False
        print("❌ python-docx 未安装")
    
    # 检查pymupdf
    try:
        import fitz
        packages['pymupdf'] = True
        print("✅ pymupdf 已安装")
    except ImportError:
        packages['pymupdf'] = False
        print("❌ pymupdf 未安装")
    
    # 检查PyPDF2
    try:
        import PyPDF2
        packages['PyPDF2'] = True
        print("✅ PyPDF2 已安装")
    except ImportError:
        packages['PyPDF2'] = False
        print("❌ PyPDF2 未安装")
    
    # 检查pdfplumber
    try:
        import pdfplumber
        packages['pdfplumber'] = True
        print("✅ pdfplumber 已安装")
    except ImportError:
        packages['pdfplumber'] = False
        print("❌ pdfplumber 未安装")
    
    return packages

def install_required_packages():
    """安装必需的包"""
    print("🔧 开始安装文档分析依赖...")
    
    # 需要安装的包列表
    required_packages = [
        "markitdown[all]",  # 包含所有格式支持
        "python-docx",
        "pymupdf",
        "PyPDF2", 
        "pdfplumber",
        "docx2txt",
        "Pillow"
    ]
    
    success_count = 0
    total_packages = len(required_packages)
    
    for package in required_packages:
        print(f"\n📦 安装 {package}...")
        if run_pip_command([sys.executable, "-m", "pip", "install", package]):
            success_count += 1
        else:
            print(f"⚠️ {package} 安装失败，但可能不影响基本功能")
    
    print(f"\n📊 安装完成: {success_count}/{total_packages} 成功")
    return success_count == total_packages

def main():
    """主函数"""
    print("🚀 AI Excel 文档分析依赖检查和安装工具")
    print("=" * 50)
    
    # 检查当前状态
    print("\n🔍 检查当前依赖状态...")
    packages = check_package_availability()
    
    # 统计缺失的包
    missing_packages = [pkg for pkg, available in packages.items() if not available]
    
    if not missing_packages:
        print("\n🎉 所有依赖都已正确安装！")
        
        # 测试功能
        print("\n🧪 测试文档分析功能...")
        try:
            from document_analyzer import DocumentAnalyzer
            analyzer = DocumentAnalyzer()
            missing_deps = analyzer.get_missing_dependencies()
            if not missing_deps:
                print("✅ 文档分析器可以正常工作")
            else:
                print(f"⚠️ 仍缺少依赖: {missing_deps}")
        except Exception as e:
            print(f"❌ 测试失败: {e}")
    else:
        print(f"\n❌ 缺少 {len(missing_packages)} 个依赖: {missing_packages}")
        
        # 询问是否安装
        response = input("\n是否自动安装缺失的依赖? (y/n): ").lower().strip()
        if response in ['y', 'yes', '是']:
            if install_required_packages():
                print("\n🎉 所有依赖安装完成！")
                
                # 重新检查
                print("\n🔍 重新检查依赖状态...")
                packages_after = check_package_availability()
                still_missing = [pkg for pkg, available in packages_after.items() if not available]
                
                if not still_missing:
                    print("\n✅ 所有依赖现在都可用了！")
                else:
                    print(f"\n⚠️ 仍有部分依赖安装失败: {still_missing}")
                    print("请手动安装这些依赖或检查网络连接")
            else:
                print("\n❌ 部分依赖安装失败，请检查错误信息")
        else:
            print("\n💡 手动安装命令:")
            print("pip install markitdown[all] python-docx pymupdf PyPDF2 pdfplumber")
    
    print("\n📚 更多信息:")
    print("- markitdown[all]: 统一的文档转换工具，支持PDF、DOCX等")
    print("- python-docx: DOCX文档结构分析")
    print("- pymupdf: 强大的PDF处理库")
    print("- pdfplumber: PDF表格和布局分析")

if __name__ == "__main__":
    main() 