#!/usr/bin/env python3
"""
文档分析依赖安装脚本
自动检测并安装缺失的文档分析依赖
"""

import subprocess
import sys
from typing import List, Dict, Tuple

def run_pip_command(packages: List[str]) -> Tuple[int, int]:
    """批量安装包，返回(成功数, 总数)"""
    success_count = 0
    total_count = len(packages)
    
    for package in packages:
        try:
            print(f"📦 安装 {package}...")
            result = subprocess.run([sys.executable, "-m", "pip", "install", package], 
                                  check=True, capture_output=True, text=True)
            print(f"✅ {package} 安装成功")
            success_count += 1
        except subprocess.CalledProcessError as e:
            print(f"❌ {package} 安装失败")
            if e.stderr:
                print(f"   错误: {e.stderr.strip()}")
    
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
    print("\n🔍 检查依赖状态...")
    
    for package_name, import_name in modules.items():
        try:
            __import__(import_name)
            status[package_name] = True
            print(f"✅ {package_name}")
        except ImportError:
            status[package_name] = False
            print(f"❌ {package_name} 缺失")
    
    return status

def test_document_analyzer():
    """测试文档分析器功能"""
    print("\n🧪 测试文档分析器...")
    try:
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
        print(f"❌ 文档分析器测试失败: {e}")
        return False

def main():
    """主函数"""
    print("🚀 AI Excel 文档分析依赖管理工具")
    print("=" * 50)
    
    # 检查当前状态
    status = check_imports()
    missing = [pkg for pkg, available in status.items() if not available]
    
    if not missing:
        print(f"\n🎉 所有依赖已安装 ({len(status)}/{len(status)})")
        test_document_analyzer()
        return
    
    print(f"\n❌ 缺少 {len(missing)} 个依赖: {', '.join(missing)}")
    
    # 询问是否安装
    response = input("\n是否安装缺失的依赖? (y/n): ").strip().lower()
    if response not in ['y', 'yes', '是']:
        print("💡 手动安装: pip install markitdown[all] python-docx pymupdf PyPDF2 pdfplumber Pillow")
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
    print(f"\n🔧 开始安装 {len(packages_to_install)} 个包...")
    success, total = run_pip_command(packages_to_install)
    
    print(f"\n📊 安装结果: {success}/{total} 成功")
    
    if success == total:
        print("🎉 所有依赖安装完成！")
        test_document_analyzer()
    elif success > 0:
        print("⚠️ 部分依赖安装成功，核心功能可用")
        test_document_analyzer()
    else:
        print("❌ 安装失败，请检查网络连接或手动安装")
    
    print("\n📚 功能说明:")
    print("- markitdown[all]: 文档内容提取和转换")
    print("- python-docx: DOCX文档结构分析") 
    print("- pymupdf: PDF处理和图片提取")
    print("- pdfplumber: PDF表格和布局分析")
    print("- Pillow: 图像处理和水印检测")

if __name__ == "__main__":
    main() 