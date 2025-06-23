#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
文档分析功能测试脚本 - 验证增强功能
测试内容：
1. 结构化分析不缩略显示（100个以内全部显示）
2. 图片预览和分析功能
3. 水印检测功能
"""

import os
import sys
import json
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

try:
    from document_analyzer import DocumentAnalyzer
    from document_utils import AdvancedDocumentProcessor
except ImportError as e:
    print(f"❌ 导入模块失败: {e}")
    print("请确保已安装所有依赖: pip install -r requirements.txt")
    sys.exit(1)

def test_enhanced_document_analysis():
    """测试增强的文档分析功能"""
    
    print("🧪 开始测试增强的文档分析功能")
    print("=" * 60)
    
    # 初始化处理器
    processor = AdvancedDocumentProcessor()
    print("✅ 文档处理器初始化成功")
    
    # 测试文件路径（您需要替换为实际的测试文件）
    test_files = [
        # "path/to/your/test.docx",
        # "path/to/your/test.pdf"
    ]
    
    # 如果没有指定测试文件，创建一个示例测试
    if not test_files:
        print("⚠️ 未指定测试文件，将创建模拟测试...")
        test_structure_display_logic()
        return
    
    for file_path in test_files:
        if not os.path.exists(file_path):
            print(f"⚠️ 测试文件不存在: {file_path}")
            continue
            
        print(f"\n📄 正在测试文件: {Path(file_path).name}")
        print("-" * 40)
        
        try:
            # 加载和分析文档
            analysis_result = processor.load_document(file_path)
            print("✅ 文档加载成功")
            
            # 测试1: 结构摘要（不缩略显示）
            print("\n1️⃣ 测试结构摘要（完整显示）")
            structure_summary = processor.get_structure_summary()
            
            # 统计标题和字体数量
            structure_analysis = analysis_result.get('structure_analysis', {})
            headings = structure_analysis.get('headings', {})
            fonts = structure_analysis.get('fonts_used', [])
            
            total_headings = sum(len(headings[level]) for level in headings.keys())
            print(f"   - 标题总数: {total_headings}")
            print(f"   - 字体总数: {len(fonts)}")
            
            if total_headings <= 100:
                print(f"   ✅ 标题数量≤100，应全部显示")
            else:
                print(f"   ⚠️ 标题数量>100，应显示前100个并提示")
                
            if len(fonts) <= 100:
                print(f"   ✅ 字体数量≤100，应全部显示")
            else:
                print(f"   ⚠️ 字体数量>100，应显示前100个并提示")
            
            # 测试2: 图片预览和分析
            print("\n2️⃣ 测试图片预览和分析")
            preview_data = analysis_result.get('preview_data', {})
            has_images = preview_data.get('has_images', False)
            images_preview = preview_data.get('images_preview', [])
            
            print(f"   - 检测到图片: {'是' if has_images else '否'}")
            print(f"   - 图片数量: {len(images_preview)}")
            
            if images_preview:
                for i, img_info in enumerate(images_preview, 1):
                    print(f"   - 图片{i}: {img_info.get('width', 0)}x{img_info.get('height', 0)}px, "
                          f"{img_info.get('size_kb', 0):.1f}KB, "
                          f"水印: {'是' if img_info.get('watermark_detected', False) else '否'}")
            
            # 测试3: 水印检测
            print("\n3️⃣ 测试水印检测")
            watermark_analysis = structure_analysis.get('watermark_analysis', {})
            
            if watermark_analysis:
                has_watermark = watermark_analysis.get('has_watermark', False)
                confidence = watermark_analysis.get('confidence', 0)
                watermark_type = watermark_analysis.get('watermark_type', '未知')
                
                print(f"   - 整体水印检测: {'检测到' if has_watermark else '未检测到'}")
                if has_watermark:
                    print(f"   - 置信度: {confidence:.2f}")
                    print(f"   - 水印类型: {watermark_type}")
                    print(f"   - 含水印图片数: {watermark_analysis.get('watermark_images', 0)}")
            else:
                print("   - 无水印分析数据")
            
            print(f"\n✅ 文件 {Path(file_path).name} 测试完成")
            
        except Exception as e:
            print(f"❌ 测试失败: {str(e)}")
            import traceback
            traceback.print_exc()

def test_structure_display_logic():
    """测试结构显示逻辑（模拟）"""
    print("\n🔍 测试结构显示逻辑")
    print("-" * 40)
    
    # 模拟不同数量的标题和字体
    test_cases = [
        {"headings": 5, "fonts": 8, "description": "小型文档"},
        {"headings": 50, "fonts": 30, "description": "中型文档"},
        {"headings": 100, "fonts": 100, "description": "大型文档（边界值）"},
        {"headings": 150, "fonts": 120, "description": "超大型文档"}
    ]
    
    for case in test_cases:
        print(f"\n📋 {case['description']}")
        print(f"   - 标题数: {case['headings']}")
        print(f"   - 字体数: {case['fonts']}")
        
        # 模拟显示逻辑
        if case['headings'] <= 100:
            print(f"   ✅ 标题: 全部显示 ({case['headings']}个)")
        else:
            print(f"   ⚠️ 标题: 显示前100个，提示还有{case['headings'] - 100}个")
            
        if case['fonts'] <= 100:
            print(f"   ✅ 字体: 全部显示 ({case['fonts']}个)")
        else:
            print(f"   ⚠️ 字体: 显示前100个，提示还有{case['fonts'] - 100}个")

def test_watermark_detection():
    """测试水印检测算法（模拟）"""
    print("\n🔍 测试水印检测算法")
    print("-" * 40)
    
    # 模拟不同类型的图片特征
    test_images = [
        {
            "name": "普通图片",
            "variance": 1000,
            "alpha_variation": 10,
            "border_activity": 5,
            "color_imbalance": 1.2
        },
        {
            "name": "透明水印图片", 
            "variance": 800,
            "alpha_variation": 50,
            "border_activity": 15,
            "color_imbalance": 1.8
        },
        {
            "name": "重复模式水印",
            "variance": 300,
            "alpha_variation": 20,
            "border_activity": 25,
            "color_imbalance": 2.5
        }
    ]
    
    for img in test_images:
        print(f"\n🖼️ {img['name']}")
        confidence = 0.0
        watermark_type = 'none'
        
        # 模拟检测逻辑
        if img['alpha_variation'] > 30:
            confidence += 0.3
            watermark_type = 'transparent'
            
        if img['variance'] < 500:
            confidence += 0.2
            if watermark_type == 'none':
                watermark_type = 'pattern'
                
        if img['border_activity'] > 20:
            confidence += 0.1
            if watermark_type == 'none':
                watermark_type = 'border'
                
        if img['color_imbalance'] > 2:
            confidence += 0.15
            if watermark_type == 'none':
                watermark_type = 'colored'
        
        has_watermark = confidence > 0.3
        confidence = min(1.0, confidence)
        
        print(f"   - 水印检测: {'是' if has_watermark else '否'}")
        print(f"   - 置信度: {confidence:.2f}")
        print(f"   - 类型: {watermark_type}")

def main():
    """主测试函数"""
    print("🚀 文档分析增强功能测试")
    print("=" * 60)
    print("测试功能:")
    print("1. 结构化分析不缩略显示（100个以内全部显示）")
    print("2. 图片预览和分析功能") 
    print("3. 水印检测功能")
    print("=" * 60)
    
    try:
        # 测试文档分析功能
        test_enhanced_document_analysis()
        
        # 测试水印检测算法
        test_watermark_detection()
        
        print("\n" + "=" * 60)
        print("✅ 所有测试完成")
        print("\n💡 使用说明:")
        print("1. 运行主应用: streamlit run app_enhanced_multiuser.py")
        print("2. 选择'📄 文档分析'模式")
        print("3. 上传PDF或DOCX文件")
        print("4. 在'📄 文档预览'标签查看图片预览和水印检测结果")
        print("5. 结构分析结果现在会完整显示100个以内的所有项目")
        
    except Exception as e:
        print(f"❌ 测试过程中出现错误: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main() 