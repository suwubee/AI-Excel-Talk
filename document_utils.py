#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
文档处理工具类 - 支持docx、pdf文档的高级处理
类似于excel_utils.py，为文档分析提供工具支持
"""

import os
import json
import tempfile
from pathlib import Path
from typing import Dict, List, Any, Tuple, Optional
from datetime import datetime

# 文档分析器
from document_analyzer import DocumentAnalyzer

class AdvancedDocumentProcessor:
    """高级文档处理器 - 提供文档加载、分析、处理的完整功能"""
    
    def __init__(self):
        self.analyzer = DocumentAnalyzer()
        self.current_document = None
        self.analysis_result = None
        self.supported_formats = ['.docx', '.pdf']
    
    def load_document(self, file_path: str) -> Dict[str, Any]:
        """
        加载文档并进行初步分析
        
        Args:
            file_path: 文档文件路径
            
        Returns:
            文档分析结果
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件不存在: {file_path}")
        
        file_ext = Path(file_path).suffix.lower()
        if file_ext not in self.supported_formats:
            raise ValueError(f"不支持的文件格式: {file_ext}, 支持的格式: {self.supported_formats}")
        
        try:
            # 使用文档分析器进行全面分析
            self.analysis_result = self.analyzer.analyze_document(file_path)
            self.current_document = file_path
            
            return self.analysis_result
            
        except Exception as e:
            raise Exception(f"文档加载失败: {str(e)}")
    
    def get_document_preview(self, max_chars: int = 10000) -> str:
        """
        获取文档预览内容
        
        Args:
            max_chars: 最大字符数
            
        Returns:
            预览内容（Markdown格式）
        """
        if not self.analysis_result:
            return "请先加载文档"
        
        preview_data = self.analysis_result.get('preview_data', {})
        if preview_data.get('status') == 'error':
            return f"预览生成失败: {preview_data.get('error', '未知错误')}"
        
        content = preview_data.get('content', '')
        if len(content) > max_chars:
            content = content[:max_chars] + "\n\n... (内容已截断)"
        
        return content
    
    def get_structure_summary(self) -> str:
        """
        获取文档结构摘要
        
        Returns:
            结构化的文档摘要（Markdown格式）
        """
        if not self.analysis_result:
            return "请先加载文档"
        
        file_info = self.analysis_result.get('file_info', {})
        structure = self.analysis_result.get('structure_analysis', {})
        
        summary_lines = []
        summary_lines.append(f"# 📄 文档结构分析")
        summary_lines.append(f"**文件名**: {file_info.get('name', 'Unknown')}")
        summary_lines.append(f"**类型**: {file_info.get('type', 'Unknown').upper()}")
        summary_lines.append(f"**大小**: {file_info.get('size_mb', 0)} MB")
        summary_lines.append("")
        
        # 基本统计
        if file_info.get('type') == 'docx':
            summary_lines.append("## 📊 文档统计")
            summary_lines.append(f"- **段落数**: {structure.get('total_paragraphs', 0)}")
            summary_lines.append(f"- **表格数**: {structure.get('tables_count', 0)}")
            summary_lines.append(f"- **图片数**: {structure.get('images_count', 0)}")
            summary_lines.append(f"- **字体种类**: {len(structure.get('fonts_used', []))}")
            
        elif file_info.get('type') == 'pdf':
            summary_lines.append("## 📊 文档统计")
            summary_lines.append(f"- **页数**: {structure.get('total_pages', 0)}")
            summary_lines.append(f"- **图片数**: {structure.get('images_count', 0)}")
            summary_lines.append(f"- **字体种类**: {len(structure.get('fonts_used', []))}")
        
        # 标题结构 - 100个以内全部显示
        headings = structure.get('headings', {})
        if headings:
            summary_lines.append("")
            summary_lines.append("## 🏷️ 标题层级结构")
            for level in sorted(headings.keys()):
                heading_list = headings[level]
                summary_lines.append(f"### {level}级标题 (共{len(heading_list)}个)")
                
                # 如果标题数量超过100个，显示前100个并提示
                if len(heading_list) > 100:
                    display_headings = heading_list[:100]
                    for heading in display_headings:
                        text = heading.get('text', str(heading))[:80]  # 增加显示长度
                        summary_lines.append(f"- {text}")
                    summary_lines.append(f"- ... 还有{len(heading_list) - 100}个{level}级标题（超过100个限制）")
                else:
                    # 100个以内全部显示
                    for heading in heading_list:
                        text = heading.get('text', str(heading))[:80]  # 增加显示长度
                        summary_lines.append(f"- {text}")
        
        # 字体信息 - 100个以内全部显示
        fonts = structure.get('fonts_used', [])
        if fonts:
            summary_lines.append("")
            summary_lines.append("## 🔤 字体使用情况")
            
            # 如果字体数量超过100个，显示前100个并提示
            if len(fonts) > 100:
                display_fonts = fonts[:100]
                for font in display_fonts:
                    summary_lines.append(f"- {font}")
                summary_lines.append(f"- ... 还有{len(fonts) - 100}种字体（超过100个限制）")
            else:
                # 100个以内全部显示
                for font in fonts:
                    summary_lines.append(f"- {font}")
        
        # 图片分析信息
        images_info = structure.get('images_info', [])
        if images_info:
            summary_lines.append("")
            summary_lines.append("## 🖼️ 图片分析")
            summary_lines.append(f"- **图片总数**: {len(images_info)}")
            for i, img_info in enumerate(images_info[:10], 1):  # 显示前10张图片的详细信息
                summary_lines.append(f"- **图片{i}**: {img_info.get('width', 0)}x{img_info.get('height', 0)}px")
                if 'page' in img_info:
                    summary_lines.append(f"  - 位置: 第{img_info['page']}页")
                if 'watermark_detected' in img_info:
                    if img_info['watermark_detected']:
                        summary_lines.append(f"  - ⚠️ 检测到可能的水印")
                    else:
                        summary_lines.append(f"  - ✅ 未检测到水印")
        
        # 水印检测总结
        watermark_summary = structure.get('watermark_analysis', {})
        if watermark_summary:
            summary_lines.append("")
            summary_lines.append("## 🔍 水印检测")
            if watermark_summary.get('has_watermark', False):
                summary_lines.append("- ⚠️ **检测到可能的水印**")
                summary_lines.append(f"- 水印类型: {watermark_summary.get('watermark_type', '未知')}")
                summary_lines.append(f"- 检测置信度: {watermark_summary.get('confidence', 0):.2f}")
            else:
                summary_lines.append("- ✅ **未检测到明显水印**")
        
        return "\n".join(summary_lines)
    
    def search_content(self, keyword: str, context_lines: int = 3) -> List[Dict[str, Any]]:
        """
        在文档中搜索关键词
        
        Args:
            keyword: 搜索关键词
            context_lines: 上下文行数
            
        Returns:
            搜索结果列表
        """
        if not self.analysis_result:
            return []
        
        return self.analyzer.search_keyword_context(keyword, context_lines)
    
    def generate_analysis_code(self, task_description: str) -> str:
        """
        生成文档分析代码
        
        Args:
            task_description: 任务描述
            
        Returns:
            生成的Python代码
        """
        if not self.analysis_result:
            return "# 请先加载文档"
        
        return self.analyzer.generate_analysis_code(task_description)
    
    def export_analysis_result(self, output_dir: str = None) -> Tuple[str, str]:
        """
        导出分析结果
        
        Args:
            output_dir: 输出目录
            
        Returns:
            (JSON文件路径, Markdown报告路径)
        """
        if not self.analysis_result:
            raise ValueError("没有分析结果可导出")
        
        if output_dir is None:
            output_dir = tempfile.gettempdir()
        
        output_dir = Path(output_dir)
        output_dir.mkdir(exist_ok=True)
        
        file_name = Path(self.current_document).stem
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # 导出JSON数据
        json_file = output_dir / f"{file_name}_analysis_{timestamp}.json"
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(self.analysis_result, f, ensure_ascii=False, indent=2, default=str)
        
        # 导出Markdown报告
        md_file = output_dir / f"{file_name}_report_{timestamp}.md"
        with open(md_file, 'w', encoding='utf-8') as f:
            f.write(self._generate_markdown_report())
        
        return str(json_file), str(md_file)
    
    def _generate_markdown_report(self) -> str:
        """生成Markdown格式的分析报告"""
        if not self.analysis_result:
            return "# 分析报告\n\n无数据"
        
        file_info = self.analysis_result.get('file_info', {})
        preview_data = self.analysis_result.get('preview_data', {})
        structure = self.analysis_result.get('structure_analysis', {})
        ai_data = self.analysis_result.get('ai_analysis_data', {})
        
        report_lines = []
        
        # 报告头部
        report_lines.extend([
            f"# 📄 文档智能分析报告",
            f"",
            f"**生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            f"**文档名称**: {file_info.get('name', 'Unknown')}",
            f"**文档类型**: {file_info.get('type', 'Unknown').upper()}",
            f"**文件大小**: {file_info.get('size_mb', 0)} MB",
            f"",
            f"---",
            f""
        ])
        
        # 文档概要
        report_lines.extend([
            f"## 📊 文档概要",
            f"",
            f"### 基本信息"
        ])
        
        doc_summary = ai_data.get('document_summary', {})
        if doc_summary:
            report_lines.extend([
                f"- **估算页数**: {doc_summary.get('estimated_pages', 'Unknown')}",
                f"- **字数统计**: {doc_summary.get('word_count', 0)} 词",
                f"- **文档规模**: {'大型文档' if doc_summary.get('estimated_pages', 0) > 20 else '中小型文档'}",
                f""
            ])
        
        # 结构分析
        structure_insights = ai_data.get('structure_insights', {})
        if structure_insights:
            report_lines.extend([
                f"### 结构特征",
                f"- **字体种类**: {structure_insights.get('fonts_count', 0)} 种",
                f"- **主要字体**: {', '.join(structure_insights.get('main_fonts', [])[:3])}",
                f""
            ])
            
            headings_info = structure_insights.get('headings', {})
            if headings_info:
                report_lines.append(f"### 标题层级")
                for level, info in headings_info.items():
                    report_lines.append(f"- **{level}**: {info['count']}个")
                    examples = info.get('examples', [])
                    if examples:
                        report_lines.append(f"  - 示例: {' | '.join(examples)}")
                report_lines.append("")
        
        # 内容预览
        if preview_data.get('status') == 'success':
            content = preview_data.get('content', '')
            if content:
                report_lines.extend([
                    f"## 📝 内容预览",
                    f"",
                    f"### 文档内容（前{min(1000, len(content))}字符）",
                    f"```",
                    content[:1000],
                    f"```",
                    f"",
                    f"---",
                    f""
                ])
        
        # AI分析建议
        suggestions = ai_data.get('analysis_suggestions', [])
        if suggestions:
            report_lines.extend([
                f"## 🎯 AI分析建议",
                f""
            ])
            for i, suggestion in enumerate(suggestions, 1):
                report_lines.append(f"{i}. {suggestion}")
            report_lines.append("")
        
        # 技术信息
        report_lines.extend([
            f"## 🔧 技术信息",
            f"",
            f"### 分析方法",
            f"- **预览生成**: MarkItDown (清除格式，保留结构)",
            f"- **结构分析**: 原生文档格式解析",
            f"- **搜索功能**: 全文索引",
            f"- **代码生成**: 智能模板",
            f"",
            f"### 分析覆盖",
            f"- ✅ 文档基本信息",
            f"- ✅ 标题层级结构", 
            f"- ✅ 字体使用情况",
            f"- ✅ 图片数量统计",
            f"- ✅ 关键词搜索支持",
            f"",
            f"---",
            f"",
            f"*报告由AI Excel智能分析工具 - 文档分析模块生成*"
        ])
        
        return "\n".join(report_lines)
    
    def get_analysis_cache(self) -> Optional[Dict[str, Any]]:
        """获取分析结果缓存"""
        return self.analysis_result
    
    def clear_cache(self):
        """清除缓存"""
        self.analysis_result = None
        self.current_document = None

class DocumentSearchEngine:
    """文档搜索引擎 - 专门处理关键词搜索和上下文提取"""
    
    def __init__(self, document_processor: AdvancedDocumentProcessor):
        self.processor = document_processor
    
    def search_multiple_keywords(self, keywords: List[str], context_lines: int = 3) -> Dict[str, List[Dict[str, Any]]]:
        """
        搜索多个关键词
        
        Args:
            keywords: 关键词列表
            context_lines: 上下文行数
            
        Returns:
            以关键词为key的搜索结果字典
        """
        results = {}
        for keyword in keywords:
            results[keyword] = self.processor.search_content(keyword, context_lines)
        return results
    
    def generate_search_report(self, keywords: List[str]) -> str:
        """
        生成搜索报告
        
        Args:
            keywords: 搜索的关键词列表
            
        Returns:
            Markdown格式的搜索报告
        """
        search_results = self.search_multiple_keywords(keywords)
        
        report_lines = []
        report_lines.append(f"# 🔍 关键词搜索报告")
        report_lines.append(f"")
        report_lines.append(f"**搜索时间**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        report_lines.append(f"**搜索关键词**: {', '.join(keywords)}")
        report_lines.append(f"")
        
        for keyword, results in search_results.items():
            report_lines.append(f"## 🎯 关键词: \"{keyword}\"")
            report_lines.append(f"")
            
            if not results:
                report_lines.append(f"❌ 未找到相关内容")
                report_lines.append(f"")
                continue
            
            report_lines.append(f"✅ 找到 {len(results)} 处匹配")
            report_lines.append(f"")
            
            for i, result in enumerate(results[:5], 1):  # 最多显示5个结果
                report_lines.append(f"### 结果 {i}")
                report_lines.append(f"**位置**: 第 {result['line_number']} 行")
                report_lines.append(f"**匹配内容**: {result['matched_line']}")
                report_lines.append(f"")
                report_lines.append(f"**上下文**:")
                report_lines.append(f"```")
                report_lines.append(result['context'])
                report_lines.append(f"```")
                report_lines.append(f"")
            
            if len(results) > 5:
                report_lines.append(f"... 还有 {len(results) - 5} 个结果")
                report_lines.append(f"")
        
        return "\n".join(report_lines)

def main():
    """测试函数"""
    processor = AdvancedDocumentProcessor()
    
    # 测试文档处理
    test_file = "test_document.docx"  # 替换为实际文件路径
    if os.path.exists(test_file):
        try:
            result = processor.load_document(test_file)
            print("✅ 文档加载成功")
            
            # 获取结构摘要
            summary = processor.get_structure_summary()
            print(summary)
            
            # 测试搜索
            search_results = processor.search_content("合同")
            print(f"\n搜索结果: {len(search_results)} 个")
            
        except Exception as e:
            print(f"❌ 测试失败: {str(e)}")
    else:
        print("❌ 测试文件不存在")

if __name__ == "__main__":
    main() 