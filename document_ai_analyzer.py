#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
文档AI智能分析器 - 专门用于docx、pdf文档的AI分析
类似于Excel分析中的EnhancedAIAnalyzer，提供文档的深度AI分析功能
"""

import openai
import json
from typing import Dict, List, Any, Optional
from document_utils import AdvancedDocumentProcessor

class EnhancedDocumentAIAnalyzer:
    """增强型文档AI分析器 - 提供深度文档分析和AI对话功能"""
    
    def __init__(self, api_key: str, base_url: str = None, model: str = "gpt-4o-mini"):
        """
        初始化AI分析器
        
        Args:
            api_key: OpenAI API密钥
            base_url: API基础URL（可选）
            model: 使用的模型
        """
        self.api_key = api_key
        self.base_url = base_url or "https://api.openai.com/v1"
        self.model = model
        
        # 初始化OpenAI客户端
        self.client = openai.OpenAI(
            api_key=api_key,
            base_url=base_url
        )
        
        # 文档处理器
        self.document_processor = AdvancedDocumentProcessor()
    
    def analyze_document_structure(self, document_analysis: Dict[str, Any]) -> str:
        """
        使用AI分析文档结构和内容
        
        Args:
            document_analysis: 文档分析结果
            
        Returns:
            AI分析报告
        """
        try:
            # 构建AI分析提示词
            prompt = self._build_document_analysis_prompt(document_analysis)
            
            # 调用AI进行分析
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {
                        "role": "system",
                        "content": """你是一位专业的文档分析专家，擅长理解和分析各种类型的文档。
请根据提供的文档结构信息，进行深度业务分析，包括：

1. 文档类型和用途识别
2. 内容主题和业务场景分析
3. 文档结构特点分析
4. 关键信息提取建议
5. 可能的分析方向和价值发现

请用专业但易懂的语言回答，提供实用的洞察和建议。"""
                    },
                    {
                        "role": "user",
                        "content": prompt
                    }
                ],
                temperature=0.7,
                max_tokens=2000
            )
            
            return response.choices[0].message.content
            
        except Exception as e:
            return f"AI分析失败: {str(e)}"
    
    def _build_document_analysis_prompt(self, document_analysis: Dict[str, Any]) -> str:
        """构建文档分析的AI提示词"""
        file_info = document_analysis.get('file_info', {})
        structure = document_analysis.get('structure_analysis', {})
        preview_data = document_analysis.get('preview_data', {})
        ai_data = document_analysis.get('ai_analysis_data', {})
        
        prompt_parts = []
        
        # 基本信息
        prompt_parts.append(f"# 文档分析请求")
        prompt_parts.append(f"")
        prompt_parts.append(f"## 文档基本信息")
        prompt_parts.append(f"- 文件名: {file_info.get('name', 'Unknown')}")
        prompt_parts.append(f"- 文档类型: {file_info.get('type', 'Unknown').upper()}")
        prompt_parts.append(f"- 文件大小: {file_info.get('size_mb', 0)} MB")
        
        # 文档摘要
        doc_summary = ai_data.get('document_summary', {})
        if doc_summary:
            prompt_parts.append(f"- 估算页数: {doc_summary.get('estimated_pages', 'Unknown')}")
            prompt_parts.append(f"- 字数统计: {doc_summary.get('word_count', 0)} 词")
        
        # 结构分析
        prompt_parts.append(f"")
        prompt_parts.append(f"## 文档结构特征")
        
        if file_info.get('type') == 'docx':
            prompt_parts.append(f"- 段落数: {structure.get('total_paragraphs', 0)}")
            prompt_parts.append(f"- 表格数: {structure.get('tables_count', 0)}")
            prompt_parts.append(f"- 图片数: {structure.get('images_count', 0)}")
        elif file_info.get('type') == 'pdf':
            prompt_parts.append(f"- 页数: {structure.get('total_pages', 0)}")
            prompt_parts.append(f"- 图片数: {structure.get('images_count', 0)}")
        
        # 标题层级结构
        headings = structure.get('headings', {})
        if headings:
            prompt_parts.append(f"")
            prompt_parts.append(f"## 标题层级结构")
            for level in sorted(headings.keys()):
                heading_list = headings[level]
                prompt_parts.append(f"### {level}级标题 (共{len(heading_list)}个)")
                for heading in heading_list[:3]:  # 显示前3个作为示例
                    text = heading.get('text', str(heading))[:100]
                    prompt_parts.append(f"- {text}")
                if len(heading_list) > 3:
                    prompt_parts.append(f"... 还有{len(heading_list) - 3}个")
        
        # 字体信息
        fonts = structure.get('fonts_used', [])
        if fonts:
            prompt_parts.append(f"")
            prompt_parts.append(f"## 字体使用情况")
            prompt_parts.append(f"- 使用了{len(fonts)}种字体")
            prompt_parts.append(f"- 主要字体: {', '.join(fonts[:5])}")
        
        # 内容预览（限制长度）
        if preview_data.get('status') == 'success':
            content = preview_data.get('content', '')
            if content:
                preview_content = content[:2000]  # 限制在2000字符
                prompt_parts.append(f"")
                prompt_parts.append(f"## 文档内容预览（前2000字符）")
                prompt_parts.append(f"```")
                prompt_parts.append(preview_content)
                prompt_parts.append(f"```")
        
        return "\n".join(prompt_parts)
    
    def chat_with_document(self, message: str, document_analysis: Dict[str, Any], context: str = "") -> str:
        """
        与文档进行AI对话
        
        Args:
            message: 用户问题
            document_analysis: 文档分析结果
            context: 对话上下文
            
        Returns:
            AI回复
        """
        try:
            # 构建对话上下文
            system_prompt = self._build_chat_system_prompt(document_analysis)
            
            messages = [
                {
                    "role": "system",
                    "content": system_prompt
                }
            ]
            
            # 添加上下文（如果有）
            if context:
                messages.append({
                    "role": "assistant", 
                    "content": f"基于之前的分析：\n{context[:1000]}"  # 限制上下文长度
                })
            
            # 添加用户问题
            messages.append({
                "role": "user",
                "content": message
            })
            
            response = self.client.chat.completions.create(
                model=self.model,
                messages=messages,
                temperature=0.7,
                max_tokens=1500
            )
            
            return response.choices[0].message.content
            
        except Exception as e:
            return f"AI对话失败: {str(e)}"
    
    def _build_chat_system_prompt(self, document_analysis: Dict[str, Any]) -> str:
        """构建对话的系统提示词"""
        file_info = document_analysis.get('file_info', {})
        structure = document_analysis.get('structure_analysis', {})
        
        system_prompt = f"""你是一位专业的文档分析助手，正在帮助用户分析一个{file_info.get('type', 'Unknown').upper()}文档。

文档基本信息：
- 文件名: {file_info.get('name', 'Unknown')}
- 文档类型: {file_info.get('type', 'Unknown').upper()}
- 文件大小: {file_info.get('size_mb', 0)} MB

请根据用户的问题，结合文档的结构和内容信息，提供专业、准确、有用的回答。

回答时请：
1. 直接回答用户的问题
2. 提供具体的分析和建议
3. 必要时给出操作步骤
4. 保持专业但易懂的语言风格"""

        return system_prompt
    
    def generate_document_code_solution(self, task_description: str, document_analysis: Dict[str, Any], filename: str) -> str:
        """
        生成文档处理的代码解决方案
        
        Args:
            task_description: 任务描述
            document_analysis: 文档分析结果
            filename: 文档文件名
            
        Returns:
            生成的Python代码
        """
        try:
            # 构建代码生成提示词
            prompt = self._build_code_generation_prompt(task_description, document_analysis, filename)
            
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {
                        "role": "system",
                        "content": """你是一位Python编程专家，擅长文档处理和数据分析。
请根据用户的需求，生成完整、可执行的Python代码。

代码要求：
1. 使用document_analyzer和document_utils模块
2. 包含完整的导入语句和错误处理
3. 提供详细的中文注释
4. 代码结构清晰，易于理解和修改
5. 包含结果输出和保存功能

专注于实用性和可执行性。"""
                    },
                    {
                        "role": "user",
                        "content": prompt
                    }
                ],
                temperature=0.3,
                max_tokens=2000
            )
            
            return response.choices[0].message.content
            
        except Exception as e:
            return f"# 代码生成失败: {str(e)}\n\n# 请检查API配置或网络连接"
    
    def _build_code_generation_prompt(self, task_description: str, document_analysis: Dict[str, Any], filename: str) -> str:
        """构建代码生成的提示词"""
        file_info = document_analysis.get('file_info', {})
        structure = document_analysis.get('structure_analysis', {})
        
        prompt_parts = []
        
        prompt_parts.append(f"# 代码生成请求")
        prompt_parts.append(f"")
        prompt_parts.append(f"## 任务描述")
        prompt_parts.append(f"{task_description}")
        prompt_parts.append(f"")
        prompt_parts.append(f"## 目标文档信息")
        prompt_parts.append(f"- 文件名: {filename}")
        prompt_parts.append(f"- 文档类型: {file_info.get('type', 'Unknown').upper()}")
        prompt_parts.append(f"- 文件大小: {file_info.get('size_mb', 0)} MB")
        
        # 提供文档结构信息
        if file_info.get('type') == 'docx':
            prompt_parts.append(f"- 段落数: {structure.get('total_paragraphs', 0)}")
            prompt_parts.append(f"- 表格数: {structure.get('tables_count', 0)}")
        elif file_info.get('type') == 'pdf':
            prompt_parts.append(f"- 页数: {structure.get('total_pages', 0)}")
        
        # 标题信息
        headings = structure.get('headings', {})
        if headings:
            prompt_parts.append(f"")
            prompt_parts.append(f"## 文档标题结构")
            for level in sorted(headings.keys()):
                heading_list = headings[level]
                prompt_parts.append(f"- {level}级标题: {len(heading_list)}个")
        
        # 可用的工具类
        prompt_parts.append(f"")
        prompt_parts.append(f"## 可用工具")
        prompt_parts.append(f"- DocumentAnalyzer: 文档分析和搜索")
        prompt_parts.append(f"- AdvancedDocumentProcessor: 高级文档处理")
        prompt_parts.append(f"- DocumentSearchEngine: 搜索引擎")
        
        prompt_parts.append(f"")
        prompt_parts.append(f"请生成完整的Python代码来完成上述任务。")
        
        return "\n".join(prompt_parts)
    
    def suggest_analysis_tasks(self, document_analysis: Dict[str, Any]) -> List[Dict[str, str]]:
        """
        基于文档特征建议分析任务
        
        Args:
            document_analysis: 文档分析结果
            
        Returns:
            建议任务列表
        """
        file_info = document_analysis.get('file_info', {})
        structure = document_analysis.get('structure_analysis', {})
        
        suggestions = []
        
        # 基本分析任务
        suggestions.append({
            "title": "🔍 关键信息提取",
            "description": "搜索和提取文档中的关键信息，如日期、金额、人名等"
        })
        
        suggestions.append({
            "title": "📋 文档结构分析",
            "description": "分析文档的组织结构，识别章节和层级关系"
        })
        
        # 基于文档类型的建议
        if file_info.get('type') == 'pdf':
            suggestions.append({
                "title": "📄 PDF内容提取",
                "description": "按页面提取PDF内容，分析每页的重点信息"
            })
        
        # 基于标题数量的建议
        headings = structure.get('headings', {})
        if headings and sum(len(h) for h in headings.values()) > 5:
            suggestions.append({
                "title": "🏷️ 标题大纲生成",
                "description": "基于标题层级生成文档大纲和目录结构"
            })
        
        # 基于文档大小的建议
        if file_info.get('size_mb', 0) > 1:
            suggestions.append({
                "title": "📊 内容统计分析",
                "description": "统计文档的字数、段落数、关键词频率等"
            })
        
        suggestions.append({
            "title": "🔗 关联内容发现",
            "description": "查找文档中的相关内容和交叉引用"
        })
        
        return suggestions

def main():
    """测试函数"""
    # 这里可以添加测试代码
    print("DocumentAIAnalyzer 初始化完成")

if __name__ == "__main__":
    main() 