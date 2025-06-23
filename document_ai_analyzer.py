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

## 严格的代码格式要求：
1. **只返回纯Python代码**，不要包含任何markdown格式标记（如```python 或 ```）
2. **不要添加任何解释性文字**，只返回可直接执行的代码
3. **代码必须是完整的**，包含所有必要的导入语句
4. **每行代码前不要有额外的空格或制表符**

## 代码质量要求：
1. 使用document_analyzer和document_utils模块进行文档处理
2. 包含完整的导入语句和错误处理
3. 提供详细的中文注释，解释每个步骤的作用
4. 代码结构清晰，逻辑分明，易于理解和修改
5. 包含结果输出和保存功能
6. 添加执行进度提示和关键节点的状态输出

## 功能实现要求：
1. **文档加载和验证**：检查文档是否存在和格式是否正确
2. **数据处理逻辑**：实现具体的分析、提取、转换等功能
3. **结果输出**：清晰展示处理结果，包含统计信息
4. **文件保存**：如需保存结果，提供保存到文件的功能
5. **错误处理**：对可能的异常进行捕获和处理

请严格按照上述要求生成代码，确保代码可以直接复制粘贴运行。"""
                    },
                    {
                        "role": "user",
                        "content": prompt
                    }
                ],
                temperature=0.3,
                max_tokens=2000
            )
            
            # 获取生成的代码
            code = response.choices[0].message.content
            
            # 彻底清理可能的markdown格式标记
            import re
            
            # 移除开头的各种markdown代码块标记
            code = re.sub(r'^```(?:python|py)?\s*\n?', '', code, flags=re.IGNORECASE | re.MULTILINE)
            
            # 移除结尾的markdown代码块标记
            code = re.sub(r'\n?\s*```\s*$', '', code, flags=re.MULTILINE)
            
            # 移除可能的行内代码标记
            code = re.sub(r'`{1,3}([^`]+)`{1,3}', r'\1', code)
            
            # 移除可能的解释性文字段落（检测非代码内容）
            lines = code.split('\n')
            filtered_lines = []
            
            for line in lines:
                stripped = line.strip()
                # 跳过纯解释性文字行（不包含代码特征的行）
                if stripped and not any([
                    stripped.startswith('#'),  # 注释
                    stripped.startswith('import '),  # 导入语句
                    stripped.startswith('from '),  # 导入语句
                    '=' in stripped,  # 赋值语句
                    stripped.startswith('def '),  # 函数定义
                    stripped.startswith('class '),  # 类定义
                    stripped.startswith('if '),  # 条件语句
                    stripped.startswith('for '),  # 循环语句
                    stripped.startswith('while '),  # 循环语句
                    stripped.startswith('try:'),  # 异常处理
                    stripped.startswith('except'),  # 异常处理
                    stripped.startswith('finally:'),  # 异常处理
                    stripped.startswith('with '),  # 上下文管理器
                    stripped.endswith(':'),  # 代码块
                    'print(' in stripped,  # 输出语句
                    stripped.startswith('    '),  # 缩进代码
                ]):
                    # 检查是否是纯解释性文字
                    if not any(char in stripped for char in ['(', ')', '[', ']', '{', '}', '=', '.', ',']):
                        continue  # 跳过纯文字说明
                
                filtered_lines.append(line)
            
            # 重新组合代码
            code = '\n'.join(filtered_lines)
            
            return code.strip()
            
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
        
        # 可用的工具类和函数
        prompt_parts.append(f"")
        prompt_parts.append(f"## 可用工具和方法")
        prompt_parts.append(f"### DocumentAnalyzer 类的可用方法:")
        prompt_parts.append(f"- analyzer.analyze_document(file_path): 完整分析文档")
        prompt_parts.append(f"- analyzer.get_page_count(file_path): 获取文档页数")
        prompt_parts.append(f"- analyzer.analyze_structure(file_path): 分析文档结构")
        prompt_parts.append(f"- analyzer.get_document_info(file_path): 获取文档基本信息")
        prompt_parts.append(f"- analyzer.search_keyword_context(keyword, context_lines): 搜索关键词")
        prompt_parts.append(f"")
        prompt_parts.append(f"### AdvancedDocumentProcessor 类的可用方法:")
        prompt_parts.append(f"- processor.load_document(file_path): 加载并分析文档")
        prompt_parts.append(f"- processor.search_content(keyword, context_lines): 搜索内容")
        prompt_parts.append(f"- processor.export_analysis_result(): 导出分析结果")
        prompt_parts.append(f"")
        prompt_parts.append(f"### 重要提示:")
        prompt_parts.append(f"- 必须传入完整的文档路径参数")
        prompt_parts.append(f"- 使用 doc_file_path 变量作为文档路径")
        prompt_parts.append(f"- 所有方法都有完整的错误处理")
        prompt_parts.append(f"")
        prompt_parts.append(f"## 可用变量和函数")
        prompt_parts.append(f"### 文档路径变量")
        prompt_parts.append(f"- doc_file_path: 当前文档的完整文件路径（字符串）")
        prompt_parts.append(f"- doc_file_name: 当前文档的文件名")
        prompt_parts.append(f"- document_analysis: 完整的文档分析结果字典")
        prompt_parts.append(f"- doc_info: 文档基本信息（file_info）")
        prompt_parts.append(f"- doc_structure: 文档结构信息（structure_analysis）")
        prompt_parts.append(f"")
        prompt_parts.append(f"### 工作空间变量")
        prompt_parts.append(f"- user_workspace: 用户工作空间目录")
        prompt_parts.append(f"- user_exports_dir: 用户导出目录")
        prompt_parts.append(f"- user_uploads_dir: 用户上传目录")
        prompt_parts.append(f"- user_temp_dir: 用户临时目录")
        prompt_parts.append(f"")
        prompt_parts.append(f"### 实用函数")
        prompt_parts.append(f"- save_to_exports(filename, data): 保存文件到导出目录")
        prompt_parts.append(f"- get_export_path(filename): 获取导出文件路径")
        prompt_parts.append(f"- get_temp_path(filename): 获取临时文件路径")
        prompt_parts.append(f"")
        prompt_parts.append(f"### 重要提示")
        prompt_parts.append(f"- 使用 doc_file_path 变量访问当前文档")
        prompt_parts.append(f"- 首先检查 doc_file_path 是否存在和有效")
        prompt_parts.append(f"- 所有生成的文件都应保存到用户导出目录")
        prompt_parts.append(f"")
        prompt_parts.append(f"## 可用库和模块")
        prompt_parts.append(f"- 标准库: os, datetime, json, shutil, tempfile, sys, time, re, Path")
        prompt_parts.append(f"- 数据处理: BytesIO, StringIO, base64, hashlib, uuid")
        prompt_parts.append(f"- PDF处理: fitz (PyMuPDF), PyPDF2 (如果需要)")
        prompt_parts.append(f"- 进度显示: tqdm (进度条)")
        prompt_parts.append(f"- 图像处理: PIL/Pillow (通过文档分析器)")
        prompt_parts.append(f"- 文档处理: python-docx, markitdown")
        
        prompt_parts.append(f"")
        prompt_parts.append(f"## 代码生成要求")
        prompt_parts.append(f"请生成完整的Python代码来完成上述任务，必须遵循以下规范：")
        prompt_parts.append(f"")
        prompt_parts.append(f"### 格式要求：")
        prompt_parts.append(f"- 只返回纯Python代码，不要包含markdown格式")
        prompt_parts.append(f"- 不要添加任何```python 或 ``` 标记")
        prompt_parts.append(f"- 代码可以直接复制粘贴执行")
        prompt_parts.append(f"")
        prompt_parts.append(f"### 代码结构：")
        prompt_parts.append(f"1. 导入所有需要的模块")
        prompt_parts.append(f"2. 检查 doc_file_path 变量是否存在和有效")
        prompt_parts.append(f"3. 文档加载和初始化处理器")
        prompt_parts.append(f"4. 具体的数据处理逻辑")
        prompt_parts.append(f"5. 结果输出和统计信息")
        prompt_parts.append(f"6. 使用 save_to_exports() 保存结果")
        prompt_parts.append(f"")
        prompt_parts.append(f"### 代码模板示例：")
        prompt_parts.append(f"```")
        prompt_parts.append(f"# 步骤1: 导入模块")
        prompt_parts.append(f"import os")
        prompt_parts.append(f"from document_analyzer import DocumentAnalyzer")
        prompt_parts.append(f"from document_utils import AdvancedDocumentProcessor")
        prompt_parts.append(f"")
        prompt_parts.append(f"# 步骤2: 检查文档路径")
        prompt_parts.append(f"if not doc_file_path or not os.path.exists(doc_file_path):")
        prompt_parts.append(f"    print('❌ 文档文件不存在，请先选择文档')")
        prompt_parts.append(f"    exit()")
        prompt_parts.append(f"")
        prompt_parts.append(f"# 步骤3: 初始化分析器")
        prompt_parts.append(f"analyzer = DocumentAnalyzer()")
        prompt_parts.append(f"processor = AdvancedDocumentProcessor()")
        prompt_parts.append(f"")
        prompt_parts.append(f"# 步骤4: 获取基本信息")
        prompt_parts.append(f"doc_info = analyzer.get_document_info(doc_file_path)")
        prompt_parts.append(f"page_count = analyzer.get_page_count(doc_file_path)")
        prompt_parts.append(f"print(f'文档页数: {{page_count}}')")
        prompt_parts.append(f"")
        prompt_parts.append(f"# 步骤5: 分析结构")
        prompt_parts.append(f"structure = analyzer.analyze_structure(doc_file_path)")
        prompt_parts.append(f"")
        prompt_parts.append(f"# 步骤6: 保存结果")
        prompt_parts.append(f"save_to_exports('analysis_result.json', structure)")
        prompt_parts.append(f"```")
        prompt_parts.append(f"")
        prompt_parts.append(f"### 必需功能：")
        prompt_parts.append(f"- 添加中文注释解释每个步骤")
        prompt_parts.append(f"- 包含完整的错误处理")
        prompt_parts.append(f"- 显示执行进度和状态信息")
        prompt_parts.append(f"- 输出详细的处理结果")
        
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