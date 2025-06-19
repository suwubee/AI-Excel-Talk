# 📄 文档智能分析功能指南

## 🎯 功能概述

基于现有的Excel分析工具，我们新增了**文档智能分析功能**，支持DOCX和PDF文档的深度分析。本功能与Excel分析完全分离，但遵循相同的流程设计，提供完整的文档处理和AI分析能力。

## ✨ 核心特性

### 1. 📝 数据预览阶段
- 使用 **MarkItDown** 解析文档内容
- 清除格式，保留结构
- 支持图片预览（前5张图片）
- 限制显示前10页内容，避免内容过载
- 提供字数统计和页数估算

### 2. 🏗️ 结构化分析阶段
- **原始文档格式解析**：直接从DOCX/PDF提取结构信息
- **标题层级识别**：自动识别1-4级标题
- **字体使用分析**：统计文档中使用的字体类型
- **内容组织分析**：段落数、表格数、图片数统计

### 3. 🤖 AI智能分析阶段
- **文档用途识别**：判断文档类型（合同、报告、手册等）
- **内容主题分析**：识别核心议题和关键信息
- **结构特点评估**：分析文档的逻辑性和可读性
- **关键信息提取**：智能识别重要数据点

### 4. 💻 代码执行阶段
- **智能代码生成**：基于任务描述生成Python代码
- **关键词搜索**：支持精确搜索和上下文提取
- **批量处理**：支持多关键词同时搜索
- **结果导出**：JSON数据和Markdown报告

## 🚀 快速开始

### 1. 安装依赖
```bash
pip install -r requirements.txt
```

新增的依赖包：
- `markitdown>=0.1.0` - 文档内容提取和清洗
- `python-docx>=0.8.11` - DOCX文档处理
- `PyPDF2>=3.0.0` - PDF基础处理
- `pymupdf>=1.23.0` - PDF高级处理
- `Pillow>=9.0.0` - 图片处理

### 2. 启动应用
```bash
python run_multiuser.py
```

### 3. 选择分析模式
在主界面选择 **📄 文档分析** 模式，然后上传DOCX或PDF文件。

## 📋 使用流程

### 步骤1：文档上传
- 支持 `.docx` 和 `.pdf` 格式
- 文件大小建议小于50MB
- 支持选择已有文档或上传新文档

### 步骤2：文档预览
- **📄 文档预览** 标签页
- 查看MarkItDown清洗后的内容
- 检查文档基本信息和统计数据
- 查看结构化分析摘要

### 步骤3：AI智能分析
- **🤖 AI文档分析** 标签页
- 进行结构化分析（无需API）
- 使用AI进行深度内容分析（需要API Key）
- 通过快速操作按钮进行专项分析

### 步骤4：代码执行
- **💻 代码执行** 标签页
- 使用AI生成文档处理代码
- 手动编写和执行Python代码
- 自动替换文档路径变量

### 步骤5：搜索和导出
- **🔍 搜索工具** 标签页
- 单个或批量关键词搜索
- 查看搜索结果和上下文
- 导出分析报告和数据

## 🛠️ 核心模块

### 1. DocumentAnalyzer
主要的文档分析器，提供：
- 文档类型检测
- MarkItDown内容提取
- 结构化信息分析
- 关键词搜索功能

```python
from document_analyzer import DocumentAnalyzer

analyzer = DocumentAnalyzer()
result = analyzer.analyze_document("document.docx")
```

### 2. AdvancedDocumentProcessor
高级文档处理器，提供：
- 文档加载和预览
- 结构摘要生成
- 搜索引擎接口
- 分析结果导出

```python
from document_utils import AdvancedDocumentProcessor

processor = AdvancedDocumentProcessor()
analysis = processor.load_document("document.pdf")
```

### 3. EnhancedDocumentAIAnalyzer
AI分析器，提供：
- 深度内容分析
- 智能对话功能
- 代码生成建议
- 分析任务推荐

```python
from document_ai_analyzer import EnhancedDocumentAIAnalyzer

ai_analyzer = EnhancedDocumentAIAnalyzer(api_key, base_url, model)
analysis = ai_analyzer.analyze_document_structure(document_data)
```

## 📊 分析示例

### 关键词搜索示例
```python
# 搜索合同编号及其上下文
search_results = processor.search_content("合同编号", context_lines=3)

for result in search_results:
    print(f"位置: 第{result['line_number']}行")
    print(f"匹配内容: {result['matched_line']}")
    print(f"上下文:\n{result['context']}")
```

### 批量分析示例
```python
# 批量搜索多个关键词
keywords = ["甲方", "乙方", "金额", "日期", "条款"]
search_engine = DocumentSearchEngine(processor)
report = search_engine.generate_search_report(keywords)
```

### AI代码生成示例
```python
# AI生成文档处理代码
task = "搜索所有包含'重要条款'的段落并提取关键信息"
code = ai_analyzer.generate_document_code_solution(task, document_data, filename)
```

## 🎯 高级功能

### 1. 图片预览
- 自动提取文档中的图片（前5张）
- 转换为Base64格式在界面中显示
- 支持PDF中的图片提取

### 2. 字体分析
- **DOCX**: 提取Run级别的字体信息
- **PDF**: 基于字体名称和大小进行分析
- 智能识别主要字体和样式

### 3. 标题识别
- **DOCX**: 基于样式名称识别标题层级
- **PDF**: 基于字体大小智能推断标题层级
- 支持1-4级标题的自动分类

### 4. 上下文搜索
- 精确关键词匹配
- 可配置上下文行数（1-10行）
- 支持大小写不敏感搜索
- 显示关键词在行中的位置

## 🔧 配置说明

### 环境变量
- `OPENAI_API_KEY`: OpenAI API密钥（用于AI分析）
- `OPENAI_BASE_URL`: API基础URL（可选）

### 参数配置
- **预览限制**: 默认20,000字符，约10页内容
- **搜索上下文**: 默认3行，可调整为1-10行
- **图片数量**: 最多显示5张图片
- **字体统计**: 最多显示前10种字体

## 📝 最佳实践

### 1. 文档准备
- 确保文档格式正确且未损坏
- 对于大型文档，建议先进行结构化分析
- PDF文档的OCR识别效果取决于原始质量

### 2. 关键词搜索
- 使用具体的关键词而非通用词汇
- 利用批量搜索功能提高效率
- 适当调整上下文行数以获得完整信息

### 3. AI分析
- 提供准确的任务描述以获得更好的代码生成
- 结合结构化分析结果进行深度AI分析
- 使用快速操作按钮进行标准化分析

### 4. 结果导出
- 及时导出重要的分析结果
- JSON格式适合程序处理，Markdown适合阅读
- 利用用户工作空间管理导出文件

## 🚨 注意事项

### 1. 安全性
- 用户文档完全隔离，互不干扰
- 支持自动文件清理和隐私保护
- 敏感文档请谨慎使用在线AI服务

### 2. 性能
- 大型文档（>50MB）可能处理较慢
- PDF文档的分析时间取决于页数和复杂度
- AI分析需要网络连接和API配额

### 3. 兼容性
- 支持现代浏览器（Chrome、Firefox、Safari、Edge）
- 需要Python 3.8+环境
- 部分功能需要internet连接

## 🛠️ 故障排除

### 常见问题

**Q: 文档上传失败？**
A: 
- 检查文件格式是否为.docx或.pdf
- 确认文件大小不超过限制
- 验证文件是否损坏

**Q: MarkItDown预览为空？**
A:
- 确认markitdown包已正确安装
- 检查文档是否包含可提取的文本内容
- 尝试重新上传文档

**Q: 结构分析无结果？**
A:
- 检查文档是否包含标题样式
- 验证python-docx或pymupdf是否正常工作
- 查看错误日志获取详细信息

**Q: AI分析失败？**
A:
- 验证API Key配置是否正确
- 检查网络连接状态
- 确认API配额是否充足

## 📞 技术支持

如需技术支持或功能建议，请：
1. 查看日志错误信息
2. 检查依赖包安装状态
3. 验证配置参数设置
4. 提供具体的错误复现步骤

---

**🎉 开始使用文档智能分析功能，体验AI驱动的文档处理新体验！** 