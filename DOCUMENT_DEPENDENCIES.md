# 📚 文档分析功能依赖指南

## 🎯 概述

AI Excel 智能分析工具的文档分析功能支持 `.docx` 和 `.pdf` 格式，需要安装额外的依赖库来实现完整功能。

## 📦 常用库说明

### 必需依赖

| 库名 | 用途 | 导入示例 | 安装命令 |
|------|------|----------|----------|
| `markitdown[all]` | 统一文档转换，支持PDF、DOCX等 | `from markitdown import MarkItDown` | `pip install markitdown[all]` |
| `python-docx` | DOCX文档结构分析 | `from docx import Document` | `pip install python-docx` |
| `pymupdf` | 强大的PDF处理（推荐） | `import fitz` | `pip install pymupdf` |
| `PyPDF2` | PDF文档处理（备选） | `import PyPDF2` | `pip install PyPDF2` |
| `Pillow` | 图像处理支持 | `from PIL import Image` | `pip install Pillow` |

### 增强依赖

| 库名 | 用途 | 导入示例 | 安装命令 |
|------|------|----------|----------|
| `pdfplumber` | PDF表格和布局分析 | `import pdfplumber` | `pip install pdfplumber` |
| `docx2txt` | DOCX纯文本提取 | `import docx2txt` | `pip install docx2txt` |
| `textract` | 多格式文档提取（可选） | `import textract` | `pip install textract` |

## 🚀 快速安装

### 方式一：一键安装（推荐）
```bash
pip install -r requirements.txt
```

### 方式二：自动检查安装
```bash
python install_document_dependencies.py
```

### 方式三：手动安装
```bash
pip install markitdown[all] python-docx pymupdf PyPDF2 pdfplumber Pillow
```

## 🔧 常见问题解决

### 1. markitdown PDF 支持问题
```
错误: PdfConverter threw MissingDependencyException
解决: pip install markitdown[pdf] 或 pip install markitdown[all]
```

### 2. Windows 编译错误
```bash
# 升级构建工具
pip install --upgrade pip setuptools wheel
# 安装 Visual C++ 构建工具
```

### 3. 依赖冲突
```bash
# 创建虚拟环境
python -m venv document_env
source document_env/bin/activate  # Linux/Mac
document_env\Scripts\activate     # Windows
pip install -r requirements.txt
```

### 4. 网络问题
```bash
# 使用国内源
pip install -i https://pypi.tuna.tsinghua.edu.cn/simple/ markitdown[all]
```

## 📋 功能对照表

| 功能 | 需要的库 | 备注 |
|------|----------|------|
| PDF预览 | markitdown[pdf] | 必需 |
| DOCX预览 | python-docx, markitdown | 必需 |
| PDF结构分析 | pymupdf 或 PyPDF2 | 至少一个 |
| DOCX结构分析 | python-docx | 必需 |
| 图片提取 | Pillow, pymupdf | 可选 |
| 表格提取 | pdfplumber | 增强功能 |
| 关键词搜索 | 任何PDF/DOCX库 | 基础功能 |

## 🧪 依赖检查代码示例

```python
from typing import Optional
from pathlib import Path

# 基础导入检查
try:
    from markitdown import MarkItDown
    print("✅ markitdown 可用")
except ImportError:
    print("❌ markitdown 未安装或缺少依赖")

try:
    from docx import Document
    print("✅ python-docx 可用")
except ImportError:
    print("❌ python-docx 未安装")

try:
    import fitz  # pymupdf
    print("✅ pymupdf 可用")
except ImportError:
    print("❌ pymupdf 未安装")

# 文档分析器导入
try:
    from document_analyzer import DocumentAnalyzer
    from document_utils import AdvancedDocumentProcessor, DocumentSearchEngine
    print("✅ 文档分析模块可用")
except ImportError as e:
    print(f"❌ 文档分析模块导入失败: {e}")
```

## 🔄 版本兼容性

| Python版本 | 支持状态 | 建议版本 |
|------------|----------|----------|
| Python 3.8+ | ✅ 完全支持 | 推荐 |
| Python 3.7 | ⚠️ 部分支持 | 不推荐 |
| Python 3.6- | ❌ 不支持 | - |

## 📝 使用示例

### 基础文档分析
```python
from document_analyzer import DocumentAnalyzer

# 初始化分析器
analyzer = DocumentAnalyzer()

# 检查依赖
missing_deps = analyzer.get_missing_dependencies()
if missing_deps:
    print(f"缺少依赖: {missing_deps}")
else:
    # 分析文档
    result = analyzer.analyze_document("document.pdf")
    print("分析完成!")
```

### 高级文档处理
```python
from document_utils import AdvancedDocumentProcessor

# 初始化处理器
processor = AdvancedDocumentProcessor()

# 加载文档
document_data = processor.load_document("document.docx")

# 获取预览
preview = processor.get_document_preview(max_chars=5000)

# 搜索内容
results = processor.search_content("关键词", context_lines=3)
```

## 💡 性能优化建议

1. **PDF处理**: 优先使用 `pymupdf`，性能更好
2. **大文件处理**: 使用 `pdfplumber` 进行表格分析
3. **内存优化**: 处理大文档时及时清理缓存
4. **多格式支持**: 安装完整的 `markitdown[all]` 依赖

## 🔗 相关链接

- [markitdown 官方文档](https://github.com/microsoft/markitdown)
- [python-docx 文档](https://python-docx.readthedocs.io/)
- [PyMuPDF 文档](https://pymupdf.readthedocs.io/)
- [pdfplumber 文档](https://github.com/jsvine/pdfplumber) 