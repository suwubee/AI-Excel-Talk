# 🚀 AI Excel 智能分析工具 - 多用户版

## 📋 项目简介

一款基于AI的Excel智能分析工具，支持多用户并发使用，提供完整的数据隔离和隐私保护。用户可以上传Excel文件，通过AI进行深度分析，并在安全的隔离环境中执行Python代码进行数据处理。

## ✨ 核心特性

- 🤖 **AI智能分析** - 深度业务理解、自然语言交互、智能代码生成
- 🔐 **多用户支持** - 完全数据隔离、会话管理、并发安全、隐私保护
- 💻 **代码执行环境** - 安全隔离、Excel原生支持、文件自动拦截
- 📁 **智能文件管理** - 自动拦截保存、分类下载、用户专属空间

## 🚀 快速开始

### 1. 环境准备
```bash
# 要求: Python 3.8+ 和 OpenAI API Key

# 安装依赖
pip install -r requirements.txt
```

### 2. 启动应用
```bash
# 方式一: 使用启动脚本（推荐）
python run_multiuser.py

# 方式二: 直接启动
streamlit run app_enhanced_multiuser.py

# 方式三: 指定参数
python run_multiuser.py --port 8502 --debug
```

### 3. 配置使用
1. 浏览器访问 `http://localhost:8501`
2. 侧边栏输入 OpenAI API Key
3. 选择模型（推荐 deepseek-v3）
4. 上传Excel文件开始分析

## 🤖 自动化部署

### GitHub Actions CI/CD

本项目已集成完整的GitHub Actions工作流，提供：

- ✅ **自动代码检查** - 语法验证、文件完整性检查
- ✅ **Docker构建测试** - 确保容器化部署可用
- ✅ **多平台配置** - 自动生成各平台部署配置
- ✅ **自动发布** - 创建Release版本和安装包
- ✅ **部署指导** - 提供详细的部署说明

### 一键部署到云端

**🌟 推荐：Streamlit Cloud（免费）**
1. Fork这个仓库
2. 访问 [share.streamlit.io](https://share.streamlit.io)
3. 连接仓库，主文件设为 `app_enhanced_multiuser.py`
4. 点击Deploy！

**🐳 Docker部署**
```bash
# 使用docker-compose
docker-compose up -d

# 或使用Docker直接运行
docker run -p 8501:8501 -v $(pwd)/user_uploads:/app/user_uploads ai-excel-tool
```

**📚 详细部署指南**
- [GitHub Actions指南](GITHUB_DEPLOYMENT_GUIDE.md) - 自动化部署流程
- [完整部署指南](DEPLOYMENT.md) - 多平台部署详解

## 📊 主要功能

### 四大核心模块
1. **📋 数据预览** - Excel文件管理、多工作表预览、数据质量分析
2. **🤖 AI分析** - 业务场景识别、数据关系分析、自然语言对话
3. **💻 代码执行** - Python编辑器、安全执行、文件自动拦截下载
4. **🛠️ 数据工具** - 数据清洗、统计分析、导出管理

### 智能文件拦截 🔥
```python
# 代码中的任何文件保存操作都会被自动拦截并重定向：
df.to_excel("分析结果.xlsx")                    # 自动重定向到用户目录
with open("报告.json", "w") as f: ...           # 自动拦截并分类
with open("分析.md", "w") as f: ...             # 提供直接下载链接
```

## 🔐 安全特性

- **数据隔离** - 用户间完全独立的文件系统
- **隐私保护** - 24小时自动清理过期数据，敏感信息脱敏
- **访问控制** - 基于会话的权限验证，防止路径遍历

## 📖 使用示例

### 基础数据处理
```python
# 分析数据
print(f"数据概况: {df_Sheet1.shape}")
processed_df = df_Sheet1.copy()
processed_df['新列'] = processed_df['原列'] * 1.2

# 保存结果（自动拦截到用户目录）
processed_df.to_excel("处理结果.xlsx")
```

### AI分析包生成
```python
# 生成完整AI分析包
from generate_ai_analysis_package import AdvancedAIAnalyzer
analyzer = AdvancedAIAnalyzer(api_key="your_key")
json_file, md_file = analyzer.generate_complete_analysis_package(df_Sheet1, ".")
# 执行完成后会自动显示下载链接
```

## 📚 文档指南

- **📖 [用户指南](USER_GUIDE.md)** - 详细的功能使用说明和最佳实践
- **🔧 [开发者指南](DEVELOPER_GUIDE.md)** - 技术架构、项目结构和扩展开发
- **🚀 [GitHub部署指南](GITHUB_DEPLOYMENT_GUIDE.md)** - 自动化部署流程详解
- **☁️ [部署指南](DEPLOYMENT.md)** - 多平台部署方案

## 🔄 版本信息

**当前版本**: v2.2.0 🔥
- ✅ **AI分析智能化革命** - 业务导向提示词，数据故事挖掘
- ✅ **Excel结构分析完全通用化** - 纯统计学算法，100%兼容性
- ✅ **技术稳定性提升** - pandas兼容性修复，代码库精简
- ✅ 智能文件保存拦截功能
- ✅ 多用户数据隔离和隐私保护

**📋 [查看v2.2.0详细升级说明](UPDATE_v2.2.0.md)**

## 💬 支持与反馈

- 🐛 **问题报告**: 提交 [GitHub Issue](https://github.com/your-repo/issues)
- 💡 **功能建议**: 欢迎提交 Pull Request
- 📧 **技术支持**: 查看文档或在GitHub Discussion讨论

## 📄 许可证

本项目采用 MIT 许可证，详情请查看 LICENSE 文件。

---

**🎉 立即开始体验安全、智能的Excel数据分析！**

> 💡 **提示**: 首次使用建议先查看[用户指南](USER_GUIDE.md)了解完整功能。 