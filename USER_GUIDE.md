# 📖 用户指南 - AI Excel 智能分析工具

## 📋 目录
- [快速入门](#-快速入门)
- [功能详解](#-功能详解)
- [多用户特性](#-多用户特性)
- [智能文件拦截](#-智能文件拦截)
- [使用技巧](#-使用技巧)
- [常见问题](#-常见问题)

## 🚀 快速入门

### 第一次使用

1. **启动应用**
   ```bash
   python run_multiuser.py
   ```

2. **基础配置**
   - 访问 `http://localhost:8501`
   - 侧边栏输入 OpenAI API Key
   - 选择模型（推荐 `deepseek-v3`）
   - 配置会自动保存，下次访问时自动恢复

3. **上传文件**
   - 支持 `.xlsx` 和 `.xls` 格式
   - 首次上传后，系统会记住您的文件
   - 可以在下拉菜单中选择已上传的文件

### 界面导航

应用分为四个主要标签页：

- **📋 数据预览** - 查看和管理Excel数据
- **🤖 AI分析** - 智能分析和对话交互  
- **💻 代码执行** - Python代码编写和执行
- **🛠️ 数据工具** - 数据清洗和导出工具

## 📊 功能详解

### 📋 数据预览与管理

#### Excel文件管理
- **上传新文件**: 直接拖拽或点击上传
- **选择已有文件**: 从下拉菜单选择之前上传的文件
- **工作表切换**: 支持多工作表Excel文件
- **数据统计**: 实时显示行数、列数、缺失值等

#### 数据质量检查
```
数据行数: 1000        数据列数: 15
缺失值: 25           重复行: 3
```

### 🤖 AI智能分析

#### 深度业务分析
AI基于您的完整数据样本进行分析：
- **业务场景识别** - 自动识别数据的业务用途
- **数据关系分析** - 发现工作表间的关联关系
- **关键指标发现** - 识别重要的业务指标
- **分析建议** - 提供具体的分析方向

#### 智能对话
您可以用中文向AI提问：
```
"分析这个销售数据的趋势"
"找出利润率最高的产品类别"
"推荐下一步的分析方向"
```

#### 快速操作
点击预设的快速操作按钮：
- 🎯 业务场景识别
- 🔗 数据关系分析  
- 💎 关键指标发现
- 📊 分析机会推荐

### 💻 代码执行环境

#### 可用变量
系统自动为您准备以下变量：
```python
# Excel数据（每个工作表对应一个DataFrame）
df_Sheet1, df_Sheet2, df_销售数据, df_产品信息

# 文件信息
excel_file_path      # 原始Excel文件路径
excel_file_name      # 文件名
sheet_names          # 所有工作表名称列表

# 用户工作空间
user_session_id      # 您的会话ID
user_workspace       # 您的专属工作空间
user_exports_dir     # 导出文件目录
```

#### 代码编写
```python
# 基础数据分析
print(f"数据概况: {df_Sheet1.shape}")
print(f"列名: {list(df_Sheet1.columns)}")

# 数据处理
processed_df = df_Sheet1.copy()
processed_df['利润率'] = processed_df['利润'] / processed_df['销售额'] * 100

# 数据筛选
high_profit = processed_df[processed_df['利润率'] > 20]

# 保存修改（重要！）
df_Sheet1 = processed_df  # 将修改保存回原变量
```

#### AI代码助手
- 点击 **🤖 AI助手** 按钮
- 描述您要完成的任务
- AI自动生成相应的Python代码
- 一键插入到编辑器中

### 🛠️ 数据工具

#### 数据清洗
- **缺失值处理**: 选择列和填充方法
- **重复数据**: 自动检测和清理
- **数据类型**: 自动优化数据类型

#### 统计分析
- **描述性统计**: 生成完整的统计摘要
- **分组分析**: 按类别进行统计
- **汇总报告**: 生成新的统计工作表

#### 文件导出
- **Excel导出**: 保存修改后的Excel文件
- **自定义文件名**: 支持时间戳命名
- **导出管理**: 查看和下载所有导出文件

## 🔐 多用户特性

### 数据隔离保护

每个用户拥有完全独立的工作空间：
```
user_uploads/
├── user_abc123/          # 您的专属空间
│   ├── uploads/         # 上传的Excel文件
│   ├── exports/         # 导出的文件
│   └── temp/           # 临时文件
└── user_def456/         # 其他用户空间（您无法访问）
```

### 会话管理

- **智能会话ID**: 基于浏览器特征生成稳定ID
- **自动超时**: 24小时后自动清理过期数据
- **数据恢复**: 同一浏览器会话中数据自动恢复

### 配置管理

#### 自动保存
- 输入API Key后自动保存到服务器
- 浏览器localStorage缓存非敏感配置
- 页面刷新后自动恢复配置

#### 配置安全
- **服务器端**: 加密保存完整配置
- **浏览器端**: 脱敏显示敏感信息
- **手动控制**: 可随时清除所有配置

### 隐私保护

- **敏感信息脱敏**: API Key显示为 `sk-1234****cdef`
- **自动清理**: 过期会话自动删除
- **手动清理**: 随时清理个人数据
- **访问控制**: 严格防止跨用户访问

## 🔄 智能文件拦截

### 自动拦截机制

系统会自动拦截并重定向以下文件保存操作：

#### 1. Excel文件保存
```python
# 原始代码
df.to_excel("分析结果.xlsx", index=False)

# 自动重定向到: user_exports/timestamp_分析结果.xlsx
```

#### 2. JSON数据保存
```python
# 原始代码
with open("分析报告.json", "w") as f:
    json.dump(data, f)

# 自动重定向到: user_exports/timestamp_分析报告.json
```

#### 3. Markdown文档保存
```python
# 原始代码
with open("分析说明.md", "w") as f:
    f.write(content)

# 自动重定向到: user_exports/timestamp_分析说明.md
```

### 自动下载界面

执行完成后，系统会自动显示：

```
📁 生成的文件
🎉 检测到 3 个生成的文件

📄 JSON数据文件:
📄 20241203_143022_分析报告.json (2.3 KB)  [⬇️ 下载]

📝 Markdown分析文件:  
📝 20241203_143022_分析说明.md (5.1 KB)   [⬇️ 下载]

📊 Excel文件:
📊 20241203_143022_分析结果.xlsx (45.7 KB) [⬇️ 下载]
```

### 文件管理

在 **🛠️ 数据工具** → **📁 我的导出文件** 中：
- 查看所有生成的文件
- 按时间排序显示
- 一键下载或删除
- 查看存储使用统计

## 💡 使用技巧

### Excel数据处理最佳实践

#### 1. 变量命名规则
```python
# 工作表名会自动转换为变量名
# "销售数据" → df_销售数据
# "Sheet 1" → df_Sheet_1  
# "产品-信息" → df_产品_信息
```

#### 2. 保存修改
```python
# 错误：修改不会保存
df_销售数据['新列'] = df_销售数据['金额'] * 1.1

# 正确：将修改保存回原变量
processed_df = df_销售数据.copy()
processed_df['新列'] = processed_df['金额'] * 1.1
df_销售数据 = processed_df  # 保存修改
```

#### 3. 跨工作表分析
```python
# 合并多个工作表
merged_data = pd.merge(
    df_销售数据, 
    df_产品信息, 
    on='产品ID', 
    how='left'
)
```

### AI分析技巧

#### 1. 提问技巧
```
具体问题 ✅ "分析Q1季度销售额下降的原因"
模糊问题 ❌ "帮我分析数据"

业务导向 ✅ "识别高价值客户群体的特征"  
技术导向 ❌ "做个机器学习模型"
```

#### 2. 充分利用AI对话
- 先让AI分析整体数据结构
- 基于AI建议深入特定方向
- 结合业务背景提供上下文

### 文件操作技巧

#### 1. 有意义的文件命名
```python
# 推荐
timestamp = datetime.now().strftime("%Y%m%d_%H%M")
df.to_excel(f"销售分析报告_{timestamp}.xlsx")

# 不推荐  
df.to_excel("result.xlsx")
```

#### 2. 批量处理
```python
# 为不同分析结果创建不同文件
for category in categories:
    result = analyze_category(category)
    result.to_excel(f"{category}_分析结果.xlsx")
```

#### 3. 结构化导出
```python
# 同时生成多种格式
analysis_data.to_excel("详细数据.xlsx")           # Excel原始数据
summary.to_json("分析摘要.json")                # 程序可读格式
report.to_file("分析报告.md")                   # 人类可读格式
```

## ❓ 常见问题

### 配置问题

**Q: API Key输入后丢失怎么办？**
A: 
- 检查是否点击了"💾 保存配置"
- 清除浏览器缓存后重新输入
- 使用"手动保存配置"功能

**Q: 页面刷新后配置消失？**
A: 
- 配置会自动从localStorage恢复
- 如果没有恢复，说明localStorage被清除
- 重新输入配置并保存即可

### 文件问题

**Q: 上传的Excel文件找不到？**
A:
- 检查文件格式是否为.xlsx或.xls
- 查看"选择已有文件"下拉菜单
- 如果列表为空，需要重新上传

**Q: 代码执行时提示变量不存在？**
A:
- 检查工作表名是否正确
- 使用 `print(sheet_names)` 查看可用工作表
- 确保变量名格式正确（空格替换为下划线）

### 代码执行问题

**Q: 数据修改后没有保存？**
A:
```python
# 必须将修改保存回原变量
df_工作表名 = 修改后的数据
```

**Q: 文件保存不成功？**
A:
- 检查文件名是否包含特殊字符
- 确保磁盘空间充足
- 查看控制台错误信息

**Q: AI代码生成不符合需求？**
A:
- 提供更详细的任务描述
- 包含具体的业务背景
- 可以多次生成并选择最佳方案

### 多用户问题

**Q: 如何确保数据隐私？**
A:
- 每个用户数据完全隔离
- 24小时自动清理过期数据
- 可随时手动清理个人数据

**Q: 会话过期怎么办？**
A:
- 系统会自动生成新的会话ID
- 保存的配置会自动恢复
- 需要重新上传Excel文件

**Q: 存储空间不足怎么办？**
A:
- 在"数据工具"中查看存储使用
- 删除不需要的导出文件
- 使用"清理我的数据"功能

### 性能问题

**Q: 大文件处理很慢？**
A:
- Excel文件建议小于50MB
- 复杂分析可以分批处理
- 使用数据抽样进行初步分析

**Q: AI分析响应缓慢？**
A:
- 检查网络连接是否正常
- 尝试使用更快的AI模型
- 简化分析任务的复杂度

## 🎯 高级技巧

### 复杂数据分析

#### 时间序列分析
```python
# 确保日期列格式正确
df_数据['日期'] = pd.to_datetime(df_数据['日期'])
df_数据.set_index('日期', inplace=True)

# 趋势分析
monthly_trend = df_数据.resample('M').sum()
monthly_trend.to_excel("月度趋势分析.xlsx")
```

#### 多维度分析
```python
# 透视表分析
pivot_table = pd.pivot_table(
    df_销售数据,
    values='销售额',
    index='产品类别',
    columns='销售月份',
    aggfunc='sum',
    fill_value=0
)
pivot_table.to_excel("产品销售透视表.xlsx")
```

### AI分析包生成

如果您需要生成完整的AI分析报告：

```python
from generate_ai_analysis_package import AdvancedAIAnalyzer

# 创建分析器
analyzer = AdvancedAIAnalyzer(
    api_key=st.session_state.api_key,
    model=st.session_state.selected_model
)

# 生成完整分析包
json_file, md_file = analyzer.generate_complete_analysis_package(
    data=df_主要数据,
    save_dir=".",  # 自动重定向到用户目录
    business_context="销售数据分析"  # 可选：提供业务背景
)

print(f"✅ 分析包已生成:")
print(f"📄 数据文件: {json_file}")
print(f"📝 分析报告: {md_file}")
```

执行后会自动生成JSON数据文件和Markdown分析报告，并提供下载链接。

---

**💡 提示**: 如需了解更多技术细节，请查看[开发者指南](DEVELOPER_GUIDE.md)。 