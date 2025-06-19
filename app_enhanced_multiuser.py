"""
AI Excel 智能分析工具 - 多用户增强版
支持多用户并发、文件隔离和隐私保护
"""

import streamlit as st
import pandas as pd
import numpy as np
import openai
import io
import json
import re
from typing import Dict, List, Any, Tuple
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import traceback
import tempfile
import os
import hashlib
from pathlib import Path
import streamlit.components.v1 as components
from user_session_manager import UserSessionManager, UserConfigManager
from excel_utils import AdvancedExcelProcessor, DataAnalyzer

# 设置pandas选项，避免FutureWarning
pd.set_option('future.no_silent_downcasting', True)

# 页面配置
st.set_page_config(
    page_title="AI Excel 智能分析工具 - 多用户版",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 添加JavaScript代码用于localStorage操作
st.markdown("""
<script>
// localStorage操作函数
window.setLocalStorageItem = function(key, value) {
    localStorage.setItem(key, JSON.stringify(value));
};

window.getLocalStorageItem = function(key) {
    const item = localStorage.getItem(key);
    return item ? JSON.parse(item) : null;
};

window.removeLocalStorageItem = function(key) {
    localStorage.removeItem(key);
};

// 页面加载时立即检查localStorage并恢复配置
window.addEventListener('load', function() {
    console.log('🔄 页面加载完成，检查localStorage配置');
    
    // 查找配置缓存
    let foundConfig = null;
    
    // 加载localStorage配置
    for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (key && key.startsWith('ai_excel_config_')) {
            const value = localStorage.getItem(key);
            if (value) {
                try {
                    const config = JSON.parse(value);
                    console.log('🔄 找到localStorage配置:', key);
                    
                    // 显示脱敏配置信息
                    const displayConfig = {...config};
                    if (displayConfig.api_key && displayConfig.api_key.length > 8) {
                        displayConfig.api_key = config.api_key.substring(0, 4) + '****' + config.api_key.substring(config.api_key.length - 4);
                    }
                    console.log('🔄 配置内容（脱敏）:', displayConfig);
                    
                    foundConfig = config;
                    break;
                } catch (e) {
                    console.error('🔄 localStorage配置解析失败:', e);
                }
            }
        }
    }
    
    if (foundConfig) {
        console.log('🔄 localStorage配置恢复完成，将通知Streamlit');
        
        // 创建一个全局标记，表示有localStorage配置需要恢复
        window.streamlitLocalStorageConfig = {
            found: true,
            config: foundConfig,
            timestamp: new Date().toISOString()
        };
        
        // 创建一个DOM元素来标记localStorage配置已恢复
        const indicator = document.createElement('div');
        indicator.id = 'localStorage_config_indicator';
        indicator.style.display = 'none';
        indicator.setAttribute('data-restored', 'true');
        indicator.setAttribute('data-api-key', foundConfig.api_key || '');
        indicator.setAttribute('data-base-url', foundConfig.base_url || '');
        indicator.setAttribute('data-model', foundConfig.selected_model || '');
        document.body.appendChild(indicator);
        
        console.log('✅ localStorage配置已准备就绪');
    } else {
        console.log('🔄 没有找到localStorage配置');
        window.streamlitLocalStorageConfig = {
            found: false,
            config: null,
            timestamp: new Date().toISOString()
        };
    }
});

// 为Streamlit提供localStorage访问接口
window.streamlitLocalStorage = {
    set: window.setLocalStorageItem,
    get: window.getLocalStorageItem,
    remove: window.removeLocalStorageItem
};
</script>
""", unsafe_allow_html=True)

# 统一的内容容器样式函数
def get_unified_content_styles():
    """获取统一的内容容器样式"""
    return """
    <style>
    /* 基础内容容器样式 */
    .content-container-base {
        overflow-y: auto;
        border-radius: 8px;
        padding: 18px;
        font-size: 14px;
        line-height: 1.6;
        margin: 10px 0;
    }
    
    /* 文档预览容器 */
    .document-preview-container {
        max-height: 500px;
        border: 1px solid #4caf50;
        background-color: #f1f8e9;
        font-family: 'Source Code Pro', monospace;
    }
    
    /* 文档结构容器 */
    .document-structure-container {
        max-height: 400px;
        border: 1px solid #2196f3;
        background-color: #f3f8ff;
        font-family: 'Roboto', sans-serif;
    }
    
    /* AI分析结果容器 */
    .ai-analysis-container {
        max-height: 600px;
        border: 1px solid #ff5722;
        background-color: #fff3e0;
        font-family: 'Roboto', sans-serif;
        box-shadow: 0 2px 4px rgba(255, 87, 34, 0.1);
    }
    
    /* Excel结构分析容器 */
    .excel-structure-container {
        max-height: 450px;
        border: 1px solid #4caf50;
        background-color: #f1f8e9;
        font-family: 'Monaco', 'Menlo', monospace;
        font-size: 13px;
        line-height: 1.5;
    }
    
    /* Excel AI分析容器 */
    .excel-ai-container {
        max-height: 550px;
        border: 1px solid #ff5722;
        background-color: #fff3e0;
        font-family: 'Roboto', sans-serif;
        box-shadow: 0 2px 4px rgba(255, 87, 34, 0.1);
    }
    
    /* 对话历史容器 */
    .chat-container-base {
        max-height: 500px;
        overflow-y: auto;
        border-radius: 8px;
        padding: 18px;
        font-family: 'Roboto', sans-serif;
        font-size: 14px;
        line-height: 1.5;
    }
    
    .doc-chat-container {
        border: 1px solid #ff9800;
        background-color: #fff8e1;
    }
    
    .excel-chat-container {
        border: 1px solid #9c27b0;
        background-color: #f3e5f5;
    }
    
    /* 对话气泡样式 */
    .chat-user {
        background-color: #e3f2fd;
        border-left: 4px solid #2196f3;
        padding: 12px;
        margin: 10px 0;
        border-radius: 5px;
    }
    
    .chat-ai {
        background-color: #e8f5e8;
        border-left: 4px solid #4caf50;
        padding: 12px;
        margin: 10px 0;
        border-radius: 5px;
    }
    
    .chat-divider {
        border-top: 1px solid #d0d0d0;
        margin: 15px 0;
    }
    
    /* 标题样式 */
    .content-container-base h1, .content-container-base h2, .content-container-base h3 {
        margin-top: 18px;
        margin-bottom: 10px;
    }
    
    .document-preview-container h1, .document-preview-container h2, .document-preview-container h3,
    .excel-structure-container h1, .excel-structure-container h2, .excel-structure-container h3 {
        color: #2e7d32;
    }
    
    .document-structure-container h1, .document-structure-container h2, .document-structure-container h3 {
        color: #1976d2;
    }
    
    .ai-analysis-container h1, .ai-analysis-container h2, .ai-analysis-container h3,
    .excel-ai-container h1, .excel-ai-container h2, .excel-ai-container h3 {
        color: #d84315;
        border-bottom: 2px solid #ffccbc;
        padding-bottom: 5px;
    }
    
    /* 列表和段落样式 */
    .content-container-base ul, .content-container-base ol {
        margin-left: 18px;
        margin-bottom: 10px;
    }
    
    .content-container-base p {
        margin-bottom: 10px;
        text-align: justify;
    }
    
    .content-container-base strong {
        font-weight: 600;
    }
    
    .document-preview-container strong, .excel-structure-container strong {
        color: #388e3c;
    }
    
    .ai-analysis-container strong, .excel-ai-container strong {
        color: #ff5722;
    }
    
    /* 代码样式 */
    .excel-structure-container code {
        background-color: #e8f5e8;
        padding: 2px 4px;
        border-radius: 3px;
        font-family: 'Courier New', monospace;
    }
    
    /* 引用样式 */
    .ai-analysis-container blockquote, .excel-ai-container blockquote {
        border-left: 4px solid #ff5722;
        padding-left: 15px;
        margin: 15px 0;
        background-color: #fbe9e7;
        font-style: italic;
    }
    </style>
    """

def render_content_container(content: str, container_type: str) -> None:
    """
    统一的内容容器渲染函数
    
    Args:
        content: 要显示的内容
        container_type: 容器类型 ('document-preview', 'document-structure', 'ai-analysis', 
                       'excel-structure', 'excel-ai')
    """
    # 如果还没有添加样式，先添加
    if not hasattr(st.session_state, '_unified_styles_added'):
        st.markdown(get_unified_content_styles(), unsafe_allow_html=True)
        st.session_state._unified_styles_added = True
    
    # 根据类型选择合适的CSS类
    container_class = f"{container_type}-container content-container-base"
    
    # 渲染内容
    st.markdown(
        f'<div class="{container_class}">{content}</div>',
        unsafe_allow_html=True
    )

def render_chat_container(chat_history: List[Dict], container_type: str = 'doc-chat') -> None:
    """
    统一的对话历史渲染函数
    
    Args:
        chat_history: 对话历史列表
        container_type: 容器类型 ('doc-chat' 或 'excel-chat')
    """
    # 添加样式
    if not hasattr(st.session_state, '_unified_styles_added'):
        st.markdown(get_unified_content_styles(), unsafe_allow_html=True)
        st.session_state._unified_styles_added = True
    
    # 构建对话内容HTML
    chat_html = f'<div class="{container_type}-container chat-container-base">'
    
    for i, chat in enumerate(chat_history):
        if chat["role"] == "user":
            if container_type == 'doc-chat':
                chat_html += f'<div class="chat-user"><strong>👤 用户第 {(i//2) + 1} 次提问：</strong><br/>{chat["content"]}</div>'
            else:
                chat_html += f'<div class="chat-user">👤 {chat["content"]}</div>'
        else:
            chat_html += f'<div class="chat-ai"><strong>🤖 AI回答：</strong><br/>{chat["content"]}</div>'
        
        if i < len(chat_history) - 1:
            chat_html += '<div class="chat-divider"></div>'
    
    chat_html += '</div>'
    
    # 渲染对话内容
    st.markdown(chat_html, unsafe_allow_html=True)

# 自定义CSS样式（保持原有样式）
st.markdown("""
<style>
    .main-header {
        font-size: 2.8rem;
        font-weight: bold;
        text-align: center;
        margin-bottom: 2rem;
        background: linear-gradient(90deg, #1f77b4, #ff7f0e);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
    }
    .chat-container {
        max-height: 400px;
        overflow-y: auto;
        padding: 1rem;
        border: 1px solid #ddd;
        border-radius: 10px;
        background-color: #f8f9fa;
        margin-bottom: 1rem;
    }
    .user-message {
        background: linear-gradient(135deg, #007bff, #0056b3);
        color: white;
        padding: 12px;
        border-radius: 15px 15px 5px 15px;
        margin: 8px 0;
        text-align: right;
        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
    }
    .ai-message {
        background: linear-gradient(135deg, #28a745, #1e7e34);
        color: white;
        padding: 12px;
        border-radius: 15px 15px 15px 5px;
        margin: 8px 0;
        text-align: left;
        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
    }
    .excel-preview {
        max-height: 600px;
        overflow-y: auto;
        border: 2px solid #ddd;
        border-radius: 10px;
        padding: 15px;
        background: linear-gradient(to bottom, #ffffff, #f8f9fa);
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .feature-card {
        background: white;
        padding: 1rem;
        border-radius: 10px;
        border: 1px solid #e0e0e0;
        margin: 0.5rem 0;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .metric-card {
        background: linear-gradient(135deg, #e3f2fd, #bbdefb);
        padding: 1rem;
        border-radius: 8px;
        text-align: center;
        margin: 0.5rem;
    }
    .session-info {
        background: linear-gradient(135deg, #f3e5f5, #e1bee7);
        padding: 10px;
        border-radius: 8px;
        font-size: 0.8rem;
        margin: 5px 0;
    }
    .config-saved {
        background-color: #d4edda;
        color: #155724;
        padding: 8px;
        border-radius: 5px;
        border: 1px solid #c3e6cb;
        margin: 5px 0;
        font-size: 0.9rem;
    }
    .config-loaded {
        background-color: #d1ecf1;
        color: #0c5460;
        padding: 8px;
        border-radius: 5px;
        border: 1px solid #bee5eb;
        margin: 5px 0;
        font-size: 0.9rem;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 24px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        padding-left: 20px;
        padding-right: 20px;
        background-color: #f0f2f6;
        border-radius: 10px 10px 0px 0px;
        color: #262730;
        font-size: 16px;
        font-weight: 600;
    }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(90deg, #1f77b4, #ff7f0e);
        color: white;
    }
</style>
""", unsafe_allow_html=True)

class EnhancedAIAnalyzer:
    """增强版AI分析器（保持原有功能）"""
    
    def __init__(self, api_key: str, base_url: str = None, model: str = "gpt-4.1-mini"):
        self.client = openai.OpenAI(
            api_key=api_key,
            base_url=base_url if base_url else None
        )
        self.model = model
    
    def analyze_excel_structure(self, excel_data: Dict[str, pd.DataFrame]) -> str:
        """深度智能分析Excel文件结构和业务逻辑"""
        try:
            prompt = "作为资深的业务数据分析专家，请对以下Excel文件进行深度业务理解和分析。我将提供每个工作表的前50行完整数据供您分析：\n\n"
            
            for sheet_name, df in excel_data.items():
                prompt += f"## 📋 工作表: {sheet_name}\n"
                prompt += f"- 数据规模: {len(df)}行 × {len(df.columns)}列\n"
                prompt += f"- 字段列表: {list(df.columns)}\n\n"
                
                if not df.empty:
                    # 提取前50行数据用于分析
                    sample_size = min(50, len(df))
                    sample_df = df.head(sample_size)
                    
                    prompt += f"### 📊 前{sample_size}行完整数据:\n"
                    prompt += "```\n"
                    # 修复to_string参数兼容性问题
                    try:
                        # 尝试使用width参数
                        data_string = sample_df.to_string(max_rows=sample_size, max_cols=20, width=None)
                    except TypeError:
                        # 如果width参数不支持，则不使用该参数
                        data_string = sample_df.to_string(max_rows=sample_size, max_cols=20)
                    except Exception:
                        # 如果还有其他错误，使用最基本的to_string
                        data_string = sample_df.to_string()
                    
                    prompt += data_string
                    prompt += "\n```\n\n"
                    
                    # 字段特征分析
                    prompt += f"### 🔍 字段特征分析:\n"
                    for col in df.columns:
                        try:
                            # 基本统计
                            non_null_count = df[col].count()
                            total_count = len(df)
                            null_count = total_count - non_null_count
                            
                            prompt += f"**{col}**:\n"
                            prompt += f"  - 数据完整性: {non_null_count}/{total_count} 非空 ({null_count}个缺失值)\n"
                            
                            # 数据类型分析
                            dtype_info = str(df[col].dtype)
                            prompt += f"  - 数据类型: {dtype_info}\n"
                            
                            # 唯一值分析
                            if non_null_count > 0:
                                unique_count = df[col].nunique()
                                prompt += f"  - 唯一值数量: {unique_count}\n"
                                
                                # 获取示例值（修复tolist错误）
                                non_null_series = df[col].dropna()
                                if len(non_null_series) > 0:
                                    sample_values = non_null_series.head(5).values.tolist()
                                    prompt += f"  - 示例值: {sample_values}\n"
                                
                                # 对于数值类型，提供统计信息
                                if pd.api.types.is_numeric_dtype(df[col]):
                                    stats = df[col].describe()
                                    prompt += f"  - 数值范围: [{stats['min']:.2f} - {stats['max']:.2f}]\n"
                                    prompt += f"  - 平均值: {stats['mean']:.2f}, 中位数: {stats['50%']:.2f}\n"
                                
                                # 对于文本类型，提供频次分析
                                elif df[col].dtype == 'object':
                                    try:
                                        value_counts = df[col].value_counts()
                                        if len(value_counts) > 0:
                                            top_values = value_counts.head(5)
                                            # 修复FutureWarning: 使用.iloc而不是位置索引
                                            top_values_dict = {}
                                            for i in range(len(top_values)):
                                                key = top_values.index[i]
                                                value = top_values.iloc[i]
                                                top_values_dict[key] = value
                                            prompt += f"  - 高频值: {top_values_dict}\n"
                                            
                                            # 文本长度分析
                                            text_lengths = df[col].dropna().astype(str).str.len()
                                            if len(text_lengths) > 0:
                                                avg_length = text_lengths.mean()
                                                max_length = text_lengths.max()
                                                prompt += f"  - 文本长度: 平均{avg_length:.1f}字符, 最大{max_length}字符\n"
                                    except Exception as e:
                                        prompt += f"  - 频次分析出错: {str(e)}\n"
                        
                        except Exception as e:
                            prompt += f"**{col}**: 分析出错 - {str(e)}\n"
                        
                        prompt += "\n"
                
                prompt += "\n" + "="*50 + "\n\n"
            
            # 分析提示
            prompt += """
作为资深业务数据分析师，请基于上述真实数据进行深度业务理解和价值挖掘分析：

🎯 **业务洞察与价值发现**:
- 基于数据模式和内容，推断这是什么业务场景（如销售、运营、财务、人力等）
- 识别数据背后的核心业务问题和管理关注点
- 分析数据的决策支撑价值和管理应用场景
- 发现数据中隐含的业务规律和趋势

📊 **数据故事与商业逻辑**:
- 从数据中读出"故事"：数据反映了什么业务现状和问题
- 分析不同工作表的业务关联性和数据流向逻辑
- 识别关键业务节点和数据流转环节
- 找出可能的业务瓶颈、机会点或风险点

💡 **实际分析价值与建议**:
- 基于数据特征，提出3-5个具体的、可操作的分析方向
- 每个分析方向都要说明：为什么重要、如何分析、预期收益
- 建议具体的分析方法和可用工具（如透视表、图表类型等）
- 指出数据中最有价值的分析维度和切入点

🎨 **可视化与呈现建议**:
- 针对关键指标，推荐最适合的图表类型和展示方式
- 建议制作哪些管理看板或报告
- 提出数据呈现的最佳实践，让非技术人员也能理解

请避免单纯的技术性描述，重点关注业务价值和实际应用，用业务语言而非技术术语进行表达。
"""
            
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": "你是一位具有15年经验的资深业务数据分析师和商业顾问。你擅长从真实数据中洞察业务本质，发现商业价值，并提供可操作的分析建议。你的分析风格注重实用性和业务价值，能够将复杂的数据转化为清晰的商业洞察，帮助管理者做出明智决策。"},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.7,
                max_tokens=3000
            )
            
            return response.choices[0].message.content
            
        except Exception as e:
            return f"❌ AI分析出错: {str(e)}"
    
    def chat_with_data(self, message: str, excel_data: Dict[str, pd.DataFrame], context: str = "") -> str:
        """与数据对话（保持原有功能）"""
        try:
            # 构建数据摘要
            data_summary = "当前Excel数据概况：\n"
            for sheet_name, df in excel_data.items():
                data_summary += f"- {sheet_name}: {len(df)}行 × {len(df.columns)}列\n"
                data_summary += f"  字段: {', '.join(df.columns.tolist()[:10])}\n"
                if len(df.columns) > 10:
                    data_summary += f"  (还有{len(df.columns)-10}个字段...)\n"
            
            prompt = f"""
你是一位专业的数据分析师。基于以下Excel数据信息回答用户问题：

{data_summary}

已有分析上下文：
{context}

用户问题：{message}

请提供专业、具体的分析建议，用中文回答。
"""
            
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": "你是一位专业的数据分析师，善于理解业务需求并提供实用的分析建议。"},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.7,
                max_tokens=1500
            )
            
            return response.choices[0].message.content
            
        except Exception as e:
            return f"❌ AI对话出错: {str(e)}"

    def generate_enhanced_code_solution(self, task_description: str, enhanced_excel_data: Dict, excel_filename: str) -> str:
        """生成增强的Excel代码解决方案，包含完整的Excel文件和工作表关系信息"""
        try:
            # 构建更详细的Excel结构信息
            excel_structure_info = f"Excel文件: {excel_filename}\n\n"
            
            # 工作表概览
            excel_structure_info += "工作表结构概览:\n"
            for sheet_name, sheet_data in enhanced_excel_data.items():
                safe_name = sheet_name.replace(' ', '_').replace('-', '_').replace('.', '_')
                excel_structure_info += f"\n📋 工作表: {sheet_name} (变量名: df_{safe_name})\n"
                excel_structure_info += f"  - 数据规模: {sheet_data['shape'][0]}行 × {sheet_data['shape'][1]}列\n"
                excel_structure_info += f"  - 列名: {sheet_data['columns']}\n"
                excel_structure_info += f"  - 数据类型: {sheet_data['dtypes']}\n"
                
                if sheet_data['sample_data']:
                    excel_structure_info += f"  - 数据样例:\n"
                    for col, values in list(sheet_data['sample_data'].items())[:3]:
                        sample_vals = list(values.values())[:2]
                        excel_structure_info += f"    * {col}: {sample_vals}\n"
            
            # 可用变量信息
            excel_structure_info += f"\n可用变量:\n"
            excel_structure_info += f"- excel_file_path: 原始Excel文件路径\n"
            excel_structure_info += f"- excel_file_name: 文件名 ({excel_filename})\n"
            excel_structure_info += f"- sheet_names: 所有工作表名称列表\n"
            excel_structure_info += f"- sheet_info: 工作表详细信息字典\n"
            
            for sheet_name in enhanced_excel_data.keys():
                safe_name = sheet_name.replace(' ', '_').replace('-', '_').replace('.', '_')
                excel_structure_info += f"- df_{safe_name}: {sheet_name}工作表的DataFrame\n"

            prompt = f"""
任务描述: {task_description}

Excel文件完整结构信息:
{excel_structure_info}

请生成完整的Python代码来完成这个任务。代码要求：

1. **Excel文件级别操作**: 
   - 可以使用excel_file_path访问原始Excel文件
   - 可以使用pd.read_excel(excel_file_path, sheet_name='xxx')读取特定工作表
   - 支持复杂的Excel操作，如条件读取、自定义解析等

2. **跨工作表分析**:
   - 分析不同工作表之间的业务关系
   - 识别关联字段和数据流向
   - 实现跨表数据合并、比较、验证

3. **综合数据处理**:
   - 使用pandas进行高级数据处理
   - 包含必要的错误处理和数据验证
   - 添加详细的中文注释说明业务逻辑

4. **结果输出**:
   - 提供清晰的处理结果和统计信息
   - 如果需要修改数据，确保将结果保存回相应的df_变量
   - 包含执行进度提示和关键节点输出

5. **代码结构**:
   - 确保代码可以直接执行
   - 包含必要的导入语句
   - 结构清晰，逻辑分明

特别注意：
- 如果任务涉及多个工作表，请分析它们的业务关系
- 如果需要生成新的汇总或分析结果，请创建新的变量
- 所有对原数据的修改都要保存回对应的df_变量中
- 充分利用sheet_names和sheet_info变量来实现动态的工作表处理

请只返回纯Python代码，不要包含任何markdown格式标记。
"""
            
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": "你是一个Excel数据分析和Python编程专家，专门生成高质量的Excel处理代码。你深度理解Excel文件结构、工作表关系和业务数据分析。只返回纯Python代码，不要包含任何markdown格式标记。"},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.1,
                max_tokens=2000
            )
            
            # 清理可能的markdown格式
            code = response.choices[0].message.content
            
            # 移除可能的markdown代码块标记
            import re
            # 移除开头的```python或```
            code = re.sub(r'^```(?:python)?\n?', '', code, flags=re.MULTILINE)
            # 移除结尾的```
            code = re.sub(r'\n?```$', '', code, flags=re.MULTILINE)
            
            return code.strip()
            
        except Exception as e:
            return f"# 增强代码生成失败: {str(e)}"


def get_session_id():
    """获取或生成用户会话ID（在同一浏览器会话中保持稳定）"""
    try:
        if 'user_session_id' not in st.session_state or not st.session_state.user_session_id:
            # 尝试基于稳定的信息生成会话ID（不包含时间戳）
            try:
                user_agent = st.context.headers.get('user-agent', 'unknown_browser')
            except Exception:
                user_agent = 'unknown_browser'
            
            # 使用更稳定的标识符：用户代理 + 日期（不包含具体时间）
            stable_identifier = f"{user_agent}_{datetime.now().strftime('%Y%m%d')}"
            session_hash = hashlib.md5(stable_identifier.encode()).hexdigest()[:16]
            
            # 生成稳定的session_id（同一天内保持一致）
            st.session_state.user_session_id = f"user_{session_hash}"
            
            print(f"[DEBUG] 生成稳定会话ID: {st.session_state.user_session_id}")
        else:
            print(f"[DEBUG] 使用现有会话ID: {st.session_state.user_session_id}")
        
        # 确保返回值不为空
        if not st.session_state.user_session_id:
            # 如果还是空，生成一个随机的session_id
            import uuid
            st.session_state.user_session_id = f"user_{uuid.uuid4().hex[:16]}"
            print(f"[DEBUG] 生成随机会话ID: {st.session_state.user_session_id}")
        
        return st.session_state.user_session_id
        
    except Exception as e:
        print(f"[ERROR] 会话ID生成失败: {e}")
        # 生成一个紧急备用的session_id
        import uuid
        backup_session_id = f"user_{uuid.uuid4().hex[:16]}"
        st.session_state.user_session_id = backup_session_id
        print(f"[DEBUG] 使用备用会话ID: {backup_session_id}")
        return backup_session_id


def save_to_browser_cache(config: Dict[str, Any], config_manager: UserConfigManager, session_id: str):
    """保存配置到浏览器localStorage（保存真实配置）"""
    try:
        # 对于localStorage，我们保存真实的配置（用户本地浏览器是安全的）
        real_config = config.copy()
        real_config['cached_at'] = datetime.now().isoformat()
        real_config['cache_type'] = 'browser_real'
        
        # 同时创建脱敏版本用于显示
        safe_config = config_manager.get_config_for_browser_cache(config)
        
        # 保存到session state（脱敏版本）
        st.session_state.browser_cached_config = safe_config
        print(f"[DEBUG] 保存到session state (脱敏): {safe_config}")
        
        # 保存到服务器文件（脱敏版本）
        file_success = config_manager.save_browser_cache_config(session_id, config)
        print(f"[DEBUG] 服务器文件保存: {'成功' if file_success else '失败'}")
        
        # 保存到浏览器localStorage（真实配置）
        storage_key = f"ai_excel_config_{session_id[:16]}"
        browser_success = set_browser_storage_item(storage_key, real_config)
        print(f"[DEBUG] localStorage保存真实配置: {'成功' if browser_success else '失败'}")
        
        return file_success and browser_success
        
    except Exception as e:
        print(f"[ERROR] 浏览器缓存保存失败: {e}")
        return False

def get_browser_storage_config(session_id: str):
    """从localStorage读取配置到session state"""
    storage_key = f"ai_excel_config_{session_id[:16]}"
    
    # 创建JavaScript来读取localStorage并写入DOM
    html_code = f"""
    <script>
        (function() {{
            const key = '{storage_key}';
            const value = localStorage.getItem(key);
            
            // 清除之前的配置div
            const oldDiv = document.getElementById('localStorage_config_reader');
            if (oldDiv) {{
                oldDiv.remove();
            }}
            
            // 创建新的配置div
            const configDiv = document.createElement('div');
            configDiv.id = 'localStorage_config_reader';
            configDiv.style.display = 'none';
            
            if (value) {{
                try {{
                    const config = JSON.parse(value);
                    console.log('📖 localStorage读取成功:', key);
                    console.log('📖 读取的配置:', config);
                    
                    // 将配置数据写入div属性
                    configDiv.setAttribute('data-found', 'true');
                    configDiv.setAttribute('data-config', value);
                    configDiv.setAttribute('data-key', key);
                    
                    // 显示脱敏版本到控制台
                    const displayConfig = {{...config}};
                    if (displayConfig.api_key && displayConfig.api_key.length > 8) {{
                        displayConfig.api_key = displayConfig.api_key.substring(0, 4) + '****' + displayConfig.api_key.substring(displayConfig.api_key.length - 4);
                    }}
                    console.table(displayConfig);
                    
                }} catch (e) {{
                    console.error('📖 localStorage配置解析失败:', e);
                    configDiv.setAttribute('data-found', 'error');
                    configDiv.setAttribute('data-error', e.message);
                }}
            }} else {{
                console.log('📖 localStorage中没有找到配置:', key);
                configDiv.setAttribute('data-found', 'false');
            }}
            
            document.body.appendChild(configDiv);
            
            // 触发事件通知配置已读取
            const event = new CustomEvent('localStorageConfigRead', {{ 
                detail: {{ key: key, found: !!value }}
            }});
            window.dispatchEvent(event);
        }})();
    </script>
    <div id="localStorage_reader_container" style="height: 1px;"></div>
    """
    
    components.html(html_code, height=1)
    
    # 尝试从session state缓存的localStorage数据中读取
    cached_config = st.session_state.get('localStorage_cached_config')
    if cached_config and cached_config.get('session_id') == session_id:
        print(f"[DEBUG] 从session state缓存读取localStorage配置: {cached_config['config'].keys()}")
        return cached_config['config']
    
    return None

def load_browser_cache_config(config_manager: UserConfigManager, session_id: str):
    """从服务器文件加载浏览器缓存配置"""
    try:
        # 从服务器文件加载浏览器缓存
        file_config = config_manager.load_browser_cache_config(session_id)
        
        if file_config:
            print(f"[DEBUG] 从服务器文件加载浏览器缓存: {file_config}")
            return file_config
        
        print(f"[DEBUG] 没有找到浏览器缓存配置文件")
        return None
        
    except Exception as e:
        print(f"[ERROR] 浏览器缓存加载失败: {e}")
        return None

def try_read_localStorage_direct(session_id: str):
    """尝试直接从localStorage读取配置并缓存到session state"""
    storage_key = f"ai_excel_config_{session_id[:16]}"
    
    # 创建JavaScript来尝试读取localStorage并将结果写入session state
    html_code = f"""
    <script>
        (function() {{
            const key = '{storage_key}';
            const value = localStorage.getItem(key);
            
            if (value) {{
                try {{
                    const config = JSON.parse(value);
                    console.log('🔄 直接读取localStorage配置成功:', key);
                    
                    // 创建一个带有特殊ID的div来传递配置
                    const resultDiv = document.createElement('div');
                    resultDiv.id = 'localStorage_direct_result';
                    resultDiv.style.display = 'none';
                    resultDiv.setAttribute('data-success', 'true');
                    resultDiv.setAttribute('data-config', JSON.stringify(config));
                    document.body.appendChild(resultDiv);
                    
                    console.log('🔄 配置已写入DOM，等待Python读取');
                    
                    // 显示脱敏信息
                    const displayConfig = {{...config}};
                    if (displayConfig.api_key && displayConfig.api_key.length > 8) {{
                        displayConfig.api_key = displayConfig.api_key.substring(0, 4) + '****' + displayConfig.api_key.substring(displayConfig.api_key.length - 4);
                    }}
                    console.log('🔄 脱敏配置:', displayConfig);
                    
                }} catch (e) {{
                    console.error('🔄 localStorage读取失败:', e);
                    const resultDiv = document.createElement('div');
                    resultDiv.id = 'localStorage_direct_result';
                    resultDiv.style.display = 'none';
                    resultDiv.setAttribute('data-success', 'false');
                    resultDiv.setAttribute('data-error', e.message);
                    document.body.appendChild(resultDiv);
                }}
            }} else {{
                console.log('🔄 localStorage中没有配置');
                const resultDiv = document.createElement('div');
                resultDiv.id = 'localStorage_direct_result';
                resultDiv.style.display = 'none';
                resultDiv.setAttribute('data-success', 'false');
                resultDiv.setAttribute('data-reason', 'not_found');
                document.body.appendChild(resultDiv);
            }}
        }})();
    </script>
    <div style="height: 1px;"></div>
    """
    
    components.html(html_code, height=1)
    return storage_key


def simulate_localStorage_recovery(config_manager: UserConfigManager, session_id: str):
    """基于服务器端文件模拟localStorage配置恢复"""
    try:
        print(f"[DEBUG] === 模拟localStorage恢复 ===")
        print(f"[DEBUG] 检查会话ID: {session_id}")
        
        # 获取用户工作空间
        workspace = config_manager.session_manager.get_user_workspace(session_id)
        print(f"[DEBUG] 用户工作空间: {workspace}")
        
        if workspace:
            # 检查各种配置文件是否存在
            config_file = workspace / "user_config.json"
            cache_file = workspace / "browser_cache.json"
            
            print(f"[DEBUG] 用户配置文件: {config_file}")
            print(f"[DEBUG] 配置文件存在: {config_file.exists()}")
            print(f"[DEBUG] 浏览器缓存文件: {cache_file}")
            print(f"[DEBUG] 缓存文件存在: {cache_file.exists()}")
            
            # 如果有缓存文件，说明之前localStorage保存过
            if cache_file.exists():
                print(f"[DEBUG] 检测到浏览器缓存文件，模拟localStorage恢复")
                
                # 获取完整的服务器配置
                full_config = config_manager.load_user_config(session_id)
                if full_config:
                    print(f"[DEBUG] 模拟localStorage恢复成功: API Key={'已设置' if full_config.get('api_key') else '未设置'}")
                    return full_config
                else:
                    print(f"[DEBUG] 无法获取完整配置进行localStorage模拟")
                    return None
            else:
                print(f"[DEBUG] 没有检测到浏览器缓存文件，无localStorage配置")
                return None
        else:
            print(f"[DEBUG] 用户工作空间不存在")
            return None
    except Exception as e:
        print(f"[ERROR] localStorage模拟恢复失败: {e}")
        return None


def load_user_config(config_manager: UserConfigManager, session_id: str):
    """加载用户配置（优先使用localStorage中的真实配置）"""
    try:
        print(f"[DEBUG] === 开始加载用户配置 ===")
        print(f"[DEBUG] 会话ID: {session_id}")
        print(f"[DEBUG] 会话ID前16位: {session_id[:16]}")
        
        # 首先尝试模拟localStorage恢复
        localStorage_config = None
        print(f"[DEBUG] 尝试模拟localStorage恢复...")
        localStorage_config = simulate_localStorage_recovery(config_manager, session_id)
        
        # 然后尝试从服务器端加载配置
        saved_config = config_manager.load_user_config(session_id)
        print(f"[DEBUG] 服务器端配置: {saved_config is not None}")
        
        # 最后尝试从服务器端浏览器缓存加载
        browser_cache_config = config_manager.load_browser_cache_config(session_id)
        if browser_cache_config:
            print(f"[DEBUG] 服务器端浏览器缓存: {browser_cache_config.keys()}")
            # 将浏览器缓存配置保存到session state中（用于显示）
            st.session_state.browser_cached_config = browser_cache_config
        
        # 配置优先级：模拟localStorage > 服务器完整配置 > 服务器浏览器缓存 > 默认值
        final_config = {}
        config_source = ""
        
        # 1. 优先使用模拟的localStorage配置
        if localStorage_config:
            final_config.update(localStorage_config)
            config_source = "localStorage"
            print(f"[DEBUG] 使用模拟localStorage配置作为主配置")
        
        # 2. 如果没有localStorage，使用服务器端完整配置
        elif saved_config:
            final_config.update(saved_config)
            config_source = "服务器端完整配置"
            print(f"[DEBUG] 使用服务器端完整配置")
        
        # 3. 如果都没有，使用服务器端浏览器缓存
        elif browser_cache_config:
            for key, value in browser_cache_config.items():
                if key not in ['cached_at', 'cache_type']:
                    final_config[key] = value
            config_source = "服务器端浏览器缓存"
            print(f"[DEBUG] 使用服务器端浏览器缓存")
        
        # 4. 如果都没有，使用默认值
        if not final_config:
            final_config = {
                'base_url': 'https://apistudy.mycache.cn/v1',
                'selected_model': 'deepseek-v3'
            }
            config_source = "默认配置"
            print(f"[DEBUG] 使用默认配置")
        
        # 将最终配置加载到session state
        if 'api_key' in final_config:
            st.session_state.saved_api_key = final_config['api_key']
        else:
            # 确保清除旧的API Key
            if 'saved_api_key' in st.session_state:
                del st.session_state.saved_api_key
        
        if 'base_url' in final_config:
            st.session_state.saved_base_url = final_config['base_url']
        if 'selected_model' in final_config:
            st.session_state.saved_model = final_config['selected_model']
        
        # 保存配置来源信息
        st.session_state.config_source = config_source
        
        print(f"[DEBUG] 最终配置加载完成: API Key={'已设置' if final_config.get('api_key') else '未设置'}")
        print(f"[DEBUG] 配置来源: {config_source}")
        print(f"[DEBUG] 配置详情: localStorage={localStorage_config is not None}, 服务器={saved_config is not None}, 缓存={browser_cache_config is not None}")
        print(f"[DEBUG] === 配置加载完成 ===")
        
        return final_config
        
    except Exception as e:
        print(f"[ERROR] 配置加载失败: {e}")
        return None


def save_user_config(config_manager: UserConfigManager, session_id: str, config: Dict[str, Any]):
    """保存用户配置到服务器端"""
    return config_manager.save_user_config(session_id, config)


def auto_save_config(config_manager: UserConfigManager, session_id: str, api_key: str, base_url: str, selected_model: str):
    """自动保存配置"""
    config_to_save = {
        'api_key': api_key,
        'base_url': base_url,
        'selected_model': selected_model,
        'save_timestamp': datetime.now().isoformat(),
        'auto_saved': True
    }
    
    # 保存到服务器
    success = save_user_config(config_manager, session_id, config_to_save)
    
    # 保存到session state
    if success:
        st.session_state.saved_api_key = api_key
        st.session_state.saved_base_url = base_url
        st.session_state.saved_model = selected_model
    
    return success


def set_browser_storage_item(key: str, value: Any):
    """设置浏览器localStorage项目"""
    try:
        # 将值转换为JSON字符串，注意转义
        json_value = json.dumps(value).replace('"', '\\"').replace("'", "\\'")
        
        html_code = f"""
        <script>
            try {{
                const value = '{json_value}';
                const parsedValue = JSON.parse(value);
                localStorage.setItem('{key}', value);
                console.log('✅ 已保存到localStorage:', '{key}', parsedValue);
                
                // 验证保存是否成功
                const saved = localStorage.getItem('{key}');
                if (saved) {{
                    console.log('✅ 验证localStorage保存成功:', JSON.parse(saved));
                }} else {{
                    console.error('❌ localStorage保存失败');
                }}
            }} catch (e) {{
                console.error('❌ localStorage保存出错:', e);
            }}
        </script>
        <div style="display: none; height: 1px;">设置localStorage</div>
        """
        
        components.html(html_code, height=1)
        return True
    except Exception as e:
        print(f"[ERROR] 设置localStorage失败: {e}")
        return False

def remove_browser_storage_item(key: str):
    """从浏览器localStorage删除项目"""
    try:
        html_code = f"""
        <script>
            try {{
                localStorage.removeItem('{key}');
                console.log('✅ 已从localStorage删除:', '{key}');
            }} catch (e) {{
                console.error('❌ localStorage删除出错:', e);
            }}
        </script>
        <div style="display: none; height: 1px;">删除localStorage</div>
        """
        
        components.html(html_code, height=1)
        return True
    except Exception as e:
        print(f"[ERROR] 删除localStorage失败: {e}")
        return False

def get_browser_cache_setting(session_id: str):
    """从localStorage获取浏览器缓存设置"""
    setting_key = f"ai_excel_browser_cache_setting_{session_id[:16]}"
    
    html_code = f"""
    <script>
        (function() {{
            const key = '{setting_key}';
            const value = localStorage.getItem(key);
            
            // 清除之前的设置div
            const oldDiv = document.getElementById('browser_cache_setting_reader');
            if (oldDiv) {{
                oldDiv.remove();
            }}
            
            // 创建新的设置div
            const settingDiv = document.createElement('div');
            settingDiv.id = 'browser_cache_setting_reader';
            settingDiv.style.display = 'none';
            
            if (value) {{
                try {{
                    const setting = JSON.parse(value);
                    console.log('📖 浏览器缓存设置读取成功:', key, setting);
                    settingDiv.setAttribute('data-found', 'true');
                    settingDiv.setAttribute('data-enabled', setting.enabled ? 'true' : 'false');
                    
                    // 通知Python更新session state
                    window.localStorage_browser_cache_setting = setting;
                    
                }} catch (e) {{
                    console.error('📖 浏览器缓存设置解析失败:', e);
                    settingDiv.setAttribute('data-found', 'false');
                }}
            }} else {{
                console.log('📖 localStorage中没有浏览器缓存设置，使用默认值');
                settingDiv.setAttribute('data-found', 'false');
                window.localStorage_browser_cache_setting = null;
            }}
            
            document.body.appendChild(settingDiv);
        }})();
    </script>
    <div style="height: 1px;"></div>
    """
    
    components.html(html_code, height=1)
    return setting_key

def save_browser_cache_setting(session_id: str, enabled: bool):
    """保存浏览器缓存设置到localStorage"""
    setting_key = f"ai_excel_browser_cache_setting_{session_id[:16]}"
    setting_value = {"enabled": enabled, "updated_at": datetime.now().isoformat()}
    
    return set_browser_storage_item(setting_key, setting_value)

def try_load_browser_cache_setting(session_id: str):
    """尝试从localStorage加载浏览器缓存设置"""
    setting_key = f"ai_excel_browser_cache_setting_{session_id[:16]}"
    
    # 创建JavaScript来读取localStorage设置并直接应用
    html_code = f"""
    <script>
        (function() {{
            const key = '{setting_key}';
            const value = localStorage.getItem(key);
            
            if (value) {{
                try {{
                    const setting = JSON.parse(value);
                    const enabled = setting.enabled;
                    console.log('🔧 从localStorage读取浏览器缓存设置:', enabled);
                    
                    // 如果设置为false，需要通过重新加载页面来应用
                    if (!enabled) {{
                        // 在页面URL中添加一个特殊参数来标记需要关闭缓存
                        const currentUrl = new URL(window.location);
                        const hasParam = currentUrl.searchParams.has('browser_cache_disabled');
                        
                        if (!hasParam) {{
                            currentUrl.searchParams.set('browser_cache_disabled', 'true');
                            console.log('🔧 检测到浏览器缓存已关闭，重新加载页面应用设置');
                            window.location.href = currentUrl.toString();
                            return;
                        }}
                    }} else {{
                        // 如果设置为true，移除禁用参数
                        const currentUrl = new URL(window.location);
                        if (currentUrl.searchParams.has('browser_cache_disabled')) {{
                            currentUrl.searchParams.delete('browser_cache_disabled');
                            console.log('🔧 检测到浏览器缓存已开启，重新加载页面应用设置');
                            window.location.href = currentUrl.toString();
                            return;
                        }}
                    }}
                    
                }} catch (e) {{
                    console.error('🔧 localStorage浏览器缓存设置解析失败:', e);
                }}
            }} else {{
                console.log('🔧 localStorage中没有浏览器缓存设置，使用默认值');
            }}
        }})();
    </script>
    """
    
    components.html(html_code, height=0)
    
    # 检查URL参数来确定当前设置
    query_params = st.query_params
    if 'browser_cache_disabled' in query_params:
        st.session_state.browser_cache_enabled = False
        print(f"[DEBUG] 从URL参数检测到浏览器缓存已关闭")
    else:
        # 保持默认值或已设置的值
        print(f"[DEBUG] 浏览器缓存设置: {st.session_state.browser_cache_enabled}")

def init_browser_cache_setting(session_id: str):
    """初始化浏览器缓存设置，从localStorage读取或使用默认值"""
    setting_key = f"ai_excel_browser_cache_setting_{session_id[:16]}"
    
    # 默认设置为开启
    default_enabled = True
    
    # 首先检查session_state中是否已有设置
    if 'browser_cache_enabled' not in st.session_state:
        st.session_state.browser_cache_enabled = default_enabled
    
    # 创建JavaScript来检查localStorage并通过URL参数传递设置
    html_code = f"""
    <script>
        (function() {{
            const key = '{setting_key}';
            const value = localStorage.getItem(key);
            
            let enabled = {str(default_enabled).lower()};  // 默认值
            
            if (value) {{
                try {{
                    const setting = JSON.parse(value);
                    enabled = setting.enabled;
                    console.log('🔧 从localStorage读取浏览器缓存设置:', enabled);
                }} catch (e) {{
                    console.error('🔧 localStorage浏览器缓存设置解析失败:', e);
                    enabled = {str(default_enabled).lower()};
                }}
            }} else {{
                console.log('🔧 localStorage中没有浏览器缓存设置，使用默认值:', enabled);
            }}
            
            // 创建唯一的div标记来传递设置
            const settingDiv = document.createElement('div');
            settingDiv.id = 'browser_cache_setting_init_{session_id[:8]}';
            settingDiv.style.display = 'none';
            settingDiv.setAttribute('data-enabled', enabled.toString());
            settingDiv.setAttribute('data-key', key);
            settingDiv.setAttribute('data-session', '{session_id[:8]}');
            document.body.appendChild(settingDiv);
            
            console.log('🔧 浏览器缓存设置初始化完成:', enabled);
        }})();
    </script>
    <div style="height: 1px;"></div>
    """
    
    components.html(html_code, height=1)
    
    print(f"[DEBUG] 初始化浏览器缓存设置: {st.session_state.browser_cache_enabled}")
    return setting_key

def init_localStorage_config(session_id: str):
    """初始化时从localStorage自动恢复配置"""
    storage_key = f"ai_excel_config_{session_id[:16]}"
    
    # 创建JavaScript代码来自动恢复localStorage配置
    html_code = f"""
    <script>
        (function() {{
            const key = '{storage_key}';
            const value = localStorage.getItem(key);
            
            if (value) {{
                try {{
                    const config = JSON.parse(value);
                    
                    // 将配置保存到一个全局变量供Python读取
                    window.streamlitLocalStorageConfig = config;
                    
                    // 触发一个自定义事件通知配置已恢复
                    const event = new CustomEvent('localStorageConfigLoaded', {{ 
                        detail: config 
                    }});
                    window.dispatchEvent(event);
                    
                }} catch (e) {{
                    console.error('🔄 localStorage配置恢复失败:', e);
                    window.streamlitLocalStorageConfig = null;
                }}
            }} else {{
                console.log('🔄 localStorage中没有配置需要恢复');
                window.streamlitLocalStorageConfig = null;
            }}
        }})();
    </script>
    """
    
    components.html(html_code, height=0)


def restore_localStorage_to_session_state(session_id: str):
    """模拟从localStorage恢复配置到session state（基于检查结果）"""
    # 由于无法直接从JavaScript获取数据，我们基于localStorage的存在来恢复配置
    # 这是一个间接的方法，但对于页面刷新场景是有效的
    
    # 首先检查localStorage
    check_localStorage_and_restore(session_id)
    
    # 基于localStorage可能存在的配置，尝试重建session state
    # 这里我们可以检查是否应该有localStorage配置
    storage_key = f"ai_excel_config_{session_id[:16]}"
    
    # 使用一个特殊的JavaScript来尝试读取并缓存配置
    cache_html = f"""
    <script>
        (function() {{
            const key = '{storage_key}';
            const value = localStorage.getItem(key);
            
            if (value) {{
                try {{
                    const config = JSON.parse(value);
                    
                    // 将配置存储到一个全局缓存变量
                    window.streamlit_localStorage_cache = {{
                        session_id: '{session_id}',
                        config: config,
                        cached_at: new Date().toISOString()
                    }};
                    
                    console.log('💾 localStorage配置已缓存到全局变量');
                    
                }} catch (e) {{
                    console.error('💾 localStorage配置缓存失败:', e);
                    window.streamlit_localStorage_cache = null;
                }}
            }} else {{
                window.streamlit_localStorage_cache = null;
            }}
        }})();
    </script>
    """
    
    st.markdown(cache_html, unsafe_allow_html=True)
    
    return storage_key

def check_localStorage_and_restore(session_id: str):
    """检查localStorage并尝试恢复配置到session state"""
    storage_key = f"ai_excel_config_{session_id[:16]}"
    
    # 使用JavaScript检查localStorage并自动恢复配置
    restore_html = f"""
    <script>
        (function() {{
            const key = '{storage_key}';
            const value = localStorage.getItem(key);
            
            console.log('🔄 检查localStorage配置恢复，键:', key);
            
            if (value) {{
                try {{
                    const config = JSON.parse(value);
                    console.log('🔄 找到localStorage配置:', config);
                    
                    // 将配置标记为已恢复
                    window.localStorage_config_restored = {{
                        session_id: '{session_id}',
                        config: config,
                        restored_at: new Date().toISOString()
                    }};
                    
                    // 创建恢复状态div
                    const restoreDiv = document.createElement('div');
                    restoreDiv.id = 'localStorage_restore_indicator';
                    restoreDiv.style.display = 'none';
                    restoreDiv.setAttribute('data-restored', 'true');
                    restoreDiv.setAttribute('data-session', '{session_id}');
                    restoreDiv.setAttribute('data-has-api-key', config.api_key ? 'true' : 'false');
                    restoreDiv.setAttribute('data-base-url', config.base_url || '');
                    restoreDiv.setAttribute('data-model', config.selected_model || '');
                    
                    // 安全地设置API key（仅前后4位）
                    if (config.api_key && config.api_key.length > 8) {{
                        restoreDiv.setAttribute('data-api-key-preview', config.api_key.substring(0, 4) + '****' + config.api_key.substring(config.api_key.length - 4));
                    }}
                    
                    document.body.appendChild(restoreDiv);
                    
                    console.log('✅ localStorage配置恢复完成');
                    console.log('✅ API Key:', config.api_key ? '已设置' : '未设置');
                    console.log('✅ Base URL:', config.base_url || '未设置');
                    console.log('✅ Model:', config.selected_model || '未设置');
                    
                }} catch (e) {{
                    console.error('❌ localStorage配置恢复失败:', e);
                    
                    const restoreDiv = document.createElement('div');
                    restoreDiv.id = 'localStorage_restore_indicator';
                    restoreDiv.style.display = 'none';
                    restoreDiv.setAttribute('data-restored', 'false');
                    restoreDiv.setAttribute('data-error', e.message);
                    document.body.appendChild(restoreDiv);
                }}
            }} else {{
                console.log('🔄 localStorage中没有配置');
                
                const restoreDiv = document.createElement('div');
                restoreDiv.id = 'localStorage_restore_indicator';
                restoreDiv.style.display = 'none';
                restoreDiv.setAttribute('data-restored', 'false');
                restoreDiv.setAttribute('data-reason', 'not_found');
                document.body.appendChild(restoreDiv);
            }}
        }})();
    </script>
    <div style="height: 1px; display: none;">localStorage检查</div>
    """
    
    components.html(restore_html, height=1)
    
    # 检查是否已经有localStorage恢复的配置缓存
    if 'localStorage_restored_config' in st.session_state:
        cached = st.session_state.localStorage_restored_config
        if cached.get('session_id') == session_id:
            print(f"[DEBUG] 使用已缓存的localStorage配置: API Key={'已设置' if cached['config'].get('api_key') else '未设置'}")
            return cached['config']
    
    return None

def main():
    """主应用程序"""
    
    # 初始化用户会话管理器
    if 'session_manager' not in st.session_state:
        st.session_state.session_manager = UserSessionManager(
            base_upload_dir="user_uploads",
            session_timeout_hours=24,
            cleanup_interval_minutes=60
        )
        st.session_state.config_manager = UserConfigManager(st.session_state.session_manager)
    
    session_manager = st.session_state.session_manager
    config_manager = st.session_state.config_manager
    
    # 初始化 session state
    if 'user_session_id' not in st.session_state:
        st.session_state.user_session_id = None
    if 'excel_data' not in st.session_state:
        st.session_state.excel_data = {}
    if 'current_sheet' not in st.session_state:
        st.session_state.current_sheet = None
    if 'chat_history' not in st.session_state:
        st.session_state.chat_history = []
    if 'excel_analysis' not in st.session_state:
        st.session_state.excel_analysis = ""
    if 'excel_processor' not in st.session_state:
        st.session_state.excel_processor = AdvancedExcelProcessor()
    
    # 获取当前用户会话ID
    session_id = get_session_id()
    
    # 验证session_id不为None
    if not session_id:
        print("[ERROR] 无法获取有效的session_id，使用默认值")
        import uuid
        session_id = f"user_{uuid.uuid4().hex[:16]}"
        st.session_state.user_session_id = session_id
    
    print(f"[DEBUG] 最终会话ID: {session_id}")
    
    # 页面加载时尝试从localStorage恢复配置（只在首次运行时）
    if 'localStorage_recovery_attempted' not in st.session_state:
        st.session_state.localStorage_recovery_attempted = True
        
        # 创建JavaScript来立即尝试恢复localStorage配置
        recovery_html = f"""
        <script>
            (function() {{
                const sessionId = '{session_id}';
                const key = 'ai_excel_config_' + sessionId.substring(0, 16);
                const value = localStorage.getItem(key);
                
                console.log('🔄 页面初始化localStorage恢复，会话ID:', sessionId);
                console.log('🔄 查找配置键:', key);
                
                if (value) {{
                    try {{
                        const config = JSON.parse(value);
                        console.log('🔄 发现localStorage配置，准备恢复...');
                        
                        // 将配置写入一个特殊的全局变量
                        window.initialLocalStorageConfig = {{
                            session_id: sessionId,
                            config: config,
                            restored_at: new Date().toISOString()
                        }};
                        
                        // 显示脱敏版本
                        const displayConfig = {{...config}};
                        if (displayConfig.api_key && displayConfig.api_key.length > 8) {{
                            displayConfig.api_key = config.api_key.substring(0, 4) + '****' + config.api_key.substring(config.api_key.length - 4);
                        }}
                        console.log('🔄 恢复的配置（脱敏）:', displayConfig);
                        
                        // 创建一个div来标记配置已恢复
                        const statusDiv = document.createElement('div');
                        statusDiv.id = 'localStorage_recovery_status';
                        statusDiv.style.display = 'none';
                        statusDiv.setAttribute('data-status', 'success');
                        statusDiv.setAttribute('data-session', sessionId);
                        document.body.appendChild(statusDiv);
                        
                    }} catch (e) {{
                        console.error('🔄 localStorage配置恢复失败:', e);
                        window.initialLocalStorageConfig = null;
                        
                        const statusDiv = document.createElement('div');
                        statusDiv.id = 'localStorage_recovery_status';
                        statusDiv.style.display = 'none';
                        statusDiv.setAttribute('data-status', 'error');
                        statusDiv.setAttribute('data-error', e.message);
                        document.body.appendChild(statusDiv);
                    }}
                }} else {{
                    console.log('🔄 localStorage中没有找到配置');
                    window.initialLocalStorageConfig = null;
                    
                    const statusDiv = document.createElement('div');
                    statusDiv.id = 'localStorage_recovery_status';
                    statusDiv.style.display = 'none';
                    statusDiv.setAttribute('data-status', 'not_found');
                    document.body.appendChild(statusDiv);
                }}
            }})();
        </script>
        """
        
        st.markdown(recovery_html, unsafe_allow_html=True)
        
        # 检查是否有全局配置可以恢复
        recovery_check_html = """
        <script>
            // 检查是否有恢复的配置
            setTimeout(function() {
                if (window.initialLocalStorageConfig) {
                    console.log('✅ localStorage配置已恢复到全局变量');
                    console.log('✅ 配置内容:', window.initialLocalStorageConfig);
                }
            }, 100);
        </script>
        """
        
        st.markdown(recovery_check_html, unsafe_allow_html=True)
    
    # 检查并处理localStorage恢复的配置
    if 'localStorage_config_processed' not in st.session_state:
        st.session_state.localStorage_config_processed = True
        
        # 尝试处理恢复的localStorage配置
        process_html = f"""
        <script>
            (function() {{
                if (window.initialLocalStorageConfig && window.initialLocalStorageConfig.session_id === '{session_id}') {{
                    const config = window.initialLocalStorageConfig.config;
                    console.log('🔄 处理恢复的localStorage配置...');
                    
                    // 创建一个配置处理结果的div
                    const resultDiv = document.createElement('div');
                    resultDiv.id = 'localStorage_process_result';
                    resultDiv.style.display = 'none';
                    resultDiv.setAttribute('data-processed', 'true');
                    resultDiv.setAttribute('data-config', JSON.stringify(config));
                    resultDiv.setAttribute('data-session', '{session_id}');
                    document.body.appendChild(resultDiv);
                    
                    console.log('🔄 localStorage配置已标记为待处理');
                }} else {{
                    console.log('🔄 没有localStorage配置需要处理');
                }}
            }})();
        </script>
        """
        
        st.markdown(process_html, unsafe_allow_html=True)
    
    # 初始化localStorage配置恢复（只在首次加载时）
    if 'localStorage_initialized' not in st.session_state:
        init_localStorage_config(session_id)
        st.session_state.localStorage_initialized = True
    
    # 初始化配置加载标记
    if 'config_loaded' not in st.session_state:
        st.session_state.config_loaded = False
    
    # 加载用户配置（只在第一次或明确需要时加载）
    if not st.session_state.config_loaded:
        loaded_config = load_user_config(config_manager, session_id)
        st.session_state.config_loaded = True
        if loaded_config:
            st.session_state.config_load_success = True
        else:
            st.session_state.config_load_success = False
    
    # 显示主标题
    st.markdown('<h1 class="main-header">🚀 AI Excel 智能分析工具 - 多用户版</h1>', unsafe_allow_html=True)
    
    # 显示会话信息
    st.markdown(f'<div class="session-info">👤 当前会话: {session_id[:20]}... | 🔐 数据隔离保护</div>', unsafe_allow_html=True)
    
    # 显示配置加载状态
    if st.session_state.get('config_load_success', False):
        config_source = st.session_state.get('config_source', '未知来源')
        if config_source == 'localStorage':
            st.markdown(f'<div class="config-loaded">✅ 已从浏览器localStorage恢复配置 🔄</div>', unsafe_allow_html=True)
        elif config_source == '服务器端完整配置':
            st.markdown(f'<div class="config-loaded">✅ 已加载服务器端保存配置 🗄️</div>', unsafe_allow_html=True)
        elif config_source == '服务器端浏览器缓存':
            st.markdown(f'<div class="config-loaded">✅ 已加载浏览器缓存配置 📱</div>', unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="config-loaded">✅ 已加载配置 ({config_source})</div>', unsafe_allow_html=True)
    
    # 侧边栏配置
    with st.sidebar:
        st.header("⚙️ 配置设置")
        
        # OpenAI配置
        st.subheader("🤖 OpenAI API 配置")
        
        # 使用保存的配置作为默认值
        default_api_key = st.session_state.get('saved_api_key', '')
        default_base_url = st.session_state.get('saved_base_url', 'https://apistudy.mycache.cn/v1')
        default_model = st.session_state.get('saved_model', 'deepseek-v3')
        
        api_key = st.text_input(
            "🔑 API Key", 
            value=default_api_key,
            type="password",
            help="输入您的OpenAI API密钥，会自动保存",
            key="api_key_input"
        )
        
        base_url = st.text_input(
            "🌐 Base URL (可选)", 
            value=default_base_url,
            help="自定义API服务地址，留空使用默认",
            key="base_url_input"
        )
        
        model_options = ["deepseek-v3","gpt-4.1-mini", "gpt-4.1"]
        selected_model = st.selectbox(
            "🧠 选择模型", 
            model_options,
            index=model_options.index(default_model) if default_model in model_options else 0,
            key="model_select"
        )
        
        # 检测配置变化并自动保存
        current_config = {
            'api_key': api_key,
            'base_url': base_url,
            'selected_model': selected_model
        }
        
        # 使用唯一的key来避免重复保存
        config_key = f"{api_key}_{base_url}_{selected_model}"
        
        # 检查是否有配置变化
        config_changed = False
        if 'last_config_key' not in st.session_state:
            st.session_state.last_config_key = config_key
            config_changed = False  # 初次加载不触发保存
        elif st.session_state.last_config_key != config_key:
            config_changed = True
        
        # 如果配置有变化且API Key不为空，自动保存
        if config_changed and api_key.strip():
            try:
                success = auto_save_config(config_manager, session_id, api_key, base_url, selected_model)
                if success:
                    st.session_state.last_config_key = config_key
                    
                    # 总是保存到浏览器缓存
                    save_to_browser_cache(current_config, config_manager, session_id)
                    
                    # 显示自动保存提示
                    st.success("✅ 配置已自动保存")
                else:
                    st.error("❌ 自动保存失败")
            except Exception as e:
                st.error(f"❌ 自动保存出错: {str(e)}")
        
        # 显示当前配置状态
        if api_key.strip():
            st.info(f"🔑 API Key: {api_key[:8]}...")
            st.info(f"🌐 Base URL: {base_url}")
            st.info(f"🧠 模型: {selected_model}")
        
        # 配置保存和缓存选项
        st.subheader("💾 配置管理")
        
        col_save, col_clear = st.columns(2)
        
        with col_save:
            if st.button("💾 手动保存配置", use_container_width=True):
                config_to_save = {
                    'api_key': api_key,
                    'base_url': base_url,
                    'selected_model': selected_model,
                    'save_timestamp': datetime.now().isoformat(),
                    'manual_save': True
                }
                
                try:
                    if save_user_config(config_manager, session_id, config_to_save):
                        st.success("✅ 配置已手动保存到服务器")
                        # 同时保存到session state
                        st.session_state.saved_api_key = api_key
                        st.session_state.saved_base_url = base_url
                        st.session_state.saved_model = selected_model
                        st.session_state.last_config_key = config_key
                        
                        # 总是保存到浏览器缓存
                        save_to_browser_cache(config_to_save, config_manager, session_id)
                        st.success("✅ 浏览器缓存已更新")
                    else:
                        st.error("❌ 配置保存失败")
                except Exception as e:
                    st.error(f"❌ 配置保存出错: {str(e)}")
        
        with col_clear:
            if st.button("🗑️ 清除配置", use_container_width=True):
                try:
                    # 清除保存的配置
                    if session_manager.get_user_workspace(session_id):
                        config_file = session_manager.get_user_workspace(session_id) / "user_config.json"
                        if config_file.exists():
                            config_file.unlink()
                            st.success("✅ 服务器端配置已清除")
                        
                        # 清除浏览器缓存文件
                        cache_file = session_manager.get_user_workspace(session_id) / "browser_cache.json"
                        if cache_file.exists():
                            cache_file.unlink()
                            st.success("✅ 浏览器缓存文件已清除")
                    
                    # 清除浏览器localStorage
                    storage_key = f"ai_excel_config_{session_id[:16]}"
                    remove_browser_storage_item(storage_key)
                    st.success("✅ 浏览器localStorage已清除")
                    
                    # 清除session state
                    for key in ['saved_api_key', 'saved_base_url', 'saved_model', 'browser_cached_config', 'last_config_key', 'config_loaded', 'config_load_success']:
                        if key in st.session_state:
                            del st.session_state[key]
                    
                    st.success("✅ 所有配置已清除，页面将刷新")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ 清除配置出错: {str(e)}")
        
        # 用户数据统计
        st.subheader("📊 我的数据统计")
        try:
            # 获取用户工作空间信息
            user_workspace = session_manager.get_user_workspace(session_id)
            if user_workspace and user_workspace.exists():
                # 统计用户文件
                user_files = list(user_workspace.glob("*"))
                user_file_count = len([f for f in user_files if f.is_file() and not f.name.startswith('.')])
                
                # 计算用户磁盘使用
                user_disk_usage = sum(f.stat().st_size for f in user_files if f.is_file()) / (1024 * 1024)  # MB
                
                # 显示用户统计
                col_user1, col_user2 = st.columns(2)
                with col_user1:
                    st.metric("我的文件数", user_file_count)
                with col_user2:
                    st.metric("我的磁盘使用", f"{user_disk_usage:.2f} MB")
                
                # 全局系统状态（可选显示）
                show_global_stats = st.checkbox("显示全局系统状态", value=False)
                if show_global_stats:
                    global_stats = session_manager.get_session_stats()
                    st.write("**全局系统状态：**")
                    col_g1, col_g2, col_g3 = st.columns(3)
                    with col_g1:
                        st.metric("总活跃会话", global_stats['active_sessions'])
                    with col_g2:
                        st.metric("总文件数", global_stats['total_files'])
                    with col_g3:
                        st.metric("总磁盘使用", f"{global_stats['disk_usage_mb']} MB")
            else:
                st.info("还没有上传任何文件")
                
        except Exception as e:
            st.error(f"统计信息获取失败: {str(e)}")
        
        # 手动清理按钮（仅清理当前用户）
        if st.button("🧹 清理我的数据", use_container_width=True):
            if session_manager.cleanup_user_session(session_id):
                st.success("✅ 您的数据已清理完成")
                st.rerun()
            else:
                st.error("❌ 清理失败")
    
    # 初始化Excel处理器和会话状态
    if 'excel_processor' not in st.session_state:
        st.session_state.excel_processor = AdvancedExcelProcessor()
    if 'excel_data' not in st.session_state:
        st.session_state.excel_data = {}
    if 'excel_analysis' not in st.session_state:
        st.session_state.excel_analysis = ""
    if 'chat_history' not in st.session_state:
        st.session_state.chat_history = []
    if 'current_sheet' not in st.session_state:
        st.session_state.current_sheet = None
    
    # 初始化文档处理器和会话状态
    if 'document_processor' not in st.session_state:
        try:
            from document_utils import AdvancedDocumentProcessor
            from document_analyzer import DocumentAnalyzer
            
            # 检查依赖
            analyzer = DocumentAnalyzer()
            missing_deps = analyzer.get_missing_dependencies()
            
            if missing_deps:
                st.session_state.document_processor = None
                st.session_state.document_processor_error = f"缺少依赖: {', '.join(missing_deps)}"
            else:
                st.session_state.document_processor = AdvancedDocumentProcessor()
                st.session_state.document_processor_error = None
                
        except ImportError as e:
            st.session_state.document_processor = None
            st.session_state.document_processor_error = f"导入错误: {str(e)}"
    if 'document_data' not in st.session_state:
        st.session_state.document_data = {}
    if 'document_analysis' not in st.session_state:
        st.session_state.document_analysis = ""
    if 'doc_chat_history' not in st.session_state:
        st.session_state.doc_chat_history = []
    
    # 文件上传 - 支持Excel和文档
    st.subheader("📁 文件上传")
    
    # 选择分析模式
    analysis_mode = st.radio(
        "🔧 选择分析模式",
        ["📊 Excel分析", "📄 文档分析"],
        horizontal=True,
        key="analysis_mode_selector"
    )
    
    # 初始化uploaded_file变量
    uploaded_file = None
    
    if analysis_mode == "📊 Excel分析":
        # Excel分析模式
        st.markdown("### 📊 Excel文件分析")
        
        # 获取用户已有的Excel文件
        existing_excel_files = session_manager.get_user_excel_files(session_id)
    
    else:
        # 文档分析模式
        st.markdown("### 📄 文档分析 (DOCX/PDF)")
        
        # 检查文档处理器是否可用
        if st.session_state.document_processor is None:
            error_msg = getattr(st.session_state, 'document_processor_error', '未知错误')
            
            st.error(f"❌ 文档分析功能不可用: {error_msg}")
            
            with st.expander("🔧 解决方案", expanded=True):
                st.markdown("""
                **缺少文档分析依赖库，请按以下步骤安装：**
                
                **方式一：自动安装（推荐）**
                ```bash
                python install_document_dependencies.py
                ```
                
                **方式二：手动安装**
                ```bash
                pip install markitdown[all] python-docx pymupdf PyPDF2 pdfplumber
                ```
                
                **方式三：一键安装最新requirements**
                ```bash
                pip install -r requirements.txt
                ```
                
                **常见问题：**
                - `markitdown[all]`: 包含PDF、DOCX等所有格式支持
                - 如果安装失败，可能需要升级pip: `pip install --upgrade pip`
                - Windows用户可能需要安装Visual C++构建工具
                """)
                
                if st.button("🔄 重新检查依赖", type="primary"):
                    # 清除缓存，重新初始化
                    if 'document_processor' in st.session_state:
                        del st.session_state.document_processor
                    if 'document_processor_error' in st.session_state:
                        del st.session_state.document_processor_error
                    st.rerun()
            
            return  # 提前退出，不显示后续界面
        
        # 获取用户已有的文档文件
        existing_doc_files = session_manager.get_user_files(session_id, extensions=['.docx', '.pdf'])
    
    # 文件选择方式
    if analysis_mode == "📊 Excel分析" and existing_excel_files:
        file_option = st.radio(
            "📂 选择文件方式",
            ["选择已有文件", "上传新文件"],
            key="file_option_radio"
        )
        
        if file_option == "选择已有文件":
            # 显示已有文件选择器
            st.markdown("**📋 您已上传的Excel文件：**")
            
            # 创建文件选择选项
            file_options = []
            file_details = {}
            
            for file_info in existing_excel_files:
                display_text = f"{file_info['display_name']} ({file_info['size_mb']} MB, {file_info['modified_time'].strftime('%Y-%m-%d %H:%M')})"
                file_options.append(display_text)
                file_details[display_text] = file_info
            
            selected_file_text = st.selectbox(
                "选择要分析的Excel文件",
                file_options,
                key="existing_file_selector"
            )
            
            if selected_file_text and st.button("📊 加载选择的文件", type="primary"):
                try:
                    from pathlib import Path  # 确保Path可用
                    selected_file_info = file_details[selected_file_text]
                    file_path = Path(selected_file_info['path'])
                    
                    with st.spinner("📤 正在加载已有文件..."):
                        # 加载Excel数据
                        excel_data = st.session_state.excel_processor.load_excel(str(file_path))
                        st.session_state.excel_data = excel_data
                        
                        sheet_names = list(excel_data.keys())
                        if sheet_names:
                            st.session_state.current_sheet = sheet_names[0]
                        
                        # 保存当前文件信息到session state
                        st.session_state.current_file_path = str(file_path)
                        st.session_state.current_file_name = selected_file_info['display_name']
                        
                        st.success(f"✅ 文件加载成功！文件: {selected_file_info['display_name']}")
                        st.rerun()
                        
                except Exception as e:
                    st.error(f"❌ 文件加载错误: {str(e)}")
        else:
            # 上传新文件
            uploaded_file = st.file_uploader(
                "选择Excel文件",
                type=['xlsx', 'xls'],
                help="支持.xlsx和.xls格式文件",
                key="new_file_uploader"
            )
    elif analysis_mode == "📄 文档分析" and existing_doc_files:
        # 文档分析模式 - 有已有文件
        file_option = st.radio(
            "📂 选择文件方式",
            ["选择已有文档", "上传新文档"],
            key="doc_file_option_radio"
        )
        
        if file_option == "选择已有文档":
            # 显示已有文档选择器
            st.markdown("**📋 您已上传的文档文件：**")
            
            # 创建文件选择选项
            file_options = []
            file_details = {}
            
            for file_info in existing_doc_files:
                display_text = f"{file_info['display_name']} ({file_info['size_mb']} MB, {file_info['modified_time'].strftime('%Y-%m-%d %H:%M')})"
                file_options.append(display_text)
                file_details[display_text] = file_info
            
            selected_file_text = st.selectbox(
                "选择要分析的文档文件",
                file_options,
                key="existing_doc_selector"
            )
            
            if selected_file_text and st.button("📄 加载选择的文档", type="primary"):
                try:
                    from pathlib import Path
                    selected_file_info = file_details[selected_file_text]
                    file_path = Path(selected_file_info['path'])
                    
                    with st.spinner("📤 正在加载已有文档..."):
                        # 加载文档数据
                        if st.session_state.document_processor:
                            document_data = st.session_state.document_processor.load_document(str(file_path))
                            st.session_state.document_data = document_data
                            
                            # 保存当前文件信息到session state
                            st.session_state.current_doc_path = str(file_path)
                            st.session_state.current_doc_name = selected_file_info['display_name']
                            
                            st.success(f"✅ 文档加载成功！文件: {selected_file_info['display_name']}")
                            st.rerun()
                        else:
                            st.error("❌ 文档处理器未初始化")
                        
                except Exception as e:
                    st.error(f"❌ 文档加载错误: {str(e)}")
        else:
            # 上传新文档
            uploaded_file = st.file_uploader(
                "选择文档文件",
                type=['docx', 'pdf'],
                help="支持.docx和.pdf格式文件",
                key="new_doc_uploader"
            )
    else:
        # 没有已有文件，直接显示上传
        if analysis_mode == "📊 Excel分析":
            st.info("💡 您还没有上传过Excel文件，请上传您的第一个文件")
            uploaded_file = st.file_uploader(
                "选择Excel文件",
                type=['xlsx', 'xls'],
                help="支持.xlsx和.xls格式文件",
                key="first_file_uploader"
            )
        else:
            st.info("💡 您还没有上传过文档文件，请上传您的第一个文档")
            uploaded_file = st.file_uploader(
                "选择文档文件",
                type=['docx', 'pdf'],
                help="支持.docx和.pdf格式文件",
                key="first_doc_uploader"
            )
    
    # 处理文件上传
    if uploaded_file is not None:
        # 检查是否已经处理过这个文件（避免重复上传）
        if not hasattr(st.session_state, 'last_uploaded_file') or st.session_state.last_uploaded_file != uploaded_file.name:
            try:
                with st.spinner("📤 正在上传和处理文件..."):
                    # 使用会话管理器保存文件
                    file_path = session_manager.save_uploaded_file(
                        session_id, 
                        uploaded_file, 
                        uploaded_file.name
                    )
                    
                    if analysis_mode == "📊 Excel分析":
                        # 加载Excel数据
                        excel_data = st.session_state.excel_processor.load_excel(str(file_path))
                        st.session_state.excel_data = excel_data
                        
                        sheet_names = list(excel_data.keys())
                        if sheet_names:
                            st.session_state.current_sheet = sheet_names[0]
                        
                        # 保存当前文件信息到session state
                        st.session_state.current_file_path = str(file_path)
                        st.session_state.current_file_name = uploaded_file.name
                        
                    else:
                        # 加载文档数据
                        if st.session_state.document_processor:
                            document_data = st.session_state.document_processor.load_document(str(file_path))
                            st.session_state.document_data = document_data
                            
                            # 保存当前文件信息到session state
                            st.session_state.current_doc_path = str(file_path)
                            st.session_state.current_doc_name = uploaded_file.name
                        else:
                            st.error("❌ 文档处理器未初始化")
                            return
                    
                    st.session_state.last_uploaded_file = uploaded_file.name  # 记录已处理的文件
                    
                    st.success(f"✅ 文件上传成功！保存位置: {file_path.name}")
                    st.rerun()
                    
            except Exception as e:
                st.error(f"❌ 文件处理错误: {str(e)}")
        else:
            # 文件已经处理过，显示当前状态
            st.info(f"📁 当前文件: {uploaded_file.name}")
    
    # 主要界面：根据分析模式显示不同的Tabs
    if analysis_mode == "📊 Excel分析" and st.session_state.excel_data:
        tab1, tab2, tab3, tab4 = st.tabs(["📋 数据预览与管理", "🤖 AI 智能分析", "💻 代码执行", "🛠️ 数据工具"])
    elif analysis_mode == "📄 文档分析" and st.session_state.document_data:
        doc_tab1, doc_tab2, doc_tab3, doc_tab4 = st.tabs(["📄 文档预览", "🤖 AI 文档分析", "💻 代码执行", "🔍 搜索工具"])
    
    # Excel分析界面
    if analysis_mode == "📊 Excel分析" and st.session_state.excel_data:
        
        # Tab 1: 数据预览与管理
        with tab1:
            st.header("📋 Excel 数据预览与管理")
            
            sheet_names = list(st.session_state.excel_data.keys())
            st.success(f"✅ 成功载入 {len(sheet_names)} 个工作表")
            
            # 工作表选择器
            if len(sheet_names) > 1:
                selected_sheet = st.selectbox("🗂️ 选择工作表", sheet_names, key="sheet_selector")
            else:
                selected_sheet = sheet_names[0]
            
            st.session_state.current_sheet = selected_sheet
            
            # 显示当前工作表预览
            if selected_sheet in st.session_state.excel_data:
                df = st.session_state.excel_data[selected_sheet]
                
                # 数据统计概览 - 优化版本
                missing_count = df.isnull().sum().sum()
                duplicates = DataAnalyzer.find_duplicates(df)
                
                with st.expander(f"📈 数据统计概览 (数据质量: {len(duplicates)} 重复行, {missing_count} 缺失值)", expanded=True):
                    # 基础统计卡片
                    col_a, col_b, col_c, col_d = st.columns(4)
                    with col_a:
                        st.markdown(f'<div class="metric-card"><h3>{len(df)}</h3><p>数据行数</p></div>', unsafe_allow_html=True)
                    with col_b:
                        st.markdown(f'<div class="metric-card"><h3>{len(df.columns)}</h3><p>数据列数</p></div>', unsafe_allow_html=True)
                    with col_c:
                        st.markdown(f'<div class="metric-card"><h3>{missing_count}</h3><p>缺失值</p></div>', unsafe_allow_html=True)
                    with col_d:
                        st.markdown(f'<div class="metric-card"><h3>{len(duplicates)}</h3><p>重复行</p></div>', unsafe_allow_html=True)
                    
                    # 详细数据质量报告
                    if missing_count > 0 or len(duplicates) > 0:
                        st.markdown("### 📋 详细数据质量报告")
                        
                        if missing_count > 0:
                            st.markdown("**🔍 缺失值分布:**")
                            missing_series = df.isnull().sum()
                            missing_df = pd.DataFrame({
                                '列名': missing_series.index,
                                '缺失数量': missing_series.values
                            })
                            missing_df = missing_df[missing_df['缺失数量'] > 0]
                            missing_df['缺失比例(%)'] = (missing_df['缺失数量'] / len(df) * 100).round(2)
                            st.dataframe(missing_df, use_container_width=True)
                        
                        if len(duplicates) > 0:
                            st.markdown(f"**🔄 重复行信息:** 发现 {len(duplicates)} 个重复行")
                            if st.button("👁️ 查看重复行", key="view_duplicates"):
                                st.dataframe(duplicates.head(10), use_container_width=True)
                    
                    # 数据类型统计
                    st.markdown("### 📊 数据类型分布")
                    dtype_counts = df.dtypes.value_counts()
                    dtype_df = pd.DataFrame({
                        '数据类型': dtype_counts.index.astype(str),
                        '列数': dtype_counts.values
                    })
                    
                    col_dtype_table, col_dtype_chart = st.columns([1, 1])
                    with col_dtype_table:
                        st.dataframe(dtype_df, use_container_width=True)
                    with col_dtype_chart:
                        # 简单的数据类型分布图
                        st.bar_chart(dtype_df.set_index('数据类型'))
                
                # 数据预览 - 优化版本
                df_shape = df.shape
                total_cells = df_shape[0] * df_shape[1]
                
                with st.expander(f"📊 数据预览 ({df_shape[0]} 行 × {df_shape[1]} 列，共 {total_cells:,} 个单元格)", expanded=True):
                    st.info("💡 此预览仅用于查看和AI理解，实际代码执行请使用'代码执行'标签页")
                    
                    # 显示数据表格 - 添加错误处理
                    try:
                        # 尝试显示前20行数据
                        preview_df = df.head(20).copy()
                        
                        # 确保所有列的数据类型一致，避免pyarrow错误
                        for col in preview_df.columns:
                            if preview_df[col].dtype == 'object':
                                # 将混合类型的object列转换为字符串
                                preview_df[col] = preview_df[col].astype(str)
                        
                        st.dataframe(preview_df, use_container_width=True)
                        
                        # 添加数据导出功能
                        col_export_preview, col_export_full = st.columns(2)
                        
                        with col_export_preview:
                            if st.button("📥 导出预览数据", key=f"export_preview_{st.session_state.current_sheet}"):
                                st.download_button(
                                    label="💾 下载预览数据(Excel)",
                                    data=preview_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig'),
                                    file_name=f"{st.session_state.current_sheet}_预览_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                                    mime="text/csv",
                                    key=f"download_preview_{st.session_state.current_sheet}"
                                )
                        
                        with col_export_full:
                            if st.button("📥 导出完整数据", key=f"export_full_{st.session_state.current_sheet}"):
                                st.download_button(
                                    label="💾 下载完整数据(Excel)",
                                    data=df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig'),
                                    file_name=f"{st.session_state.current_sheet}_完整_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                                    mime="text/csv",
                                    key=f"download_full_{st.session_state.current_sheet}"
                                )
                        
                    except Exception as e:
                        st.warning(f"⚠️ 数据预览显示出现问题，使用文本格式展示")
                        st.write(f"**错误信息**: {str(e)}")
                        
                        # 回退到文本格式显示
                        try:
                            preview_text = df.head(10).to_string()
                            st.text(preview_text)
                        except Exception as e2:
                            st.error(f"❌ 无法显示数据预览: {str(e2)}")
                            st.write("**数据基本信息:**")
                            st.write(f"- 行数: {len(df)}")
                            st.write(f"- 列数: {len(df.columns)}")
                            st.write(f"- 列名: {list(df.columns)}")
        
        # Tab 2: AI智能分析（保持原有功能）
        with tab2:
            st.header("🤖 AI 智能分析")
            
            # 轻量级Excel结构分析（无需API）
            st.subheader("📋 Excel文件结构分析")
            st.info("💡 即使没有配置AI API，您也可以获得Excel文件的结构分析")
            
            # 导入轻量级分析器
            try:
                from ai_tab_analyzer import AITabAnalyzer
            except ImportError:
                st.error("❌ 无法导入AI分析器，请确保ai_tab_analyzer.py文件存在")
                AITabAnalyzer = None
            
            # 添加分析按钮和结果显示
            col_quick_analyze, col_clear_analysis = st.columns([3, 1])
            
            with col_quick_analyze:
                if st.button("🔍 快速分析Excel结构", type="secondary", use_container_width=True):
                    if AITabAnalyzer is None:
                        st.error("❌ AI分析器不可用")
                    elif hasattr(st.session_state, 'current_file_path') and st.session_state.current_file_path:
                        try:
                            with st.spinner("📊 正在分析Excel文件结构..."):
                                analyzer = AITabAnalyzer()
                                analysis_result = analyzer.analyze_for_ai(st.session_state.current_file_path)
                                st.session_state.quick_excel_analysis = analysis_result
                                st.success("✅ Excel智能结构分析完成！")
                                st.rerun()
                        except Exception as e:
                            st.error(f"❌ 结构分析失败: {str(e)}")
                    else:
                        st.warning("⚠️ 请先上传Excel文件")
            
            with col_clear_analysis:
                if st.button("🗑️ 清除分析", use_container_width=True):
                    if 'quick_excel_analysis' in st.session_state:
                        del st.session_state.quick_excel_analysis
                        st.rerun()
            
            # 显示快速分析结果
            if 'quick_excel_analysis' in st.session_state and st.session_state.quick_excel_analysis:
                st.subheader("📊 Excel结构分析结果")
                
                # 计算统计信息
                analysis_content = st.session_state.quick_excel_analysis
                char_count = len(analysis_content)
                word_count = len(analysis_content.split())
                line_count = analysis_content.count('\n') + 1
                
                # 使用折叠框显示，标题包含统计信息
                with st.expander(f"📋 查看详细分析 (共 {word_count} 词，{char_count} 字符，{line_count} 行)", expanded=True):
                     # 使用统一的容器渲染函数
                     render_content_container(analysis_content, 'excel-structure')
                     
                     # 添加操作按钮
                     col_download, col_copy, col_refresh = st.columns(3)
                    
                     with col_download:
                         if st.button("📥 下载分析结果", key="download_excel_structure"):
                             # 生成下载内容
                             download_content = f"""# Excel结构分析结果
                            
                             文件名: {getattr(st.session_state, 'current_file_name', '未知文件')}
                             生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

{analysis_content}
"""
                             st.download_button(
                                 label="💾 下载为.md文件",
                                 data=download_content,
                                 file_name=f"Excel结构分析_{datetime.now().strftime('%Y%m%d_%H%M%S')}.md",
                                 mime="text/markdown",
                                 key="download_excel_structure_md"
                             )
                    
                     with col_copy:
                         if st.button("📋 复制到剪贴板", key="copy_excel_structure"):
                             st.success("✅ 内容已复制到剪贴板")
                             # 使用JavaScript复制到剪贴板
                             copy_js = f"""
                             <script>
                             navigator.clipboard.writeText(`{analysis_content.replace('`', '\\`')}`);
                             </script>
"""
                             st.markdown(copy_js, unsafe_allow_html=True)
                    
                     with col_refresh:
                         if st.button("🔄 重新分析结构", key="refresh_excel_structure"):
                            if 'quick_excel_analysis' in st.session_state:
                                del st.session_state.quick_excel_analysis
                                st.rerun()
                
                # 功能说明和提示
                st.info("📝 **智能分析说明**：\n"
                       "- 🟢 **标准二维表格**：直接列出字段和筛选项\n"
                       "- 🟡 **复杂表格**：智能处理合并单元格，识别字段递进关系\n"
                       "- 🏷️ **自动筛选项识别**：≤10个唯一值的字段显示全部可选值\n"
                       "- 📈 **F列后字段关系**：提供横向字段递进说明，便于AI理解")
                
                # 如果有API配置，提供将分析结果作为AI分析基础的选项
                if api_key:
                    st.success("💡 **AI分析提示**：上述结构分析将自动作为深度AI分析的基础信息，提高AI理解准确性！")
            
            # 分隔线
            st.markdown("---")
            
            # 原有的AI分析功能
            st.subheader("🧠 深度AI分析")
            
            if not api_key:
                st.warning("⚠️ 请在侧边栏配置OpenAI API Key以使用深度AI分析功能")
            else:
                # 初始化AI分析器
                ai_analyzer = EnhancedAIAnalyzer(api_key, base_url, selected_model)
                
                # 显示已有的深度分析结果
                if 'excel_analysis' in st.session_state and st.session_state.excel_analysis:
                    # 计算分析结果的统计信息
                    analysis_content = st.session_state.excel_analysis
                    char_count = len(analysis_content)
                    word_count = len(analysis_content.split())
                    line_count = analysis_content.count('\n') + 1
                    
                    # 使用折叠框显示深度分析结果
                    with st.expander(f"🎯 AI深度分析结果 (共 {word_count} 词，{char_count} 字符，{line_count} 行)", expanded=True):
                        # 使用统一的容器渲染函数
                        render_content_container(analysis_content, 'excel-ai')
                        
                        # 添加操作按钮
                        col_download_ai, col_copy_ai, col_refresh_ai = st.columns(3)
                        
                        with col_download_ai:
                            if st.button("📥 下载AI分析", key="download_ai_analysis"):
                                # 生成下载内容
                                download_content = f"""# Excel AI深度分析报告
                                
文件名: {getattr(st.session_state, 'current_file_name', '未知文件')}
分析时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
模型: {selected_model}

{analysis_content}
"""
                                st.download_button(
                                    label="💾 下载为.md文件",
                                    data=download_content,
                                    file_name=f"Excel_AI分析_{datetime.now().strftime('%Y%m%d_%H%M%S')}.md",
                                    mime="text/markdown",
                                    key="download_ai_analysis_md"
                                )
                        
                        with col_copy_ai:
                            if st.button("📋 复制分析结果", key="copy_ai_analysis"):
                                st.success("✅ AI分析结果已复制到剪贴板")
                                # 使用JavaScript复制到剪贴板
                                copy_js = f"""
                                <script>
                                navigator.clipboard.writeText(`{analysis_content.replace('`', '\\`')}`);
                                </script>
                                """
                                st.markdown(copy_js, unsafe_allow_html=True)
                        
                        with col_refresh_ai:
                            if st.button("🔄 重新生成分析", key="refresh_ai_analysis"):
                                st.session_state.excel_analysis = ""
                                st.rerun()
                
                # AI分析控制
                col_analyze, col_refresh = st.columns([3, 1])
                
                with col_analyze:
                    if st.button("🔍 开始AI深度分析", type="primary", use_container_width=True):
                        with st.spinner("🧠 AI正在深度分析您的数据..."):
                            # 获取Excel结构分析结果
                            structure_info = ""
                            if 'quick_excel_analysis' in st.session_state and st.session_state.quick_excel_analysis:
                                structure_info = st.session_state.quick_excel_analysis
                            
                            # 进行AI深度分析（已包含数据内容和特征）
                            analysis = ai_analyzer.analyze_excel_structure(st.session_state.excel_data)
                            
                            # 构建完整的分析报告，将结构信息与业务分析结合
                            if structure_info:
                                combined_analysis = f"""## 📋 Excel文件结构解析

{structure_info}

---

## 🎯 AI深度业务分析

{analysis}"""
                            else:
                                combined_analysis = analysis
                            
                            st.session_state.excel_analysis = combined_analysis
                            st.session_state.chat_history.append({
                                "role": "assistant",
                                "content": f"**📋 Excel深度分析报告**\n\n{combined_analysis}"
                            })
                            st.rerun()
                
                with col_refresh:
                    if st.button("🔄 重新分析", use_container_width=True):
                        st.session_state.excel_analysis = ""
                        st.session_state.chat_history = []
                        st.rerun()
                
                # 快速操作按钮
                st.subheader("⚡ 智能业务分析")
                col_quick1, col_quick2 = st.columns(2)
                
                quick_actions = [
                    ("🎯 业务场景识别", "请分析这些数据的业务场景和用途，识别核心业务主题和流程"),
                    ("🔗 数据关系分析", "请分析不同工作表之间的业务逻辑关系，识别关键字段和数据流向"),
                    ("💎 关键指标发现", "请识别数据中的核心业务指标和度量字段，分析它们的业务价值"),
                    ("📊 分析机会推荐", "基于数据特征，推荐具体的分析方向和业务洞察机会")
                ]
                
                for i, (title, prompt) in enumerate(quick_actions):
                    col = col_quick1 if i % 2 == 0 else col_quick2
                    with col:
                        if st.button(title, use_container_width=True, key=f"quick_{i}"):
                            st.session_state.chat_history.append({
                                "role": "user",
                                "content": prompt
                            })
                            
                            with st.spinner("AI正在分析..."):
                                response = ai_analyzer.chat_with_data(
                                    prompt,
                                    st.session_state.excel_data,
                                    st.session_state.excel_analysis
                                )
                                st.session_state.chat_history.append({
                                    "role": "assistant", 
                                    "content": response
                                })
                            st.rerun()
                
                # 聊天历史显示 - 优化版本
                if st.session_state.chat_history:
                    # 计算对话历史统计信息
                    total_conversations = len(st.session_state.chat_history)
                    user_messages = len([chat for chat in st.session_state.chat_history if chat["role"] == "user"])
                    ai_messages = len([chat for chat in st.session_state.chat_history if chat["role"] == "assistant"])
                    
                    with st.expander(f"💬 AI对话历史 (共 {total_conversations} 条消息: {user_messages} 个问题, {ai_messages} 个回答)", expanded=False):
                        # 使用统一的对话容器渲染函数
                        render_chat_container(st.session_state.chat_history, 'excel-chat')
                        
                        # 对话历史操作按钮
                        col_export_chat, col_clear_chat = st.columns(2)
                        
                        with col_export_chat:
                            if st.button("📥 导出对话历史", key="export_chat_history"):
                                # 生成对话历史文本
                                chat_text = f"""# Excel AI对话历史
                                
文件名: {getattr(st.session_state, 'current_file_name', '未知文件')}
导出时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
总对话数: {total_conversations} 条

"""
                                for i, chat in enumerate(st.session_state.chat_history, 1):
                                    role_name = "用户" if chat["role"] == "user" else "AI助手"
                                    chat_text += f"""
## {i}. {role_name}

{chat["content"]}

"""
                                
                                st.download_button(
                                    label="💾 下载为.md文件",
                                    data=chat_text,
                                    file_name=f"Excel对话历史_{datetime.now().strftime('%Y%m%d_%H%M%S')}.md",
                                    mime="text/markdown",
                                    key="download_chat_history_md"
                                )
                        
                        with col_clear_chat:
                            if st.button("🗑️ 清空对话历史", key="clear_chat_history"):
                                st.session_state.chat_history = []
                                st.rerun()
                
                else:
                    st.info("💬 还没有AI对话记录，请开始提问或使用快速分析功能")
                
                # 用户输入
                st.subheader("💭 向AI提问")
                user_input = st.text_area(
                    "输入您的问题",
                    placeholder="例如：分析销售趋势、查找数据异常、提供业务建议等...",
                    height=80,
                    key="ai_chat_input"
                )
                
                col_send, col_clear = st.columns([1, 1])
                
                with col_send:
                    if st.button("📤 发送", type="primary", use_container_width=True):
                        if user_input.strip():
                            st.session_state.chat_history.append({
                                "role": "user",
                                "content": user_input
                            })
                            
                            with st.spinner("🤔 AI正在思考..."):
                                response = ai_analyzer.chat_with_data(
                                    user_input,
                                    st.session_state.excel_data,
                                    st.session_state.excel_analysis
                                )
                                st.session_state.chat_history.append({
                                    "role": "assistant",
                                    "content": response
                                })
                            st.rerun()
                
                with col_clear:
                    if st.button("🗑️ 清空输入", use_container_width=True):
                        # 清空输入框（通过重置key）
                        if 'ai_chat_input' in st.session_state:
                            del st.session_state.ai_chat_input
                        st.rerun()
        
        # Tab 3: 代码执行（简化版）
        with tab3:
            st.header("💻 Excel 代码执行环境")
            st.info("🔐 您的代码在隔离环境中执行，数据完全私有")
            
            if st.session_state.current_sheet:
                # 显示可用的Excel数据和文件信息
                st.subheader("📊 可用的Excel数据和文件")
                
                # 数据变量信息
                col_info1, col_info2 = st.columns(2)
                
                with col_info1:
                    st.markdown("**📋 可用的DataFrame变量:**")
                    for sheet_name in st.session_state.excel_data.keys():
                        safe_name = sheet_name.replace(' ', '_').replace('-', '_').replace('.', '_')
                        df_shape = st.session_state.excel_data[sheet_name].shape
                        st.code(f"df_{safe_name}  # {sheet_name} ({df_shape[0]}行×{df_shape[1]}列)")
                    
                    st.markdown("**📁 原始Excel文件访问:**")
                    if hasattr(st.session_state, 'current_file_name') and st.session_state.current_file_name:
                        st.code(f"excel_file_path  # 原始Excel文件路径")
                        st.code(f"excel_file_name  # 文件名: {st.session_state.current_file_name}")
                    else:
                        st.info("选择或上传文件后可获得文件路径变量")
                
                with col_info2:
                    st.markdown("**🔧 可用的库:**")
                    st.code("pd  # pandas\nnp  # numpy\npx  # plotly.express\ngo  # plotly.graph_objects\nos  # 文件操作")
                    
                    st.markdown("**📊 工作表关系信息:**")
                    st.code(f"sheet_names  # 所有工作表名称列表\nsheet_info  # 工作表详细信息字典")
                
                # 代码编辑器
                st.subheader("🖥️ Python代码编辑器")
                
                # 默认代码模板 - 包含Excel文件操作
                current_safe_name = st.session_state.current_sheet.replace(' ', '_').replace('-', '_').replace('.', '_')
                default_code = f"""# Excel文件和数据处理代码 - 多用户环境
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import os
from pathlib import Path

# ===========================================
# 🔐 用户工作空间信息（多用户隔离环境）
# ===========================================

# 当前用户会话ID
user_session_id = "{session_id}"

# 用户工作空间路径
user_workspace = Path(r"{session_manager.get_user_workspace(session_id)}")
user_uploads_dir = user_workspace / "uploads"
user_exports_dir = user_workspace / "exports"  # 导出文件请保存到这里
user_temp_dir = user_workspace / "temp"

print("🔐 用户工作空间信息:")
print(f"   会话ID: {{user_session_id}}")
print(f"   工作空间: {{user_workspace}}")
print(f"   上传目录: {{user_uploads_dir}}")
print(f"   导出目录: {{user_exports_dir}}")
print(f"   临时目录: {{user_temp_dir}}")
print()

# 用户工作空间操作函数
def save_to_exports(filename, data_or_path):
    '''将文件保存到用户导出目录'''
    from datetime import datetime
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_filename = f"{{timestamp}}_{{filename}}"
    export_path = user_exports_dir / safe_filename
    
    if isinstance(data_or_path, (pd.DataFrame,)):
        # 如果是DataFrame，保存为Excel
        if not filename.endswith(('.xlsx', '.xls')):
            filename += '.xlsx'
        data_or_path.to_excel(export_path, index=False)
    elif isinstance(data_or_path, str) and os.path.exists(data_or_path):
        # 如果是文件路径，复制文件
        import shutil
        shutil.copy2(data_or_path, export_path)
    else:
        # 其他情况，尝试写入文本
        with open(export_path, 'w', encoding='utf-8') as f:
            f.write(str(data_or_path))
    
    print(f"✅ 文件已保存到用户导出目录: {{export_path.name}}")
    return str(export_path)

def get_temp_path(filename):
    '''获取临时文件路径'''
    return str(user_temp_dir / filename)

# ===========================================
# 📊 原始Excel文件信息
# ==========================================="""
                
                # 添加当前文件信息
                if hasattr(st.session_state, 'current_file_path') and st.session_state.current_file_path:
                    default_code += f"""
# 当前Excel文件信息
excel_file_path = r"{st.session_state.current_file_path}"
excel_file_name = "{st.session_state.get('current_file_name', 'unknown.xlsx')}"

print("📊 当前Excel文件信息:")
print(f"   文件路径: {{excel_file_path}}")
print(f"   文件名: {{excel_file_name}}")
print()"""
                else:
                    default_code += f"""
# Excel文件信息（需要先选择或上传文件）
excel_file_path = None
excel_file_name = "请先选择Excel文件"

print("⚠️  请先在'📁 上传Excel文件'部分选择或上传Excel文件")
print()"""
                
                default_code += f"""
# ===========================================
# 📋 工作表数据概览
# ===========================================

# 所有工作表概览
print("📊 工作表概览:")
for i, sheet in enumerate(sheet_names, 1):
    safe_name = sheet.replace(' ', '_').replace('-', '_').replace('.', '_')
    df_shape = eval(f'df_{{safe_name}}').shape
    print(f"{{i}}. {{sheet}}: {{df_shape[0]}}行 × {{df_shape[1]}}列")
print()

# 当前工作表数据
current_df = df_{current_safe_name}
print(f"🎯 当前工作表: {st.session_state.current_sheet}")
print(f"数据形状: {{current_df.shape}}")
print("\\n数据类型:")
print(current_df.dtypes)
print("\\n前5行数据:")
print(current_df.head())

# ===========================================
# 💡 使用示例和最佳实践
# ===========================================

# 示例1: 跨工作表数据处理
print("\\n" + "="*50)
print("💡 示例1: 跨工作表分析")
print("="*50)
for sheet_name in sheet_names:
    safe_name = sheet_name.replace(' ', '_').replace('-', '_').replace('.', '_')
    df = eval(f'df_{{safe_name}}')
    print(f"{{sheet_name}} 工作表: {{len(df)}} 行数据, {{len(df.columns)}} 列")

# 示例2: 数据处理和导出（重要！）
print("\\n" + "="*50)
print("💡 示例2: 数据处理和导出到用户目录")
print("="*50)

# 处理数据示例
processed_df = current_df.copy()
# processed_df['新列'] = processed_df['某列'] * 2  # 根据实际列名修改

# 导出到用户导出目录（推荐方式）
# save_to_exports("处理后的数据.xlsx", processed_df)

# 示例3: 使用原始Excel文件进行高级操作
if excel_file_path:
    print("\\n" + "="*50)
    print("💡 示例3: 高级Excel操作")
    print("="*50)
    print("# 可以用于需要直接访问Excel文件的场景")
    print("# wb = pd.ExcelFile(excel_file_path)")
    print("# custom_df = pd.read_excel(excel_file_path, sheet_name='特定工作表', header=2)")

print("\\n" + "="*50)
print("🔐 数据安全提醒:")
print("- 所有文件自动保存到您的专属工作空间")
print("- 使用 save_to_exports() 函数将结果保存到导出目录")
print("- 导出的文件可在'数据工具'标签页下载")
print("="*50)

# ===========================================
# 🚀 开始您的数据分析
# ===========================================

# 在这里编写您的分析代码
# 记住：
# 1. 修改数据后，将结果保存回对应的df_变量
# 2. 导出文件使用 save_to_exports() 函数
# 3. 所有操作都在您的专属隔离环境中进行

# 保存修改到工作表变量（重要！）
# df_{current_safe_name} = processed_df  # 取消注释以保存修改
"""
                
                # 检查是否有保存的代码
                if 'excel_code' not in st.session_state:
                    st.session_state.excel_code = default_code
                
                # 代码输入框
                excel_code = st.text_area(
                    "在此编写处理Excel数据的Python代码",
                    value=st.session_state.excel_code,
                    height=400,
                    help="编写Python代码来处理Excel数据。可以访问原始Excel文件和所有工作表数据。"
                )
                
                # 保存代码
                st.session_state.excel_code = excel_code
                
                # 执行控制按钮
                col_exec, col_clear, col_reset, col_ai = st.columns([2, 1, 1, 1])
                
                with col_exec:
                    if st.button("▶️ 执行Excel代码", type="primary", use_container_width=True):
                        try:
                            # 准备执行环境 - 包含原始Excel文件访问
                            exec_globals = {
                                'pd': pd,
                                'np': np,
                                'px': px,
                                'go': go,
                                'st': st,
                                'os': os
                            }
                            
                            # 添加所有Excel工作表数据
                            for sheet_name, df in st.session_state.excel_data.items():
                                safe_name = sheet_name.replace(' ', '_').replace('-', '_').replace('.', '_')
                                exec_globals[f'df_{safe_name}'] = df.copy()  # 使用副本避免意外修改
                            
                            # 添加Excel文件信息
                            if hasattr(st.session_state, 'current_file_path') and st.session_state.current_file_path:
                                exec_globals['excel_file_path'] = st.session_state.current_file_path
                                exec_globals['excel_file_name'] = st.session_state.get('current_file_name', 'unknown.xlsx')
                            else:
                                exec_globals['excel_file_path'] = None
                                exec_globals['excel_file_name'] = "请先选择Excel文件"
                            
                            # 添加工作表关系信息
                            exec_globals['sheet_names'] = list(st.session_state.excel_data.keys())
                            exec_globals['sheet_info'] = {
                                name: {
                                    'shape': df.shape,
                                    'columns': list(df.columns),
                                    'dtypes': df.dtypes.to_dict()
                                }
                                for name, df in st.session_state.excel_data.items()
                            }
                            
                            # 添加用户工作空间相关变量和函数
                            from pathlib import Path
                            user_workspace = session_manager.get_user_workspace(session_id)
                            
                            exec_globals['user_session_id'] = session_id
                            exec_globals['user_workspace'] = user_workspace
                            exec_globals['user_uploads_dir'] = user_workspace / "uploads"
                            exec_globals['user_exports_dir'] = user_workspace / "exports"
                            exec_globals['user_temp_dir'] = user_workspace / "temp"
                            exec_globals['Path'] = Path
                            
                            # 定义用户工作空间操作函数
                            def save_to_exports(filename, data_or_path):
                                """将文件保存到用户导出目录"""
                                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                safe_filename = f"{timestamp}_{filename}"
                                export_path = user_workspace / "exports" / safe_filename
                                
                                # 确保导出目录存在
                                export_path.parent.mkdir(parents=True, exist_ok=True)
                                
                                if isinstance(data_or_path, pd.DataFrame):
                                    # 如果是DataFrame，保存为Excel
                                    if not filename.endswith(('.xlsx', '.xls')):
                                        export_path = export_path.with_suffix('.xlsx')
                                    data_or_path.to_excel(export_path, index=False)
                                elif isinstance(data_or_path, str) and os.path.exists(data_or_path):
                                    # 如果是文件路径，复制文件
                                    import shutil
                                    shutil.copy2(data_or_path, export_path)
                                else:
                                    # 其他情况，尝试写入文本
                                    with open(export_path, 'w', encoding='utf-8') as f:
                                        f.write(str(data_or_path))
                                
                                print(f"✅ 文件已保存到用户导出目录: {export_path.name}")
                                return str(export_path)
                            
                            def get_temp_path(filename):
                                """获取临时文件路径"""
                                temp_path = user_workspace / "temp" / filename
                                temp_path.parent.mkdir(parents=True, exist_ok=True)
                                return str(temp_path)
                            
                            def get_export_path(filename):
                                """获取导出文件路径"""
                                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                safe_filename = f"{timestamp}_{filename}"
                                export_path = user_workspace / "exports" / safe_filename
                                export_path.parent.mkdir(parents=True, exist_ok=True)
                                return str(export_path)
                            
                            # 文件保存拦截器 - 重定向常见的文件保存操作
                            original_open = open
                            created_files = []  # 记录创建的文件
                            
                            def intercepted_open(file, mode='r', **kwargs):
                                """拦截open函数，重定向文件保存到用户目录"""
                                if isinstance(file, str):
                                    # 检查是否是写入模式
                                    if 'w' in mode or 'a' in mode or 'x' in mode:
                                        # 获取文件名
                                        filename = os.path.basename(file)
                                        
                                        # 如果是相对路径或当前目录，重定向到用户导出目录
                                        if not os.path.isabs(file) or file.startswith('./') or file.startswith('.\\'):
                                            redirect_path = user_workspace / "exports" / filename
                                            redirect_path.parent.mkdir(parents=True, exist_ok=True)
                                            
                                            print(f"🔄 文件保存重定向: {file} -> {redirect_path}")
                                            created_files.append(str(redirect_path))
                                            return original_open(redirect_path, mode, **kwargs)
                                
                                return original_open(file, mode, **kwargs)
                            
                            # 拦截pandas to_excel等方法
                            original_to_excel = pd.DataFrame.to_excel
                            def intercepted_to_excel(self, excel_writer, **kwargs):
                                """拦截DataFrame.to_excel方法"""
                                if isinstance(excel_writer, str):
                                    filename = os.path.basename(excel_writer)
                                    if not os.path.isabs(excel_writer):
                                        redirect_path = user_workspace / "exports" / filename
                                        redirect_path.parent.mkdir(parents=True, exist_ok=True)
                                        print(f"🔄 Excel保存重定向: {excel_writer} -> {redirect_path}")
                                        created_files.append(str(redirect_path))
                                        return original_to_excel(self, redirect_path, **kwargs)
                                return original_to_excel(self, excel_writer, **kwargs)
                            
                            # 拦截json.dump方法
                            import json
                            original_json_dump = json.dump
                            def intercepted_json_dump(obj, fp, **kwargs):
                                """拦截json.dump方法"""
                                if hasattr(fp, 'name') and isinstance(fp.name, str):
                                    filename = os.path.basename(fp.name)
                                    if not os.path.isabs(fp.name):
                                        redirect_path = user_workspace / "exports" / filename
                                        redirect_path.parent.mkdir(parents=True, exist_ok=True)
                                        print(f"🔄 JSON保存重定向: {fp.name} -> {redirect_path}")
                                        created_files.append(str(redirect_path))
                                        with original_open(redirect_path, 'w', encoding='utf-8') as f:
                                            return original_json_dump(obj, f, **kwargs)
                                return original_json_dump(obj, fp, **kwargs)
                            
                            # 应用拦截器
                            exec_globals['open'] = intercepted_open
                            exec_globals['json'] = type('json_module', (), {
                                'dump': intercepted_json_dump,
                                'dumps': json.dumps,
                                'load': json.load,
                                'loads': json.loads
                            })()
                            
                            # 临时替换pandas方法
                            pd.DataFrame.to_excel = intercepted_to_excel
                            
                            # 添加函数到执行环境
                            exec_globals['save_to_exports'] = save_to_exports
                            exec_globals['get_temp_path'] = get_temp_path
                            exec_globals['get_export_path'] = get_export_path
                            exec_globals['created_files'] = created_files  # 让代码可以访问创建的文件列表
                            
                            # 重定向输出
                            import sys
                            from io import StringIO
                            old_stdout = sys.stdout
                            sys.stdout = mystdout = StringIO()
                            
                            # 执行代码
                            exec(excel_code, exec_globals)
                            
                            # 恢复原始函数
                            pd.DataFrame.to_excel = original_to_excel
                            
                            # 恢复输出
                            sys.stdout = old_stdout
                            output = mystdout.getvalue()
                            
                            # 检查并更新修改的数据
                            updated_sheets = []
                            for sheet_name in st.session_state.excel_data.keys():
                                safe_name = sheet_name.replace(' ', '_').replace('-', '_').replace('.', '_')
                                if f'df_{safe_name}' in exec_globals:
                                    old_shape = st.session_state.excel_data[sheet_name].shape
                                    new_df = exec_globals[f'df_{safe_name}']
                                    
                                    # 检查是否有修改
                                    if not new_df.equals(st.session_state.excel_data[sheet_name]):
                                        # 更新数据
                                        st.session_state.excel_processor.modified_data[sheet_name] = new_df
                                        st.session_state.excel_data[sheet_name] = new_df
                                        updated_sheets.append(f"{sheet_name} ({old_shape} → {new_df.shape})")
                            
                            # 检测生成的文件
                            generated_files = created_files.copy()
                            
                            # 额外检查导出目录中的新文件
                            exports_dir = user_workspace / "exports"
                            if exports_dir.exists():
                                # 获取5分钟内创建的文件
                                import time
                                current_time = time.time()
                                recent_files = []
                                
                                for file_path in exports_dir.iterdir():
                                    if file_path.is_file():
                                        file_mtime = file_path.stat().st_mtime
                                        if current_time - file_mtime < 300:  # 5分钟内
                                            file_path_str = str(file_path)
                                            if file_path_str not in generated_files:
                                                recent_files.append(file_path_str)
                                
                                generated_files.extend(recent_files)
                            
                            # 显示执行结果
                            st.success("✅ Excel代码执行成功")
                            
                            if output:
                                st.subheader("📄 执行输出:")
                                st.code(output, language="text")
                            
                            if updated_sheets:
                                st.info(f"📊 已更新的Excel工作表: {', '.join(updated_sheets)}")
                                st.info("💡 数据已保存到Excel处理器，可在'数据预览'中查看或直接导出")
                            else:
                                st.info("📋 代码执行完成，未检测到数据修改")
                            
                            # 处理生成的文件
                            if generated_files:
                                st.subheader("📁 生成的文件")
                                st.success(f"🎉 检测到 {len(generated_files)} 个生成的文件")
                                
                                # 分类显示文件
                                json_files = [f for f in generated_files if f.lower().endswith('.json')]
                                md_files = [f for f in generated_files if f.lower().endswith(('.md', '.markdown'))]
                                excel_files = [f for f in generated_files if f.lower().endswith(('.xlsx', '.xls'))]
                                other_files = [f for f in generated_files if f not in json_files + md_files + excel_files]
                                
                                # 显示JSON文件
                                if json_files:
                                    st.markdown("**📄 JSON数据文件:**")
                                    for json_file in json_files:
                                        file_path = Path(json_file)
                                        file_size = file_path.stat().st_size / 1024  # KB
                                        
                                        col1, col2 = st.columns([3, 1])
                                        with col1:
                                            st.write(f"📄 {file_path.name} ({file_size:.1f} KB)")
                                        with col2:
                                            try:
                                                with open(json_file, 'rb') as f:
                                                    file_data = f.read()
                                                st.download_button(
                                                    label="⬇️ 下载",
                                                    data=file_data,
                                                    file_name=file_path.name,
                                                    mime="application/json",
                                                    key=f"download_json_{file_path.stem}",
                                                    use_container_width=True
                                                )
                                            except Exception as e:
                                                st.error(f"下载失败: {e}")
                                
                                # 显示Markdown文件
                                if md_files:
                                    st.markdown("**📝 Markdown分析文件:**")
                                    for md_file in md_files:
                                        file_path = Path(md_file)
                                        file_size = file_path.stat().st_size / 1024  # KB
                                        
                                        col1, col2 = st.columns([3, 1])
                                        with col1:
                                            st.write(f"📝 {file_path.name} ({file_size:.1f} KB)")
                                        with col2:
                                            try:
                                                with open(md_file, 'rb') as f:
                                                    file_data = f.read()
                                                st.download_button(
                                                    label="⬇️ 下载",
                                                    data=file_data,
                                                    file_name=file_path.name,
                                                    mime="text/markdown",
                                                    key=f"download_md_{file_path.stem}",
                                                    use_container_width=True
                                                )
                                            except Exception as e:
                                                st.error(f"下载失败: {e}")
                                
                                # 显示Excel文件
                                if excel_files:
                                    st.markdown("**📊 Excel文件:**")
                                    for excel_file in excel_files:
                                        file_path = Path(excel_file)
                                        file_size = file_path.stat().st_size / 1024  # KB
                                        
                                        col1, col2 = st.columns([3, 1])
                                        with col1:
                                            st.write(f"📊 {file_path.name} ({file_size:.1f} KB)")
                                        with col2:
                                            try:
                                                with open(excel_file, 'rb') as f:
                                                    file_data = f.read()
                                                st.download_button(
                                                    label="⬇️ 下载",
                                                    data=file_data,
                                                    file_name=file_path.name,
                                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                                    key=f"download_excel_{file_path.stem}",
                                                    use_container_width=True
                                                )
                                            except Exception as e:
                                                st.error(f"下载失败: {e}")
                                
                                # 显示其他文件
                                if other_files:
                                    st.markdown("**📁 其他文件:**")
                                    for other_file in other_files:
                                        file_path = Path(other_file)
                                        file_size = file_path.stat().st_size / 1024  # KB
                                        
                                        col1, col2 = st.columns([3, 1])
                                        with col1:
                                            st.write(f"📁 {file_path.name} ({file_size:.1f} KB)")
                                        with col2:
                                            try:
                                                with open(other_file, 'rb') as f:
                                                    file_data = f.read()
                                                st.download_button(
                                                    label="⬇️ 下载",
                                                    data=file_data,
                                                    file_name=file_path.name,
                                                    mime="application/octet-stream",
                                                    key=f"download_other_{file_path.stem}",
                                                    use_container_width=True
                                                )
                                            except Exception as e:
                                                st.error(f"下载失败: {e}")
                                
                                # 提示信息
                                st.info("💡 所有生成的文件已保存到您的专属导出目录，您也可以在'🛠️ 数据工具'标签页中管理这些文件")
                            
                            else:
                                # 检查是否有新文件保存到导出目录（兜底检查）
                                exports_dir = user_workspace / "exports"
                                if exports_dir.exists():
                                    export_files = list(exports_dir.glob("*"))
                                    if export_files:
                                        # 找到最新的文件
                                        latest_files = sorted(export_files, key=lambda x: x.stat().st_mtime, reverse=True)[:3]
                                        if latest_files:
                                            st.info(f"📁 检测到导出文件: {', '.join([f.name for f in latest_files])}")
                                            st.info("💡 您可以在'🛠️ 数据工具'标签页中下载导出的文件")
                            
                        except Exception as e:
                            # 确保恢复原始函数
                            try:
                                pd.DataFrame.to_excel = original_to_excel
                            except:
                                pass
                            st.error(f"❌ 代码执行错误: {str(e)}")
                            st.code(f"错误详情:\n{traceback.format_exc()}", language="text")
                
                with col_clear:
                    if st.button("🗑️ 清空", use_container_width=True):
                        st.session_state.excel_code = ""
                        st.rerun()
                
                with col_reset:
                    if st.button("🔄 重置", use_container_width=True):
                        st.session_state.excel_code = default_code
                        st.rerun()
                
                with col_ai:
                    if st.button("🤖 AI助手", use_container_width=True, help="使用AI生成代码"):
                        st.session_state.show_ai_helper = not st.session_state.get('show_ai_helper', False)
                        st.rerun()
                
                # AI代码生成助手 - 增强版，包含完整Excel信息
                if st.session_state.get('show_ai_helper', False):
                    with st.expander("🤖 AI代码生成助手", expanded=True):
                        if not api_key:
                            st.warning("⚠️ 请先配置OpenAI API Key")
                        else:
                            ai_analyzer = EnhancedAIAnalyzer(api_key, base_url, selected_model)
                            
                            # 提供更详细的任务描述输入
                            col_task, col_context = st.columns([2, 1])
                            
                            with col_task:
                                task_description = st.text_area(
                                    "描述您需要完成的Excel数据处理任务",
                                    placeholder="例如：分析所有工作表的数据关系、生成跨表汇总报告、执行复杂的业务计算等...",
                                    height=100,
                                    key="excel_ai_task"
                                )
                            
                            with col_context:
                                st.markdown("**📊 当前Excel信息:**")
                                current_file_name = st.session_state.get('current_file_name', '未选择')
                                st.info(f"文件: {current_file_name}")
                                st.info(f"工作表数: {len(st.session_state.excel_data)}")
                                for name in list(st.session_state.excel_data.keys())[:3]:
                                    st.info(f"• {name}")
                                if len(st.session_state.excel_data) > 3:
                                    st.info(f"... 还有{len(st.session_state.excel_data)-3}个工作表")
                            
                            # 增强的代码生成，包含工作表关系信息
                            if st.button("🚀 生成Excel处理代码", type="secondary", use_container_width=True):
                                if task_description.strip():
                                    with st.spinner("正在生成Excel处理代码..."):
                                        # 传递更完整的Excel结构信息给AI
                                        enhanced_excel_data = {}
                                        for sheet_name, df in st.session_state.excel_data.items():
                                            enhanced_excel_data[sheet_name] = {
                                                'dataframe': df,
                                                'shape': df.shape,
                                                'columns': list(df.columns),
                                                'sample_data': df.head(3).to_dict() if not df.empty else {},
                                                'dtypes': df.dtypes.to_dict()
                                            }
                                        
                                        code = ai_analyzer.generate_enhanced_code_solution(
                                            task_description, 
                                            enhanced_excel_data,
                                            st.session_state.get('current_file_name', 'Excel文件')
                                        )
                                        st.session_state.excel_code = code
                                        st.success("✅ 代码已生成并插入到上方编辑器")
                                        st.rerun()
                
                # 快速导出修改后的Excel
                st.subheader("📤 快速导出")
                if st.button("💾 导出修改后的Excel文件", type="secondary", use_container_width=True):
                    try:
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        output_filename = f"processed_excel_{timestamp}.xlsx"
                        output_path = st.session_state.excel_processor.export_to_excel(output_filename)
                        
                        with open(output_path, 'rb') as f:
                            file_data = f.read()
                        
                        st.download_button(
                            label="⬇️ 下载Excel文件",
                            data=file_data,
                            file_name=output_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                        
                        st.success("✅ Excel文件已准备就绪")
                        os.unlink(output_path)
                        
                    except Exception as e:
                        st.error(f"❌ 导出失败: {str(e)}")
            
            else:
                st.warning("⚠️ 请先在'数据预览'标签中选择工作表")
        
        # Tab 4: 数据工具
        with tab4:
            st.header("🛠️ 数据工具")
            
            if st.session_state.current_sheet:
                current_df = st.session_state.excel_data[st.session_state.current_sheet]
                
                # 数据清洗工具
                with st.expander("🧹 数据清洗工具", expanded=False):
                    st.subheader("填充缺失值")
                    
                    # 修复Series比较错误 - 完全修复版本
                    columns_with_missing = []
                    for col in current_df.columns:
                        try:
                            missing_count = current_df[col].isnull().sum()
                            # 确保missing_count是标量值进行比较
                            if isinstance(missing_count, (int, float)) and missing_count > 0:
                                columns_with_missing.append(col)
                        except Exception as e:
                            # 如果检查失败，跳过该列
                            continue
                    
                    if columns_with_missing:
                        selected_col = st.selectbox("选择列", columns_with_missing, key="missing_col_selector")
                        
                        fill_methods = {
                            "均值": "mean",
                            "中位数": "median", 
                            "众数": "mode",
                            "前向填充": "forward",
                            "后向填充": "backward",
                            "自定义值": "custom"
                        }
                        
                        fill_method = st.selectbox("填充方法", list(fill_methods.keys()))
                        
                        custom_value = None
                        if fill_method == "自定义值":
                            custom_value = st.text_input("自定义填充值")
                        
                        if st.button("执行填充", type="primary", key="fill_missing_btn"):
                            success, message = st.session_state.excel_processor.fill_missing_values(
                                st.session_state.current_sheet,
                                selected_col,
                                fill_methods[fill_method],
                                custom_value
                            )
                            
                            if success:
                                st.markdown(f'<div class="success-message">{message}</div>', unsafe_allow_html=True)
                                st.session_state.excel_data[st.session_state.current_sheet] = st.session_state.excel_processor.modified_data[st.session_state.current_sheet]
                                st.rerun()
                            else:
                                st.markdown(f'<div class="error-message">{message}</div>', unsafe_allow_html=True)
                    else:
                        st.info("当前工作表没有缺失值")
                
                # 统计分析工具
                with st.expander("📊 统计分析"):
                    if st.button("生成统计汇总", type="primary", key="stats_btn"):
                        success, message = st.session_state.excel_processor.add_summary_statistics(st.session_state.current_sheet)
                        if success:
                            st.markdown(f'<div class="success-message">{message}</div>', unsafe_allow_html=True)
                            new_sheet_name = f"{st.session_state.current_sheet}_统计汇总"
                            st.session_state.excel_data[new_sheet_name] = st.session_state.excel_processor.modified_data[new_sheet_name]
                            st.rerun()
                        else:
                            st.markdown(f'<div class="error-message">{message}</div>', unsafe_allow_html=True)
                
                # 导出工具
                with st.expander("📤 导出数据"):
                    st.subheader("导出处理后的Excel文件")
                    
                    export_filename = st.text_input(
                        "文件名",
                        value=f"processed_excel_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        key="export_filename"
                    )
                    
                    if st.button("🔄 生成并导出Excel文件", type="primary", use_container_width=True):
                        try:
                            # 使用用户导出路径
                            export_path = session_manager.get_export_path(session_id, export_filename)
                            
                            # 导出到用户工作空间
                            st.session_state.excel_processor.export_to_excel(str(export_path))
                            
                            with open(export_path, 'rb') as f:
                                file_data = f.read()
                            
                            st.download_button(
                                label="⬇️ 下载处理后的文件",
                                data=file_data,
                                file_name=export_filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                            
                            st.markdown(f'<div class="success-message">✅ 文件已准备就绪，点击下载按钮获取</div>', unsafe_allow_html=True)
                            
                        except Exception as e:
                            st.markdown(f'<div class="error-message">❌ 导出失败: {str(e)}</div>', unsafe_allow_html=True)
                
                # 用户导出文件管理
                with st.expander("📁 我的导出文件", expanded=False):
                    st.subheader("管理您通过代码生成的文件")
                    
                    # 获取用户导出目录的文件
                    user_workspace = session_manager.get_user_workspace(session_id)
                    if user_workspace:
                        exports_dir = user_workspace / "exports"
                        
                        if exports_dir.exists():
                            export_files = []
                            for file_path in exports_dir.iterdir():
                                if file_path.is_file():
                                    stat_info = file_path.stat()
                                    export_files.append({
                                        'name': file_path.name,
                                        'path': file_path,
                                        'size_mb': round(stat_info.st_size / (1024 * 1024), 2),
                                        'modified': datetime.fromtimestamp(stat_info.st_mtime)
                                    })
                            
                            # 按修改时间排序
                            export_files.sort(key=lambda x: x['modified'], reverse=True)
                            
                            if export_files:
                                st.info(f"📊 找到 {len(export_files)} 个导出文件")
                                
                                # 显示文件列表
                                for i, file_info in enumerate(export_files):
                                    col1, col2, col3 = st.columns([3, 1, 1])
                                    
                                    with col1:
                                        st.write(f"**{file_info['name']}**")
                                        st.caption(f"大小: {file_info['size_mb']} MB | 修改时间: {file_info['modified'].strftime('%Y-%m-%d %H:%M')}")
                                    
                                    with col2:
                                        # 下载按钮
                                        try:
                                            with open(file_info['path'], 'rb') as f:
                                                file_data = f.read()
                                            
                                            st.download_button(
                                                label="⬇️ 下载",
                                                data=file_data,
                                                file_name=file_info['name'],
                                                mime="application/octet-stream",
                                                key=f"download_export_{i}",
                                                use_container_width=True
                                            )
                                        except Exception as e:
                                            st.error(f"下载失败: {e}")
                                    
                                    with col3:
                                        # 删除按钮
                                        if st.button("🗑️ 删除", key=f"delete_export_{i}", use_container_width=True):
                                            try:
                                                file_info['path'].unlink()
                                                st.success(f"✅ 已删除 {file_info['name']}")
                                                st.rerun()
                                            except Exception as e:
                                                st.error(f"❌ 删除失败: {e}")
                                    
                                    st.markdown("---")
                                
                                # 批量操作
                                st.subheader("批量操作")
                                col_batch1, col_batch2 = st.columns(2)
                                
                                with col_batch1:
                                    if st.button("🗑️ 清空所有导出文件", use_container_width=True):
                                        try:
                                            for file_info in export_files:
                                                file_info['path'].unlink()
                                            st.success(f"✅ 已清空 {len(export_files)} 个导出文件")
                                            st.rerun()
                                        except Exception as e:
                                            st.error(f"❌ 清空失败: {e}")
                                
                                with col_batch2:
                                    # 计算总大小
                                    total_size = sum(f['size_mb'] for f in export_files)
                                    st.metric("导出文件总大小", f"{total_size:.2f} MB")
                            
                            else:
                                st.info("📂 您还没有生成任何导出文件")
                                st.markdown("""
                                **💡 如何生成导出文件：**
                                1. 在"💻 代码执行"标签页中编写代码
                                2. 使用 `save_to_exports("文件名.xlsx", dataframe)` 函数保存文件
                                3. 导出的文件会自动出现在这里供下载
                                """)
                        
                        else:
                            st.info("📂 导出目录不存在，将在首次导出时创建")
                    
                    else:
                        st.error("❌ 无法访问用户工作空间")
            
            else:
                st.info("📋 请先在'数据预览'标签中选择工作表")
    
    # 文档分析界面
    elif analysis_mode == "📄 文档分析" and st.session_state.document_data:
        # Tab 1: 文档预览
        with doc_tab1:
            st.header("📄 文档预览与管理")
            
            file_info = st.session_state.document_data.get('file_info', {})
            st.success(f"✅ 成功载入文档: {file_info.get('name', 'Unknown')}")
            
            # 文档基本信息卡片
            col_a, col_b, col_c, col_d = st.columns(4)
            with col_a:
                st.markdown(f'<div class="metric-card"><h3>{file_info.get("type", "Unknown").upper()}</h3><p>文档类型</p></div>', unsafe_allow_html=True)
            with col_b:
                st.markdown(f'<div class="metric-card"><h3>{file_info.get("size_mb", 0)}</h3><p>文件大小(MB)</p></div>', unsafe_allow_html=True)
            with col_c:
                preview_data = st.session_state.document_data.get('preview_data', {})
                page_count = preview_data.get('page_count', 0)
                st.markdown(f'<div class="metric-card"><h3>{page_count}</h3><p>估算页数</p></div>', unsafe_allow_html=True)
            with col_d:
                word_count = preview_data.get('word_count', 0)
                st.markdown(f'<div class="metric-card"><h3>{word_count}</h3><p>字数统计</p></div>', unsafe_allow_html=True)
            
            # 文档预览
            st.subheader("📝 文档内容预览")
            st.info("💡 此预览已清除格式，保留结构，限制前10页内容，用于AI理解")
            
            try:
                if st.session_state.document_processor:
                    preview_content = st.session_state.document_processor.get_document_preview(max_chars=20000)
                    
                    if preview_content and preview_content != "请先加载文档":
                        # 计算内容统计信息
                        char_count = len(preview_content)
                        word_count = len(preview_content.split())
                        line_count = len(preview_content.split('\n'))
                        
                        # 使用可折叠的下拉框显示预览内容
                        with st.expander(f"📄 查看文档内容 ({word_count:,} 词, {char_count:,} 字符, {line_count} 行)", expanded=False):
                            st.markdown("### MarkItDown 清洗结果")
                            st.markdown("此内容已格式化为Markdown，便于AI理解和分析：")
                            
                            # 使用统一的容器渲染函数
                            render_content_container(preview_content, 'document-preview')
                            
                            # 添加一些操作按钮
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                if st.button("📋 复制内容"):
                                    st.info("💡 请使用浏览器的选择和复制功能")
                            with col2:
                                st.download_button(
                                    label="💾 下载预览",
                                    data=preview_content,
                                    file_name=f"{file_info.get('name', 'document')}_preview.md",
                                    mime="text/markdown"
                                )
                            with col3:
                                if st.button("🔄 刷新预览"):
                                    st.rerun()
                    else:
                        st.warning("⚠️ 无法生成文档预览")
                else:
                    st.error("❌ 文档处理器不可用")
                    
            except Exception as e:
                st.error(f"❌ 预览生成失败: {str(e)}")
            
            # 结构化分析摘要
            st.subheader("🏗️ 文档结构摘要")
            try:
                if st.session_state.document_processor:
                    structure_summary = st.session_state.document_processor.get_structure_summary()
                    if structure_summary and structure_summary != "请先加载文档":
                        # 计算结构摘要的统计信息
                        summary_lines = len(structure_summary.split('\n'))
                        summary_chars = len(structure_summary)
                        
                        with st.expander(f"📋 查看详细结构分析 ({summary_lines} 行)", expanded=True):
                            st.markdown("### 文档结构化分析结果")
                            st.markdown("基于原始文档格式提取的结构信息：")
                            
                            # 使用统一的容器渲染函数
                            render_content_container(structure_summary, 'document-structure')
                            
                            # 添加操作按钮
                            col1, col2 = st.columns(2)
                            with col1:
                                st.download_button(
                                    label="💾 下载结构分析",
                                    data=structure_summary,
                                    file_name=f"{file_info.get('name', 'document')}_structure.md",
                                    mime="text/markdown"
                                )
                            with col2:
                                if st.button("🔄 重新分析结构"):
                                    st.rerun()
                    else:
                        st.warning("⚠️ 无法生成结构摘要")
                else:
                    st.error("❌ 文档处理器不可用")
            except Exception as e:
                st.error(f"❌ 结构分析失败: {str(e)}")
        
        # Tab 2: AI文档分析
        with doc_tab2:
            st.header("🤖 AI 文档智能分析")
            
            # 轻量级文档结构分析（无需API）
            st.subheader("📋 文档结构化分析")
            st.info("💡 即使没有配置AI API，您也可以获得文档的结构化分析")
            
            # 添加分析按钮和结果显示
            col_doc_analyze, col_doc_clear = st.columns([3, 1])
            
            with col_doc_analyze:
                if st.button("🔍 结构化分析文档", type="secondary", use_container_width=True):
                    if hasattr(st.session_state, 'current_doc_path') and st.session_state.current_doc_path:
                        try:
                            with st.spinner("📊 正在进行文档结构化分析..."):
                                # 获取已有的分析结果
                                structure_analysis = st.session_state.document_data.get('structure_analysis', {})
                                if structure_analysis:
                                    # 格式化结构分析结果
                                    analysis_text = "# 📊 文档结构化分析结果\n\n"
                                    
                                    # 基本信息
                                    file_info = st.session_state.document_data.get('file_info', {})
                                    analysis_text += f"**文件名**: {file_info.get('name', 'Unknown')}\n"
                                    analysis_text += f"**文档类型**: {file_info.get('type', 'Unknown').upper()}\n"
                                    analysis_text += f"**文件大小**: {file_info.get('size_mb', 0)} MB\n\n"
                                    
                                    # 结构特征
                                    if file_info.get('type') == 'docx':
                                        analysis_text += "## 📋 DOCX文档结构\n"
                                        analysis_text += f"- **段落数**: {structure_analysis.get('total_paragraphs', 0)}\n"
                                        analysis_text += f"- **表格数**: {structure_analysis.get('tables_count', 0)}\n"
                                        analysis_text += f"- **图片数**: {structure_analysis.get('images_count', 0)}\n"
                                    elif file_info.get('type') == 'pdf':
                                        analysis_text += "## 📋 PDF文档结构\n"
                                        analysis_text += f"- **页数**: {structure_analysis.get('total_pages', 0)}\n"
                                        analysis_text += f"- **图片数**: {structure_analysis.get('images_count', 0)}\n"
                                    
                                    # 标题层级
                                    headings = structure_analysis.get('headings', {})
                                    if headings:
                                        analysis_text += "\n## 🏷️ 标题层级结构\n"
                                        for level in sorted(headings.keys()):
                                            heading_list = headings[level]
                                            analysis_text += f"### {level}级标题 (共{len(heading_list)}个)\n"
                                            for heading in heading_list[:3]:
                                                text = heading.get('text', str(heading))[:100]
                                                analysis_text += f"- {text}\n"
                                            if len(heading_list) > 3:
                                                analysis_text += f"- ... 还有{len(heading_list) - 3}个\n"
                                    
                                    # 字体使用
                                    fonts = structure_analysis.get('fonts_used', [])
                                    if fonts:
                                        analysis_text += f"\n## 🔤 字体使用情况\n"
                                        analysis_text += f"- **字体种类数**: {len(fonts)}\n"
                                        analysis_text += f"- **主要字体**: {', '.join(fonts[:5])}\n"
                                    
                                    st.session_state.quick_doc_analysis = analysis_text
                                    st.success("✅ 文档结构化分析完成！")
                                    st.rerun()
                                else:
                                    st.error("❌ 无法获取结构分析数据")
                        except Exception as e:
                            st.error(f"❌ 结构分析失败: {str(e)}")
                    else:
                        st.warning("⚠️ 请先上传文档文件")
            
            with col_doc_clear:
                if st.button("🗑️ 清除分析", use_container_width=True):
                    if 'quick_doc_analysis' in st.session_state:
                        del st.session_state.quick_doc_analysis
                        st.rerun()
            
            # 显示结构分析结果
            if 'quick_doc_analysis' in st.session_state and st.session_state.quick_doc_analysis:
                analysis_content = st.session_state.quick_doc_analysis
                analysis_lines = len(analysis_content.split('\n'))
                analysis_words = len(analysis_content.split())
                
                st.subheader("📊 文档结构分析结果")
                with st.expander(f"📋 查看详细分析 ({analysis_words} 词, {analysis_lines} 行)", expanded=True):
                    st.markdown("### 结构化分析报告")
                    st.markdown("基于文档原始格式提取的完整结构信息：")
                    
                    # 以markdown格式渲染分析内容
                    st.markdown(analysis_content)
                    
                    # 添加操作按钮
                    col1, col2 = st.columns(2)
                    with col1:
                        st.download_button(
                            label="💾 下载分析报告",
                            data=analysis_content,
                            file_name=f"{file_info.get('name', 'document')}_analysis.md",
                            mime="text/markdown"
                        )
                    with col2:
                        if st.button("📋 复制分析结果"):
                            st.info("💡 请使用浏览器的选择和复制功能")
                
                # 功能说明和提示
                st.info("📝 **结构化分析说明**：\n"
                       "- 📄 **原始格式解析**：直接从文档原始格式提取结构信息\n"
                       "- 🏷️ **标题层级识别**：自动识别不同级别的标题和章节\n"
                       "- 🔤 **字体样式分析**：统计文档中使用的字体类型\n"
                       "- 📊 **内容组织结构**：分析段落、表格、图片的分布")
                
                # 如果有API配置，提供将分析结果作为AI分析基础的选项
                if api_key:
                    st.success("💡 **AI分析提示**：上述结构分析将自动作为深度AI分析的基础信息，提高AI理解准确性！")
            
            # 分隔线
            st.markdown("---")
            
            # 深度AI分析功能
            st.subheader("🧠 深度AI文档分析")
            
            if not api_key:
                st.warning("⚠️ 请在侧边栏配置OpenAI API Key以使用深度AI分析功能")
            else:
                # 初始化文档AI分析器
                try:
                    from document_ai_analyzer import EnhancedDocumentAIAnalyzer
                    doc_ai_analyzer = EnhancedDocumentAIAnalyzer(api_key, base_url, selected_model)
                    
                    # AI分析控制
                    col_doc_ai_analyze, col_doc_ai_refresh = st.columns([3, 1])
                    
                    with col_doc_ai_analyze:
                        if st.button("🔍 开始AI深度文档分析", type="primary", use_container_width=True):
                            with st.spinner("🧠 AI正在深度分析您的文档..."):
                                # 获取文档结构分析结果
                                structure_info = ""
                                if 'quick_doc_analysis' in st.session_state and st.session_state.quick_doc_analysis:
                                    structure_info = st.session_state.quick_doc_analysis
                                
                                # 进行AI深度分析
                                analysis = doc_ai_analyzer.analyze_document_structure(st.session_state.document_data)
                                
                                # 构建完整的分析报告
                                if structure_info:
                                    combined_analysis = f"""## 📋 文档结构解析

{structure_info}

---

## 🎯 AI深度文档分析

{analysis}"""
                                else:
                                    combined_analysis = analysis
                                
                                st.session_state.document_analysis = combined_analysis
                                st.session_state.doc_chat_history.append({
                                    "role": "assistant",
                                    "content": f"**📋 文档深度分析报告**\n\n{combined_analysis}"
                                })
                    
                    with col_doc_ai_refresh:
                        if st.button("🔄 重新分析", use_container_width=True):
                            st.session_state.document_analysis = ""
                            st.session_state.doc_chat_history = []
                            st.rerun()
                    
                    # 快速操作按钮
                    st.subheader("⚡ 智能文档分析")
                    col_doc_quick1, col_doc_quick2 = st.columns(2)
                    
                    doc_quick_actions = [
                        ("🎯 文档用途识别", "请分析这个文档的用途和类型，识别其主要功能和应用场景"),
                        ("📋 内容主题分析", "请分析文档的主要内容主题，识别核心议题和关键信息"),
                        ("🏗️ 结构特点分析", "请分析文档的组织结构特点，评估其逻辑性和可读性"),
                        ("🔍 关键信息提取", "请识别文档中的关键信息，如重要日期、金额、人名、条款等")
                    ]
                    
                    for i, (title, prompt) in enumerate(doc_quick_actions):
                        col = col_doc_quick1 if i % 2 == 0 else col_doc_quick2
                        with col:
                            if st.button(title, use_container_width=True, key=f"doc_quick_{i}"):
                                st.session_state.doc_chat_history.append({
                                    "role": "user",
                                    "content": prompt
                                })
                                
                                with st.spinner("🤔 AI正在分析..."):
                                    response = doc_ai_analyzer.chat_with_document(
                                        prompt,
                                        st.session_state.document_data,
                                        st.session_state.document_analysis
                                    )
                                    st.session_state.doc_chat_history.append({
                                        "role": "assistant",
                                        "content": response
                                    })
                                st.rerun()
                    
                    # 显示分析结果
                    if st.session_state.document_analysis:
                        ai_analysis = st.session_state.document_analysis
                        ai_lines = len(ai_analysis.split('\n'))
                        ai_words = len(ai_analysis.split())
                        ai_chars = len(ai_analysis)
                        
                        st.subheader("📊 AI文档分析结果")
                        with st.expander(f"📋 查看完整分析报告 ({ai_words:,} 词, {ai_chars:,} 字符, {ai_lines} 行)", expanded=True):
                            st.markdown("### AI深度文档分析报告")
                            st.markdown("由AI结合结构分析和内容理解生成的完整分析：")
                            
                            # 使用统一的容器渲染函数
                            render_content_container(ai_analysis, 'ai-analysis')
                            
                            # 添加操作按钮
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                st.download_button(
                                    label="💾 下载AI报告",
                                    data=ai_analysis,
                                    file_name=f"{file_info.get('name', 'document')}_ai_analysis.md",
                                    mime="text/markdown"
                                )
                            with col2:
                                if st.button("📋 复制AI分析"):
                                    st.info("💡 请使用浏览器的选择和复制功能")
                            with col3:
                                if st.button("🔄 刷新AI分析"):
                                    st.rerun()
                    
                    # 对话历史
                    if st.session_state.doc_chat_history:
                        chat_count = len(st.session_state.doc_chat_history)
                        st.subheader("💬 AI对话历史")
                        
                        with st.expander(f"💬 查看对话记录 (共 {chat_count} 条)", expanded=False):
                            st.markdown("### 文档分析对话记录")
                            st.markdown("您与AI关于文档分析的完整对话：")
                            
                            # 使用统一的对话容器渲染函数
                            render_chat_container(st.session_state.doc_chat_history, 'doc-chat')
                            
                            # 导出对话历史
                            chat_export = "\n\n".join([
                                f"{'👤 用户: ' if chat['role'] == 'user' else '🤖 AI: '}{chat['content']}"
                                for chat in st.session_state.doc_chat_history
                            ])
                            
                            st.download_button(
                                label="💾 下载对话记录",
                                data=chat_export,
                                file_name=f"{file_info.get('name', 'document')}_chat_history.md",
                                mime="text/markdown"
                            )
                    
                    # 用户输入
                    user_input = st.text_area(
                        "💭 向AI提问",
                        placeholder="例如：分析文档重点、查找关键信息、提供改进建议等...",
                        height=80,
                        key="doc_ai_chat_input"
                    )
                    
                    col_doc_send, col_doc_clear_chat = st.columns([1, 1])
                    
                    with col_doc_send:
                        if st.button("📤 发送", type="primary", use_container_width=True):
                            if user_input.strip():
                                st.session_state.doc_chat_history.append({
                                    "role": "user",
                                    "content": user_input
                                })
                                
                                with st.spinner("🤔 AI正在思考..."):
                                    response = doc_ai_analyzer.chat_with_document(
                                        user_input,
                                        st.session_state.document_data,
                                        st.session_state.document_analysis
                                    )
                                    st.session_state.doc_chat_history.append({
                                        "role": "assistant",
                                        "content": response
                                    })
                                st.rerun()
                    
                    with col_doc_clear_chat:
                        if st.button("🗑️ 清空对话", use_container_width=True):
                            st.session_state.doc_chat_history = []
                            st.rerun()
                            
                except ImportError:
                    st.error("❌ 无法导入文档AI分析器，请确保document_ai_analyzer.py文件存在")
                except Exception as e:
                    st.error(f"❌ AI分析器初始化失败: {str(e)}")
        
        # Tab 3: 代码执行
        with doc_tab3:
            st.header("💻 文档代码执行")
            
            if not api_key:
                st.warning("⚠️ AI代码生成需要配置API Key")
            else:
                # AI代码生成助手
                col_doc_ai, col_doc_manual = st.columns([1, 1])
                
                with col_doc_ai:
                    if st.button("🤖 AI助手", use_container_width=True, help="使用AI生成文档处理代码"):
                        st.session_state.show_doc_ai_helper = not st.session_state.get('show_doc_ai_helper', False)
                        st.rerun()
                
                # AI代码生成助手
                if st.session_state.get('show_doc_ai_helper', False):
                    with st.expander("🤖 AI文档代码生成助手", expanded=True):
                        try:
                            from document_ai_analyzer import EnhancedDocumentAIAnalyzer
                            doc_ai_analyzer = EnhancedDocumentAIAnalyzer(api_key, base_url, selected_model)
                            
                            task_description = st.text_area(
                                "描述您需要完成的文档处理任务",
                                placeholder="例如：搜索所有包含'合同编号'的段落并提取上下文、分析文档中的关键信息、生成文档摘要等...",
                                height=100,
                                key="doc_ai_task"
                            )
                            
                            if st.button("🔮 AI生成代码", type="primary", use_container_width=True):
                                if task_description.strip():
                                    with st.spinner("🤖 AI正在生成代码..."):
                                        generated_code = doc_ai_analyzer.generate_document_code_solution(
                                            task_description,
                                            st.session_state.document_data,
                                            st.session_state.current_doc_name
                                        )
                                        st.session_state.doc_generated_code = generated_code
                                        st.success("✅ 代码生成完成！")
                                        st.rerun()
                                else:
                                    st.warning("⚠️ 请描述您的任务需求")
                        except ImportError:
                            st.error("❌ 无法导入文档AI分析器")
                        except Exception as e:
                            st.error(f"❌ AI代码生成失败: {str(e)}")
            
            # 显示生成的代码
            if 'doc_generated_code' in st.session_state:
                st.subheader("🔮 AI生成的代码")
                st.code(st.session_state.doc_generated_code, language='python')
                
                if st.button("📋 复制到编辑器", use_container_width=True):
                    st.session_state.doc_code_input = st.session_state.doc_generated_code
                    st.success("✅ 代码已复制到编辑器")
                    st.rerun()
            
            # 代码编辑器
            st.subheader("📝 Python代码编辑器")
            
            # 提供示例代码
            default_doc_code = '''# 文档分析示例代码
from document_analyzer import DocumentAnalyzer
from document_utils import AdvancedDocumentProcessor

# 初始化处理器
processor = AdvancedDocumentProcessor()

# 分析文档（文件路径会自动替换）
doc_path = "current_document_path"
analysis_result = processor.load_document(doc_path)

# 获取文档基本信息
file_info = analysis_result["file_info"]
print(f"文档名: {file_info['name']}")
print(f"类型: {file_info['type']}")
print(f"大小: {file_info['size_mb']} MB")

# 搜索关键词示例
keyword = "重要信息"  # 修改为您要搜索的关键词
search_results = processor.search_content(keyword, context_lines=2)

print(f"\\n=== 搜索关键词: {keyword} ===")
for i, result in enumerate(search_results, 1):
    print(f"结果 {i}:")
    print(f"  位置: 第{result['line_number']}行")
    print(f"  内容: {result['matched_line']}")
    print(f"  上下文:\\n{result['context']}")
    print("-" * 50)

# 导出分析结果
json_file, md_file = processor.export_analysis_result()
print(f"\\n分析结果已导出:")
print(f"JSON文件: {json_file}")
print(f"报告文件: {md_file}")
'''
            
            doc_code_input = st.text_area(
                "输入Python代码",
                value=st.session_state.get('doc_code_input', default_doc_code),
                height=300,
                key="doc_code_editor"
            )
            
            if st.button("🚀 执行文档分析代码", type="primary", use_container_width=True):
                if doc_code_input.strip():
                    with st.spinner("🔄 正在执行代码..."):
                        try:
                            # 创建安全的执行环境
                            exec_globals = {
                                '__builtins__': __builtins__,
                                'print': print,
                                'len': len,
                                'str': str,
                                'int': int,
                                'float': float,
                                'list': list,
                                'dict': dict,
                                'enumerate': enumerate,
                                'range': range,
                            }
                            
                            # 导入文档处理模块
                            try:
                                from document_analyzer import DocumentAnalyzer
                                from document_utils import AdvancedDocumentProcessor, DocumentSearchEngine
                                exec_globals['DocumentAnalyzer'] = DocumentAnalyzer
                                exec_globals['AdvancedDocumentProcessor'] = AdvancedDocumentProcessor
                                exec_globals['DocumentSearchEngine'] = DocumentSearchEngine
                            except ImportError as e:
                                st.error(f"❌ 导入文档处理模块失败: {str(e)}")
                                return
                            
                            # 替换当前文档路径
                            if hasattr(st.session_state, 'current_doc_path'):
                                doc_code_input = doc_code_input.replace('current_document_path', st.session_state.current_doc_path)
                                doc_code_input = doc_code_input.replace('"current_document_path"', f'"{st.session_state.current_doc_path}"')
                            
                            # 执行代码
                            exec_output = io.StringIO()
                            
                            # 重定向print输出
                            import sys
                            old_stdout = sys.stdout
                            sys.stdout = exec_output
                            
                            try:
                                exec(doc_code_input, exec_globals)
                                result = exec_output.getvalue()
                                
                                if result:
                                    st.subheader("📋 执行结果")
                                    st.text(result)
                                else:
                                    st.success("✅ 代码执行完成（无输出）")
                                    
                            finally:
                                sys.stdout = old_stdout
                                
                        except Exception as e:
                            st.error(f"❌ 代码执行错误: {str(e)}")
                            st.error("请检查代码语法和逻辑")
                else:
                    st.warning("⚠️ 请输入要执行的代码")
        
        # Tab 4: 搜索工具
        with doc_tab4:
            st.header("🔍 文档搜索工具")
            
            if st.session_state.document_processor:
                # 关键词搜索
                st.subheader("🎯 关键词搜索")
                
                col_search1, col_search2 = st.columns([3, 1])
                
                with col_search1:
                    search_keyword = st.text_input(
                        "输入搜索关键词",
                        placeholder="例如: 合同编号、重要条款、日期等...",
                        key="doc_search_keyword"
                    )
                
                with col_search2:
                    context_lines = st.number_input(
                        "上下文行数",
                        min_value=1,
                        max_value=10,
                        value=3,
                        key="doc_context_lines"
                    )
                
                if st.button("🔍 搜索", type="primary", use_container_width=True):
                    if search_keyword.strip():
                        with st.spinner(f"🔍 正在搜索关键词: {search_keyword}"):
                            search_results = st.session_state.document_processor.search_content(
                                search_keyword, 
                                context_lines
                            )
                            
                            if search_results:
                                st.success(f"✅ 找到 {len(search_results)} 个匹配结果")
                                
                                for i, result in enumerate(search_results, 1):
                                    with st.expander(f"📍 结果 {i} - 第{result['line_number']}行", expanded=i <= 3):
                                        st.markdown(f"**匹配内容**: {result['matched_line']}")
                                        st.markdown("**上下文**:")
                                        st.code(result['context'], language='text')
                            else:
                                st.warning(f"❌ 未找到关键词: {search_keyword}")
                    else:
                        st.warning("⚠️ 请输入搜索关键词")
                
                # 批量搜索
                st.subheader("📋 批量关键词搜索")
                
                batch_keywords = st.text_area(
                    "输入多个关键词（每行一个）",
                    placeholder="合同编号\n甲方\n乙方\n金额\n日期",
                    height=100,
                    key="doc_batch_keywords"
                )
                
                if st.button("🔍 批量搜索", use_container_width=True):
                    if batch_keywords.strip():
                        keywords = [kw.strip() for kw in batch_keywords.split('\n') if kw.strip()]
                        
                        if keywords:
                            with st.spinner(f"🔍 正在搜索 {len(keywords)} 个关键词..."):
                                try:
                                    from document_utils import DocumentSearchEngine
                                    search_engine = DocumentSearchEngine(st.session_state.document_processor)
                                    
                                    # 生成搜索报告
                                    search_report = search_engine.generate_search_report(keywords)
                                    
                                    st.subheader("📊 批量搜索报告")
                                    st.markdown(search_report)
                                    
                                except Exception as e:
                                    st.error(f"❌ 批量搜索失败: {str(e)}")
                        else:
                            st.warning("⚠️ 请输入有效的关键词")
                    else:
                        st.warning("⚠️ 请输入关键词")
                
                # 导出搜索结果
                st.subheader("📤 导出功能")
                if st.button("📋 导出完整分析报告", use_container_width=True):
                    try:
                        # 生成导出文件
                        user_exports_dir = session_manager.get_user_workspace(session_id) / "exports"
                        user_exports_dir.mkdir(exist_ok=True)
                        
                        json_file, md_file = st.session_state.document_processor.export_analysis_result(str(user_exports_dir))
                        
                        # 提供下载
                        col_download1, col_download2 = st.columns(2)
                        
                        with col_download1:
                            try:
                                with open(json_file, 'rb') as f:
                                    json_data = f.read()
                                st.download_button(
                                    label="⬇️ 下载JSON数据",
                                    data=json_data,
                                    file_name=os.path.basename(json_file),
                                    mime="application/json",
                                    use_container_width=True
                                )
                            except Exception as e:
                                st.error(f"JSON下载失败: {e}")
                        
                        with col_download2:
                            try:
                                with open(md_file, 'rb') as f:
                                    md_data = f.read()
                                st.download_button(
                                    label="⬇️ 下载分析报告",
                                    data=md_data,
                                    file_name=os.path.basename(md_file),
                                    mime="text/markdown",
                                    use_container_width=True
                                )
                            except Exception as e:
                                st.error(f"报告下载失败: {e}")
                        
                        st.success("✅ 文件导出成功！请点击下载按钮获取文件")
                        
                    except Exception as e:
                        st.error(f"❌ 导出失败: {str(e)}")
            else:
                st.error("❌ 文档处理器不可用，请重新加载页面")
    
    else:
        # 欢迎界面
        if analysis_mode == "📊 Excel分析":
            st.info("👋 欢迎使用AI Excel智能分析工具！请上传Excel文件开始分析。")
        else:
            st.info("👋 欢迎使用AI文档智能分析工具！请上传DOCX或PDF文档开始分析。")
        
        # 功能介绍
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("""
            ### 🔐 多用户特性
            - **数据隔离**: 每个用户拥有独立的工作空间
            - **隐私保护**: 自动清理过期文件
            - **配置管理**: 安全保存个人配置
            - **会话管理**: 智能会话超时机制
            """)
        
        with col2:
            if analysis_mode == "📊 Excel分析":
                st.markdown("""
                ### ⚡ Excel分析功能
                - **AI深度分析**: 智能理解业务数据
                - **代码执行**: 隔离环境处理数据
                - **实时预览**: 多工作表支持
                - **数据导出**: 安全文件管理
                """)
            else:
                st.markdown("""
                ### ⚡ 文档分析功能
                - **智能预览**: MarkItDown清洗格式
                - **结构分析**: 标题层级和字体识别
                - **AI理解**: 深度内容分析
                - **关键词搜索**: 精确查找和上下文提取
                """)
        
        # 系统状态展示
        stats = session_manager.get_session_stats()
        st.subheader("📊 系统状态")
        
        col_stat1, col_stat2, col_stat3 = st.columns(3)
        with col_stat1:
            st.metric("活跃用户", stats['active_sessions'])
        with col_stat2:
            st.metric("处理文件", stats['total_files'])
        with col_stat3:
            st.metric("存储使用", f"{stats['disk_usage_mb']} MB")


if __name__ == "__main__":
    main() 