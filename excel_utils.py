import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
import numpy as np
from typing import Dict, List, Any, Tuple, Optional
import io
import tempfile
import os

class AdvancedExcelProcessor:
    """高级Excel处理类"""
    
    def __init__(self):
        self.workbook = None
        self.file_path = None
        self.modified_data = {}
    
    def load_excel(self, file_path: str) -> Dict[str, pd.DataFrame]:
        """加载Excel文件并返回所有工作表的数据"""
        self.file_path = file_path
        self.workbook = openpyxl.load_workbook(file_path, data_only=False)
        
        excel_data = {}
        for sheet_name in self.workbook.sheetnames:
            df, _ = self.read_excel_with_merged_cells(file_path, sheet_name)
            excel_data[sheet_name] = df
            self.modified_data[sheet_name] = df.copy()
        
        return excel_data
    
    @staticmethod
    def read_excel_with_merged_cells(file_path: str, sheet_name: str = None, max_rows: int = 1000) -> Tuple[pd.DataFrame, str]:
        """读取Excel文件并处理合并单元格，支持更多行数"""
        try:
            wb = openpyxl.load_workbook(file_path, data_only=True)
            
            if sheet_name is None:
                sheet_name = wb.sheetnames[0]
            
            ws = wb[sheet_name]
            
            # 获取合并单元格信息
            merged_cells = ws.merged_cells.ranges
            merged_dict = {}
            
            # 创建合并单元格映射
            for merged_range in merged_cells:
                top_left_cell = ws.cell(row=merged_range.min_row, column=merged_range.min_col)
                top_left_value = top_left_cell.value
                
                for row in range(merged_range.min_row, merged_range.max_row + 1):
                    for col in range(merged_range.min_col, merged_range.max_col + 1):
                        merged_dict[(row, col)] = top_left_value
            
            # 读取数据
            data = []
            max_col = min(ws.max_column, 100)  # 限制最大列数
            max_row = min(ws.max_row, max_rows)
            
            for row in range(1, max_row + 1):
                row_data = []
                for col in range(1, max_col + 1):
                    # 检查是否是合并单元格
                    if (row, col) in merged_dict:
                        value = merged_dict[(row, col)]
                    else:
                        cell = ws.cell(row=row, column=col)
                        value = cell.value
                    
                    row_data.append(value if value is not None else "")
                
                data.append(row_data)
            
            # 转换为DataFrame
            if data and len(data) > 1:
                # 使用第一行作为列名，确保列名是字符串
                columns = []
                column_counts = {}  # 用于跟踪重复列名
                
                for i, col in enumerate(data[0]):
                    if col is None or str(col).strip() == "":
                        base_name = f"列{i+1}"
                    else:
                        base_name = str(col).strip()
                    
                    # 处理重复列名
                    if base_name in column_counts:
                        column_counts[base_name] += 1
                        unique_name = f"{base_name}_{column_counts[base_name]}"
                    else:
                        column_counts[base_name] = 0
                        unique_name = base_name
                    
                    columns.append(unique_name)
                
                # 创建DataFrame，确保数据行存在
                data_rows = data[1:] if len(data) > 1 else []
                if data_rows:
                    df = pd.DataFrame(data_rows, columns=columns)
                else:
                    df = pd.DataFrame(columns=columns)
            else:
                df = pd.DataFrame()
            
            # 清理数据 - 添加错误处理
            try:
                if not df.empty:
                    # 修复pandas FutureWarning并确保数据类型一致
                    df = df.replace('', np.nan)
                    
                    # 对每列进行数据类型规范化，避免混合类型
                    for col in df.columns:
                        try:
                            # 检查列是否包含混合类型
                            if df[col].dtype == 'object':
                                # 尝试将全部转换为字符串，避免混合类型
                                df[col] = df[col].astype(str)
                                # 将字符串'nan'重新替换为真正的NaN
                                df[col] = df[col].replace('nan', np.nan)
                        except Exception as e:
                            print(f"列 {col} 数据类型转换警告: {e}")
                            # 如果转换失败，强制转换为字符串
                            df[col] = df[col].astype(str).replace('nan', np.nan)
                    
                    # 最后应用infer_objects
                    df = df.infer_objects(copy=False)
                    
                    # 清理列名中的特殊字符
                    df.columns = [str(col).replace('\n', ' ').replace('\r', ' ').strip() for col in df.columns]
            except Exception as e:
                print(f"数据清理时出现问题: {e}")
                # 如果清理失败，至少确保基本的数据结构
                if not df.empty:
                    try:
                        df.columns = [str(col).replace('\n', ' ').replace('\r', ' ').strip() for col in df.columns]
                    except:
                        pass
            
            # 生成markdown格式预览
            markdown_content = AdvancedExcelProcessor.df_to_markdown(df, sheet_name)
            
            return df, markdown_content
            
        except Exception as e:
            raise Exception(f"读取Excel文件时出错: {str(e)}")
    
    @staticmethod
    def df_to_markdown(df: pd.DataFrame, sheet_name: str = "") -> str:
        """将DataFrame转换为Markdown格式，支持更好的格式化"""
        if df.empty:
            return f"## 📋 {sheet_name}\n\n*此工作表为空*\n\n"
        
        markdown = f"## 📋 {sheet_name}\n\n"
        
        # 数据基本信息
        markdown += f"**数据概览**: {len(df)} 行 × {len(df.columns)} 列\n\n"
        
        # 限制显示的行数和列数
        display_rows = min(20, len(df))
        display_cols = min(8, len(df.columns))
        display_df = df.head(display_rows).iloc[:, :display_cols]
        
        # 处理列名，确保不会太长
        headers = ["#"] + [str(col)[:15] + "..." if len(str(col)) > 15 else str(col) for col in display_df.columns]
        markdown += "| " + " | ".join(headers) + " |\n"
        markdown += "| " + " | ".join(["---"] * len(headers)) + " |\n"
        
        # 添加数据行
        for idx, row in display_df.iterrows():
            row_data = [str(idx + 1)]  # 行号从1开始
            for val in row:
                str_val = str(val) if pd.notna(val) else ""
                # 限制单元格内容长度
                if len(str_val) > 20:
                    str_val = str_val[:17] + "..."
                row_data.append(str_val)
            markdown += "| " + " | ".join(row_data) + " |\n"
        
        # 添加省略信息
        if len(df) > display_rows:
            markdown += f"\n*📝 仅显示前{display_rows}行，总共{len(df)}行*\n"
        
        if len(df.columns) > display_cols:
            markdown += f"*📊 仅显示前{display_cols}列，总共{len(df.columns)}列*\n"
        
        # 添加数据类型信息 - 修复错误处理
        markdown += "\n### 📈 数据类型概览\n"
        try:
            cols_to_show = list(df.columns)[:5]  # 只显示前5列的类型
            for col in cols_to_show:
                try:
                    if col in df.columns:
                        dtype = str(df[col].dtype)
                        non_null_count = df[col].count()
                        markdown += f"- **{col}**: {dtype} ({non_null_count}/{len(df)} 非空)\n"
                except Exception as e:
                    # 如果获取某列的数据类型失败，跳过该列
                    markdown += f"- **{col}**: 未知类型\n"
            
            if len(df.columns) > 5:
                markdown += f"- *... 还有 {len(df.columns) - 5} 列*\n"
        except Exception as e:
            markdown += "- *数据类型信息获取失败*\n"
        
        return markdown
    
    def update_dataframe(self, sheet_name: str, df: pd.DataFrame):
        """更新指定工作表的数据"""
        self.modified_data[sheet_name] = df
    
    def add_calculated_column(self, sheet_name: str, column_name: str, formula_func, description: str = ""):
        """添加计算列"""
        if sheet_name in self.modified_data:
            df = self.modified_data[sheet_name]
            try:
                df[column_name] = formula_func(df)
                self.modified_data[sheet_name] = df
                return True, f"成功添加计算列: {column_name}"
            except Exception as e:
                return False, f"添加计算列失败: {str(e)}"
        return False, "工作表不存在"
    
    def fill_missing_values(self, sheet_name: str, column: str, method: str = "mean", custom_value=None):
        """填充缺失值"""
        if sheet_name not in self.modified_data:
            return False, "工作表不存在"
        
        df = self.modified_data[sheet_name]
        if column not in df.columns:
            return False, f"列 '{column}' 不存在"
        
        try:
            if method == "mean" and pd.api.types.is_numeric_dtype(df[column]):
                df[column].fillna(df[column].mean(), inplace=True)
            elif method == "median" and pd.api.types.is_numeric_dtype(df[column]):
                df[column].fillna(df[column].median(), inplace=True)
            elif method == "mode":
                mode_value = df[column].mode().iloc[0] if not df[column].mode().empty else ""
                df[column].fillna(mode_value, inplace=True)
            elif method == "forward":
                df[column].fillna(method='ffill', inplace=True)
            elif method == "backward":
                df[column].fillna(method='bfill', inplace=True)
            elif method == "custom" and custom_value is not None:
                df[column].fillna(custom_value, inplace=True)
            else:
                return False, f"不支持的填充方法: {method}"
            
            self.modified_data[sheet_name] = df
            return True, f"成功填充列 '{column}' 的缺失值"
            
        except Exception as e:
            return False, f"填充失败: {str(e)}"
    
    def add_summary_statistics(self, sheet_name: str, target_columns: List[str] = None):
        """添加汇总统计"""
        if sheet_name not in self.modified_data:
            return False, "工作表不存在"
        
        df = self.modified_data[sheet_name]
        
        try:
            if target_columns is None:
                target_columns = df.select_dtypes(include=[np.number]).columns.tolist()
            
            if not target_columns:
                return False, "没有找到数值列"
            
            # 创建汇总统计
            summary_data = []
            for col in target_columns:
                if col in df.columns and pd.api.types.is_numeric_dtype(df[col]):
                    stats = {
                        '列名': col,
                        '计数': df[col].count(),
                        '均值': df[col].mean(),
                        '中位数': df[col].median(),
                        '标准差': df[col].std(),
                        '最小值': df[col].min(),
                        '最大值': df[col].max(),
                        '缺失值': df[col].isnull().sum()
                    }
                    summary_data.append(stats)
            
            if summary_data:
                summary_df = pd.DataFrame(summary_data)
                summary_sheet_name = f"{sheet_name}_统计汇总"
                self.modified_data[summary_sheet_name] = summary_df
                return True, f"成功创建统计汇总工作表: {summary_sheet_name}"
            else:
                return False, "没有可统计的数值数据"
                
        except Exception as e:
            return False, f"创建统计汇总失败: {str(e)}"
    
    def export_to_excel(self, output_path: str = None) -> str:
        """导出修改后的数据到Excel文件"""
        try:
            if output_path is None:
                timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
                output_path = f"modified_excel_{timestamp}.xlsx"
            
            # 创建新的工作簿
            wb = openpyxl.Workbook()
            wb.remove(wb.active)  # 删除默认工作表
            
            for sheet_name, df in self.modified_data.items():
                ws = wb.create_sheet(title=sheet_name)
                
                # 写入列标题
                for col_idx, column in enumerate(df.columns, 1):
                    cell = ws.cell(row=1, column=col_idx, value=column)
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(start_color="E6E6FA", end_color="E6E6FA", fill_type="solid")
                
                # 写入数据
                for row_idx, row in df.iterrows():
                    for col_idx, value in enumerate(row, 1):
                        ws.cell(row=row_idx + 2, column=col_idx, value=value)
                
                # 自动调整列宽
                for column in ws.columns:
                    max_length = 0
                    column_letter = get_column_letter(column[0].column)
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    ws.column_dimensions[column_letter].width = adjusted_width
            
            # 保存文件
            wb.save(output_path)
            return output_path
            
        except Exception as e:
            raise Exception(f"导出Excel文件失败: {str(e)}")
    
    def get_data_preview(self, sheet_name: str) -> str:
        """获取数据预览"""
        if sheet_name in self.modified_data:
            df = self.modified_data[sheet_name]
            return self.df_to_markdown(df, sheet_name)
        return "工作表不存在"

class DataAnalyzer:
    """数据分析类"""
    
    @staticmethod
    def detect_data_types(df: pd.DataFrame) -> Dict[str, str]:
        """检测数据类型"""
        type_info = {}
        try:
            for col in df.columns:
                try:
                    if col in df.columns and not df[col].empty:
                        if pd.api.types.is_numeric_dtype(df[col]):
                            if df[col].dtype == 'int64':
                                type_info[col] = "整数"
                            else:
                                type_info[col] = "浮点数"
                        elif pd.api.types.is_datetime64_any_dtype(df[col]):
                            type_info[col] = "日期时间"
                        elif pd.api.types.is_bool_dtype(df[col]):
                            type_info[col] = "布尔值"
                        else:
                            type_info[col] = "文本"
                    else:
                        type_info[col] = "未知"
                except Exception as e:
                    type_info[col] = "检测失败"
        except Exception as e:
            print(f"数据类型检测出错: {e}")
        return type_info
    
    @staticmethod
    def find_duplicates(df: pd.DataFrame) -> pd.DataFrame:
        """查找重复行"""
        try:
            if df.empty:
                return pd.DataFrame()
            duplicates = df[df.duplicated(keep=False)]
            return duplicates
        except Exception as e:
            print(f"查找重复数据时出错: {e}")
            return pd.DataFrame()
    
    @staticmethod
    def get_missing_value_report(df: pd.DataFrame) -> pd.DataFrame:
        """获取缺失值报告"""
        missing_data = []
        try:
            for col in df.columns:
                try:
                    missing_count = df[col].isnull().sum()
                    missing_percent = (missing_count / len(df)) * 100 if len(df) > 0 else 0
                    missing_data.append({
                        '列名': str(col),
                        '缺失数量': missing_count,
                        '缺失百分比': f"{missing_percent:.2f}%",
                        '数据类型': str(df[col].dtype) if col in df.columns else "未知"
                    })
                except Exception as e:
                    missing_data.append({
                        '列名': str(col),
                        '缺失数量': 0,
                        '缺失百分比': "0.00%",
                        '数据类型': "检测失败"
                    })
        except Exception as e:
            print(f"获取缺失值报告时出错: {e}")
        
        return pd.DataFrame(missing_data)
    
    @staticmethod
    def detect_outliers(df: pd.DataFrame, method: str = "iqr") -> Dict[str, List]:
        """检测异常值"""
        outliers = {}
        try:
            numeric_columns = df.select_dtypes(include=[np.number]).columns
            
            for col in numeric_columns:
                try:
                    if method == "iqr":
                        Q1 = df[col].quantile(0.25)
                        Q3 = df[col].quantile(0.75)
                        IQR = Q3 - Q1
                        lower_bound = Q1 - 1.5 * IQR
                        upper_bound = Q3 + 1.5 * IQR
                        outlier_indices = df[(df[col] < lower_bound) | (df[col] > upper_bound)].index.tolist()
                        outliers[col] = outlier_indices
                    elif method == "zscore":
                        try:
                            from scipy import stats
                            z_scores = np.abs(stats.zscore(df[col].dropna()))
                            outlier_indices = df[col].dropna().index[z_scores > 3].tolist()
                            outliers[col] = outlier_indices
                        except ImportError:
                            # 如果scipy不可用，回退到IQR方法
                            Q1 = df[col].quantile(0.25)
                            Q3 = df[col].quantile(0.75)
                            IQR = Q3 - Q1
                            lower_bound = Q1 - 1.5 * IQR
                            upper_bound = Q3 + 1.5 * IQR
                            outlier_indices = df[(df[col] < lower_bound) | (df[col] > upper_bound)].index.tolist()
                            outliers[col] = outlier_indices
                except Exception as e:
                    print(f"检测列 {col} 的异常值时出错: {e}")
                    outliers[col] = []
        except Exception as e:
            print(f"异常值检测出错: {e}")
        
        return outliers 