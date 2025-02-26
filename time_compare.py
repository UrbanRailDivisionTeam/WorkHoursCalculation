import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from pathlib import Path

class ExcelCompareTool:
    def __init__(self, root):
        self.root = root
        self.root.title("工序工时对比工具 (以EAS系统数据为基准)")
        self.root.geometry("1200x700")
        
        # 存储文件路径
        self.eas_file_path = None  # EAS系统数据文件
        self.compare_file_path = None  # 待比较文件
        
        # 需要比较的列
        self.required_columns = [
            '项目号', '节车号', '工序编码', '工序名称', 
            '工位', '工序工时'
        ]
        
        # 创建界面
        self.create_widgets()
        
    def create_widgets(self):
        # 文件选择区域
        file_frame = ttk.LabelFrame(self.root, text="文件选择", padding=10)
        file_frame.pack(fill="x", padx=10, pady=5)
        
        # EAS系统数据文件
        ttk.Button(file_frame, text="选择EAS系统数据文件（基准值）", 
                  command=self.select_eas_file).grid(row=0, column=0, padx=5)
        self.eas_file_label = ttk.Label(file_frame, text="未选择文件")
        self.eas_file_label.grid(row=0, column=1, padx=5)
        
        # 待比较文件
        ttk.Button(file_frame, text="选择待比较Excel文件", 
                  command=self.select_compare_file).grid(row=1, column=0, padx=5, pady=5)
        self.compare_file_label = ttk.Label(file_frame, text="未选择文件")
        self.compare_file_label.grid(row=1, column=1, padx=5)
        
        # 对比按钮
        ttk.Button(self.root, text="对比工序工时", 
                  command=self.compare_files).pack(pady=10)
        
        # 结果显示区域
        result_frame = ttk.LabelFrame(self.root, text="对比结果", padding=10)
        result_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 创建表格显示结果
        self.tree = ttk.Treeview(result_frame, columns=(
            "项目号", "节车号", "工序编码", "工序名称", 
            "工位", "EAS工时", "比较文件工时", "工时差异"
        ), show="headings")
        
        # 设置列标题和宽度
        column_widths = {
            "项目号": 100,
            "节车号": 100,
            "工序编码": 100,
            "工序名称": 200,
            "工位": 80,
            "EAS工时": 100,
            "比较文件工时": 100,
            "工时差异": 100
        }
        
        for col, width in column_widths.items():
            self.tree.heading(col, text=col)
            self.tree.column(col, width=width)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(result_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 统计信息显示
        self.stats_label = ttk.Label(self.root, text="")
        self.stats_label.pack(pady=5)
        
        # 导出按钮
        ttk.Button(self.root, text="导出结果", 
                  command=self.export_results).pack(pady=10)

    def select_eas_file(self):
        self.eas_file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if self.eas_file_path:
            self.eas_file_label.config(text=Path(self.eas_file_path).name)

    def select_compare_file(self):
        self.compare_file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if self.compare_file_path:
            self.compare_file_label.config(text=Path(self.compare_file_path).name)

    def set_cell_color(self, item, column, color):
        """设置单元格文字颜色"""
        self.tree.tag_configure(f'color_{item}_{column}', foreground=color)
        col_idx = self.tree["columns"].index(column)
        current_value = self.tree.set(item, col_idx)
        self.tree.set(item, col_idx, current_value)
        self.tree.item(item, tags=(f'color_{item}_{column}',))

    def compare_files(self):
        if not self.eas_file_path or not self.compare_file_path:
            messagebox.showerror("错误", "请先选择两个Excel文件")
            return
            
        try:
            # 读取Excel文件
            df_eas = pd.read_excel(self.eas_file_path)
            df_compare = pd.read_excel(self.compare_file_path)
            
            # 检查必要的列是否存在
            for col in self.required_columns:
                if col not in df_eas.columns or col not in df_compare.columns:
                    messagebox.showerror("错误", f"文件中缺少必要的列：{col}")
                    return
            
            # 清空现有结果
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            # 存储差异结果
            self.differences = []
            
            # 合并关键字段作为索引
            df_eas['key'] = df_eas['项目号'].astype(str) + '_' + \
                           df_eas['节车号'].astype(str) + '_' + \
                           df_eas['工序编码'].astype(str)
            df_compare['key'] = df_compare['项目号'].astype(str) + '_' + \
                               df_compare['节车号'].astype(str) + '_' + \
                               df_compare['工序编码'].astype(str)
            
            # 找出共同的记录
            common_keys = set(df_eas['key']).intersection(set(df_compare['key']))
            
            # 统计变量
            total_records = len(common_keys)
            changed_records = 0
            
            # 比较工序工时
            for key in common_keys:
                row_eas = df_eas[df_eas['key'] == key].iloc[0]
                row_compare = df_compare[df_compare['key'] == key].iloc[0]
                
                time_eas = float(row_eas['工序工时'])
                time_compare = float(row_compare['工序工时'])
                
                if time_eas != time_compare:
                    changed_records += 1
                    time_diff = time_compare - time_eas
                    
                    diff = {
                        '项目号': row_eas['项目号'],
                        '节车号': row_eas['节车号'],
                        '工序编码': row_eas['工序编码'],
                        '工序名称': row_eas['工序名称'],
                        '工位': row_eas['工位'],
                        'EAS工时': time_eas,
                        '比较文件工时': time_compare,
                        '工时差异': round(time_diff, 2)
                    }
                    self.differences.append(diff)
                    
                    # 插入数据行
                    item_id = self.tree.insert("", "end", values=(
                        diff['项目号'], diff['节车号'], 
                        diff['工序编码'], diff['工序名称'],
                        diff['工位'], diff['EAS工时'],
                        diff['比较文件工时'], diff['工时差异']
                    ))
                    
                    # 设置工时差异列的颜色
                    color = 'red' if time_diff > 0 else 'green'
                    self.set_cell_color(item_id, '工时差异', color)
            
            # 更新统计信息
            stats_text = (
                f"统计信息：\n"
                f"总记录数: {total_records}  "
                f"变化记录数: {changed_records}"
            )
            self.stats_label.config(text=stats_text)
            
            if not self.differences:
                messagebox.showinfo("结果", "两个文件中的工序工时完全相同！")
            
        except Exception as e:
            messagebox.showerror("错误", f"比较过程中出现错误：{str(e)}")

    def export_results(self):
        if not hasattr(self, 'differences') or not self.differences:
            messagebox.showinfo("提示", "没有差异需要导出")
            return
            
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if file_path:
            # 创建DataFrame并导出到Excel
            df_result = pd.DataFrame(self.differences)
            
            # 添加统计信息
            total_records = len(self.tree.get_children())
            
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df_result.to_excel(writer, sheet_name='详细对比结果', index=False)
                
                # 创建统计信息表
                stats_df = pd.DataFrame({
                    '统计项': ['总记录数', '变化记录数'],
                    '数量': [total_records, len(self.differences)]
                })
                stats_df.to_excel(writer, sheet_name='统计信息', index=False)
            
            messagebox.showinfo("成功", "结果已成功导出！")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelCompareTool(root)
    root.mainloop() 