import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from pathlib import Path
import datetime

class ExcelCompareTool:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel工序对比工具")
        self.root.geometry("1200x800")  # 加宽窗口以适应更多列
        
        # 存储文件路径和比较结果
        self.base_file = None
        self.compare_file = None
        self.comparison_result = None
        
        # 添加缺少工序数量统计
        self.missing_count = 0
        
        # 定义所需的列名
        self.required_columns = [
            '项目号', '节车号', '工序编码', '工序名称', '工位', '工序工时'
        ]
        
        # 创建界面
        self.create_widgets()
        
    def create_widgets(self):
        # 文件选择区域
        file_frame = ttk.LabelFrame(self.root, text="文件选择", padding=10)
        file_frame.pack(fill="x", padx=10, pady=5)
        
        # 基准表选择
        ttk.Button(file_frame, text="选择基准表", command=self.select_base_file).grid(row=0, column=0, padx=5)
        self.base_label = ttk.Label(file_frame, text="未选择文件")
        self.base_label.grid(row=0, column=1, padx=5)
        
        # 比较表选择
        ttk.Button(file_frame, text="选择比较表", command=self.select_compare_file).grid(row=1, column=0, padx=5)
        self.compare_label = ttk.Label(file_frame, text="未选择文件")
        self.compare_label.grid(row=1, column=1, padx=5)
        
        # 按钮区域
        button_frame = ttk.Frame(self.root)
        button_frame.pack(pady=10)
        ttk.Button(button_frame, text="开始对比", command=self.compare_files).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="导出结果", command=self.export_results).pack(side=tk.LEFT, padx=5)
        
        # 结果显示区域
        result_frame = ttk.LabelFrame(self.root, text="对比结果", padding=10)
        result_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 使用Treeview显示结果
        self.result_tree = ttk.Treeview(result_frame, columns=(
            '类型',
            '项目号',
            '节车号',
            '工序编码',
            '工序名称',
            '工位',
            '基准表工时',
            '比较表工时'
        ))
        
        # 设置列标题
        self.result_tree.heading('类型', text='变更类型')
        self.result_tree.heading('项目号', text='项目号')
        self.result_tree.heading('节车号', text='节车号')
        self.result_tree.heading('工序编码', text='工序编码')
        self.result_tree.heading('工序名称', text='工序名称')
        self.result_tree.heading('工位', text='工位')
        self.result_tree.heading('基准表工时', text='基准表工时')
        self.result_tree.heading('比较表工时', text='比较表工时')
        
        # 设置列宽
        self.result_tree.column('类型', width=100)
        self.result_tree.column('项目号', width=100)
        self.result_tree.column('节车号', width=100)
        self.result_tree.column('工序编码', width=100)
        self.result_tree.column('工序名称', width=200)
        self.result_tree.column('工位', width=100)
        self.result_tree.column('基准表工时', width=100)
        self.result_tree.column('比较表工时', width=100)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(result_frame, orient="vertical", command=self.result_tree.yview)
        self.result_tree.configure(yscrollcommand=scrollbar.set)
        
        self.result_tree.pack(side=tk.LEFT, fill="both", expand=True)
        scrollbar.pack(side=tk.RIGHT, fill="y")
        
        # 添加统计信息显示区域
        stats_frame = ttk.LabelFrame(self.root, text="统计信息", padding=10)
        stats_frame.pack(fill="x", padx=10, pady=5)
        
        self.stats_label = ttk.Label(stats_frame, text="")
        self.stats_label.pack(fill="x")
        
    def select_base_file(self):
        filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filename:
            self.base_file = filename
            self.base_label.config(text=Path(filename).name)
            
    def select_compare_file(self):
        filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filename:
            self.compare_file = filename
            self.compare_label.config(text=Path(filename).name)
    
    def validate_dataframe(self, df, file_name):
        missing_columns = [col for col in self.required_columns if col not in df.columns]
        if missing_columns:
            raise ValueError(f"{file_name}中缺少必要的列：{', '.join(missing_columns)}")
    
    def compare_files(self):
        if not self.base_file or not self.compare_file:
            messagebox.showerror("错误", "请先选择两个Excel文件")
            return
            
        try:
            # 重置统计数据
            self.missing_count = 0
            
            # 清空现有结果
            for item in self.result_tree.get_children():
                self.result_tree.delete(item)
            
            # 读取Excel文件
            df_base = pd.read_excel(self.base_file)
            df_compare = pd.read_excel(self.compare_file)
            
            # 验证数据框架是否包含所需列
            self.validate_dataframe(df_base, "基准表")
            self.validate_dataframe(df_compare, "比较表")
            
            # 创建用于比较的键
            df_base['比较键'] = df_base['项目号'].astype(str) + '_' + df_base['节车号'].astype(str)
            df_compare['比较键'] = df_compare['项目号'].astype(str) + '_' + df_compare['节车号'].astype(str)
            
            # 找出所有唯一的比较键
            all_keys = set(df_base['比较键'].unique()) | set(df_compare['比较键'].unique())
            
            # 存储比较结果
            self.comparison_result = []
            
            for key in all_keys:
                base_rows = df_base[df_base['比较键'] == key]
                compare_rows = df_compare[df_compare['比较键'] == key]
                
                if len(compare_rows) == 0:  # 比较表中完全缺少的项目和节车号组合
                    for _, row in base_rows.iterrows():
                        values = [
                            "缺少",
                            row['项目号'],
                            row['节车号'],
                            row['工序编码'],
                            row['工序名称'],
                            row['工位'],
                            row['工序工时'],
                            '-'  # 比较表中不存在此工序
                        ]
                        self.result_tree.insert('', 'end', values=values)
                        self.missing_count += 1
                else:  # 比较工序编码
                    base_processes = set(base_rows['工序编码'])
                    compare_processes = set(compare_rows['工序编码'])
                    
                    # 缺少的工序
                    missing_processes = base_processes - compare_processes
                    for _, row in base_rows[base_rows['工序编码'].isin(missing_processes)].iterrows():
                        values = [
                            "缺少工序",
                            row['项目号'],
                            row['节车号'],
                            row['工序编码'],
                            row['工序名称'],
                            row['工位'],
                            row['工序工时'],
                            '-'  # 比较表中不存在此工序
                        ]
                        self.result_tree.insert('', 'end', values=values)
                        self.missing_count += 1
                    
                    # 检查工序工时变化
                    common_processes = base_processes & compare_processes
                    for process in common_processes:
                        base_row = base_rows[base_rows['工序编码'] == process].iloc[0]
                        compare_row = compare_rows[compare_rows['工序编码'] == process].iloc[0]
                        
                        if base_row['工序工时'] != compare_row['工序工时']:
                            values = [
                                "工时变化",
                                base_row['项目号'],
                                base_row['节车号'],
                                base_row['工序编码'],
                                base_row['工序名称'],
                                base_row['工位'],
                                base_row['工序工时'],
                                compare_row['工序工时']
                            ]
                            self.result_tree.insert('', 'end', values=values)
            
            # 更新统计信息显示
            self.update_stats_display()
            
            messagebox.showinfo("完成", "对比完成！")
            
        except Exception as e:
            messagebox.showerror("错误", f"处理文件时出错：{str(e)}")
    
    def export_results(self):
        if not self.result_tree.get_children():
            messagebox.showwarning("警告", "没有可导出的结果！")
            return
            
        try:
            # 生成默认文件名
            default_filename = f"工序对比结果_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            
            # 选择保存位置
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile=default_filename
            )
            
            if filename:
                # 获取所有结果
                results = []
                for item in self.result_tree.get_children():
                    results.append(self.result_tree.item(item)['values'])
                
                # 创建DataFrame并保存
                df_result = pd.DataFrame(
                    results,
                    columns=[
                        '变更类型',
                        '项目号',
                        '节车号',
                        '工序编码',
                        '工序名称',
                        '工位',
                        '基准表工时',
                        '比较表工时'
                    ]
                )
                df_result.to_excel(filename, index=False)
                messagebox.showinfo("成功", "结果已成功导出！")
                
        except Exception as e:
            messagebox.showerror("错误", f"导出结果时出错：{str(e)}")

    def update_stats_display(self):
        """更新统计信息显示"""
        stats_text = f"比较文件中缺少的工序数量: {self.missing_count} 个"
        self.stats_label.config(text=stats_text)

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelCompareTool(root)
    root.mainloop() 