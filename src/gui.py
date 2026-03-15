import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import json
import sys

# 添加当前目录到路径
sys.path.append(os.path.dirname(__file__))

from excel_selector import ExcelSelector
from config import ConfigManager

class ExcelSelectorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 指标勾选文件处理工具")
        self.root.geometry("800x800")
        
        # 配置管理器
        self.config_manager = ConfigManager()
        
        # 配置变量
        self.config = {
            'source_file': '',
            'target_file': '',
            'source_config': {
                'category_row': 1,
                'indicator_start_row': 3
            },
            'target_config': {
                'indicator_start_row': 3
            },
            'category_mapping': {},
            'month_config': {
                'selected_months': []
            }
        }
        
        # 产品列表
        self.source_products = []
        self.target_sheets = []
        
        # 创建主框架
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建配置管理区域
        self.create_config_management()
        
        # 创建文件选择区域
        self.create_file_selection()
        
        # 创建产品行配置区域
        self.create_product_row_config()
        
        # 创建中间内容区域（产品映射和月份选择左右分布）
        middle_frame = ttk.Frame(self.main_frame)
        middle_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 左侧：产品映射区域
        mapping_frame = ttk.LabelFrame(middle_frame, text="产品映射", padding="10")
        mapping_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        
        # 创建滚动条
        self.mapping_canvas = tk.Canvas(mapping_frame)
        self.mapping_scrollbar = ttk.Scrollbar(mapping_frame, orient=tk.VERTICAL, command=self.mapping_canvas.yview)
        self.mapping_scrollable_frame = ttk.Frame(self.mapping_canvas)
        
        self.mapping_scrollable_frame.bind(
            "<Configure>",
            lambda e: self.mapping_canvas.configure(
                scrollregion=self.mapping_canvas.bbox("all")
            )
        )
        
        self.mapping_canvas.create_window((0, 0), window=self.mapping_scrollable_frame, anchor="nw")
        self.mapping_canvas.configure(yscrollcommand=self.mapping_scrollbar.set)
        
        self.mapping_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.mapping_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 初始提示
        ttk.Label(self.mapping_scrollable_frame, text="请先加载参考文件和目标文件，然后点击'加载产品'").pack(pady=20)
        
        # 右侧：月份选择区域
        month_frame = ttk.LabelFrame(middle_frame, text="月份选择", padding="10")
        month_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5)
        
        self.month_vars = []
        for i in range(12):
            var = tk.IntVar()
            self.month_vars.append(var)
            ttk.Checkbutton(month_frame, text=f"{i+1}月", variable=var).grid(row=i//3, column=i%3, padx=10, pady=5, sticky=tk.W)
        
        # 创建指标开始行配置区域
        self.create_indicator_start_row_config()
        
        # 创建操作按钮区域
        self.create_action_buttons()
        
        # 配置修改监听
        self.setup_config_listeners()
        
        # 自动加载上次配置
        self.auto_load_last_config()
    
    def create_file_selection(self):
        """创建文件选择区域"""
        file_frame = ttk.LabelFrame(self.main_frame, text="文件配置", padding="10")
        file_frame.pack(fill=tk.X, pady=5)
        
        # 参考文件选择
        ttk.Label(file_frame, text="参考文件:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.source_file_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.source_file_var, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="浏览", command=self.browse_source_file).grid(row=0, column=2, padx=5, pady=5)
        
        # 目标文件选择
        ttk.Label(file_frame, text="目标文件:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.target_file_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.target_file_var, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="浏览", command=self.browse_target_file).grid(row=1, column=2, padx=5, pady=5)
    
    def create_product_row_config(self):
        """创建产品行配置区域"""
        product_row_frame = ttk.LabelFrame(self.main_frame, text="产品行配置", padding="10")
        product_row_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(product_row_frame, text="参考文件产品行 (Excel行号):").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.category_row_var = tk.StringVar(value="1")
        ttk.Entry(product_row_frame, textvariable=self.category_row_var, width=10).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(product_row_frame, text="加载产品", command=self.load_products).grid(row=0, column=2, padx=5, pady=5)
        # 加载产品状态
        self.product_status_var = tk.StringVar(value="")
        ttk.Label(product_row_frame, textvariable=self.product_status_var, foreground="green").grid(row=0, column=3, padx=5, pady=5)
    
    def create_product_mapping(self):
        """创建产品映射区域"""
        mapping_frame = ttk.LabelFrame(self.main_frame, text="产品映射", padding="10")
        mapping_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 创建滚动条
        self.mapping_canvas = tk.Canvas(mapping_frame)
        self.mapping_scrollbar = ttk.Scrollbar(mapping_frame, orient=tk.VERTICAL, command=self.mapping_canvas.yview)
        self.mapping_scrollable_frame = ttk.Frame(self.mapping_canvas)
        
        self.mapping_scrollable_frame.bind(
            "<Configure>",
            lambda e: self.mapping_canvas.configure(
                scrollregion=self.mapping_canvas.bbox("all")
            )
        )
        
        self.mapping_canvas.create_window((0, 0), window=self.mapping_scrollable_frame, anchor="nw")
        self.mapping_canvas.configure(yscrollcommand=self.mapping_scrollbar.set)
        
        self.mapping_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.mapping_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 初始提示
        ttk.Label(self.mapping_scrollable_frame, text="请先加载参考文件和目标文件，然后点击'加载产品'").pack(pady=20)
    
    def create_indicator_start_row_config(self):
        """创建指标开始行配置区域"""
        indicator_frame = ttk.LabelFrame(self.main_frame, text="指标配置", padding="10")
        indicator_frame.pack(fill=tk.X, pady=5)
        
        # 左右分布
        left_frame = ttk.Frame(indicator_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        right_frame = ttk.Frame(indicator_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=5)
        
        # 左侧：指标开始行配置
        ttk.Label(left_frame, text="参考文件指标开始行 (Excel行号):").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.source_indicator_row_var = tk.StringVar(value="3")
        ttk.Entry(left_frame, textvariable=self.source_indicator_row_var, width=10).grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(left_frame, text="目标文件指标开始行 (Excel行号):").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.target_indicator_row_var = tk.StringVar(value="3")
        ttk.Entry(left_frame, textvariable=self.target_indicator_row_var, width=10).grid(row=1, column=1, padx=5, pady=5)
        
        # 右侧：标记配置
        ttk.Label(right_frame, text="源文件标记:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.source_mark_var = tk.StringVar(value="✓")
        ttk.Entry(right_frame, textvariable=self.source_mark_var, width=10).grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(right_frame, text="目标文件标记:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.target_mark_var = tk.StringVar(value="✓")
        ttk.Entry(right_frame, textvariable=self.target_mark_var, width=10).grid(row=1, column=1, padx=5, pady=5)
    
    def create_month_selection(self):
        """创建月份选择区域"""
        month_frame = ttk.LabelFrame(self.main_frame, text="月份选择", padding="10")
        month_frame.pack(fill=tk.X, pady=5)
        
        self.month_vars = []
        for i in range(12):
            var = tk.IntVar()
            self.month_vars.append(var)
            ttk.Checkbutton(month_frame, text=f"{i+1}月", variable=var).grid(row=i//4, column=i%4, padx=10, pady=5, sticky=tk.W)
    
    def create_config_management(self):
        """创建配置管理区域"""
        config_frame = ttk.LabelFrame(self.main_frame, text="配置管理", padding="10")
        config_frame.pack(fill=tk.X, pady=5)
        
        # 配置下拉框
        ttk.Label(config_frame, text="配置:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.config_var = tk.StringVar()
        self.config_combobox = ttk.Combobox(config_frame, textvariable=self.config_var, width=30)
        self.config_combobox.grid(row=0, column=1, padx=5, pady=5)
        
        # 加载配置按钮和状态
        ttk.Button(config_frame, text="加载配置", command=self.load_config_from_combobox, width=15).grid(row=0, column=2, padx=5, pady=5)
        self.config_status_var = tk.StringVar(value="")
        ttk.Label(config_frame, textvariable=self.config_status_var, foreground="green").grid(row=0, column=3, padx=5, pady=5)
        
        # 保存配置按钮
        ttk.Button(config_frame, text="保存配置", command=self.save_config, width=15).grid(row=0, column=4, padx=5, pady=5)
        
        # 刷新配置列表
        self.refresh_config_list()
    
    def create_action_buttons(self):
        """创建操作按钮区域"""
        action_frame = ttk.Frame(self.main_frame, padding="10")
        action_frame.pack(fill=tk.X, pady=5)
        
        # 执行按钮和状态
        execute_frame = ttk.Frame(action_frame)
        execute_frame.pack(side=tk.LEFT, padx=10)
        ttk.Button(execute_frame, text="执行", command=self.execute, width=20).pack(side=tk.LEFT)
        self.execute_status_var = tk.StringVar(value="")
        ttk.Label(execute_frame, textvariable=self.execute_status_var, foreground="green").pack(side=tk.LEFT, padx=5)
        
        ttk.Button(action_frame, text="查看结果", command=self.view_result, width=20).pack(side=tk.LEFT, padx=10)
    
    def browse_source_file(self):
        """浏览参考文件"""
        file_path = filedialog.askopenfilename(
            title="选择参考文件",
            filetypes=[("Excel文件", "*.xlsx")]
        )
        if file_path:
            self.source_file_var.set(file_path)
    
    def browse_target_file(self):
        """浏览目标文件"""
        # 默认打开workspace文件夹
        workspace_dir = os.path.join(os.path.dirname(__file__), "..", "test")
        if not os.path.exists(workspace_dir):
            workspace_dir = os.path.dirname(__file__)
        
        file_path = filedialog.askopenfilename(
            title="选择目标文件",
            initialdir=workspace_dir,
            filetypes=[("Excel文件", "*.xlsx")]
        )
        if file_path:
            self.target_file_var.set(file_path)
    
    def view_result(self):
        """查看结果文件"""
        target_file = self.target_file_var.get()
        if target_file and os.path.exists(target_file):
            os.system(f"open '{target_file}'")  # macOS
        else:
            messagebox.showerror("错误", "请先选择目标文件")
    
    def setup_config_listeners(self):
        """设置配置修改监听"""
        # 监听文件选择
        self.source_file_var.trace_add('write', lambda *args: self.config_status_var.set(""))
        self.target_file_var.trace_add('write', lambda *args: self.config_status_var.set(""))
        
        # 监听产品行配置
        self.category_row_var.trace_add('write', lambda *args: self.config_status_var.set(""))
        
        # 监听指标行配置
        self.source_indicator_row_var.trace_add('write', lambda *args: self.config_status_var.set(""))
        self.target_indicator_row_var.trace_add('write', lambda *args: self.config_status_var.set(""))
        
        # 监听标记配置
        self.source_mark_var.trace_add('write', lambda *args: self.config_status_var.set(""))
        self.target_mark_var.trace_add('write', lambda *args: self.config_status_var.set(""))
        
        # 监听月份选择
        for var in self.month_vars:
            var.trace_add('write', lambda *args: self.config_status_var.set(""))
    
    def load_products(self):
        """加载参考文件和目标文件的产品信息"""
        source_file = self.source_file_var.get()
        target_file = self.target_file_var.get()
        
        if not source_file or not target_file:
            self.product_status_var.set("❌")
            return
        
        try:
            # 读取参考文件的产品信息
            import pandas as pd
            df = pd.read_excel(source_file, sheet_name='Sheet1')
            category_row = int(self.category_row_var.get()) - 1
            # 加载所有产品（从第1列开始）
            all_category_cols = df.columns[1:]
            self.source_products = df.iloc[category_row, 1:].tolist()
            
            # 读取目标文件的sheet页信息
            xls = pd.ExcelFile(target_file)
            self.target_sheets = xls.sheet_names
            
            # 更新产品映射区域
            self.update_product_mapping()
            
            self.product_status_var.set("✅")
        except Exception as e:
            print(f"加载产品失败: {str(e)}")
            self.product_status_var.set("❌")
    
    def update_product_mapping(self):
        """更新产品映射区域"""
        # 清空现有内容
        for widget in self.mapping_scrollable_frame.winfo_children():
            widget.destroy()
        
        # 创建产品映射表格
        for i, product in enumerate(self.source_products):
            if product:
                ttk.Label(self.mapping_scrollable_frame, text=product).grid(row=i, column=0, padx=5, pady=5, sticky=tk.W)
                # 创建下拉框
                var = tk.StringVar()
                # 尝试从配置中加载映射
                if product in self.config['category_mapping']:
                    target_product = self.config['category_mapping'][product]
                    if target_product in self.target_sheets:
                        var.set(target_product)
                # 尝试自动匹配
                elif product in self.target_sheets:
                    var.set(product)
                combobox = ttk.Combobox(self.mapping_scrollable_frame, textvariable=var, values=self.target_sheets, width=20)
                combobox.grid(row=i, column=1, padx=5, pady=5)
                # 保存映射关系
                self.config['category_mapping'][product] = var
    
    def execute(self):
        """执行映射操作"""
        # 收集配置
        self.config['source_file'] = self.source_file_var.get()
        self.config['target_file'] = self.target_file_var.get()
        self.config['source_config']['category_row'] = int(self.category_row_var.get())
        self.config['source_config']['indicator_start_row'] = int(self.source_indicator_row_var.get())
        self.config['target_config']['indicator_start_row'] = int(self.target_indicator_row_var.get())
        
        # 收集产品映射，只包含有选择目标产品的映射
        filtered_mapping = {}
        for product, var in self.config['category_mapping'].items():
            target_product = var.get()
            if target_product:
                filtered_mapping[product] = target_product
        # 保存原始映射，执行后不会清空
        original_mapping = {}
        for product, var in self.config['category_mapping'].items():
            original_mapping[product] = var
        
        # 检查是否有多个参考产品映射到同一个目标产品
        target_to_sources = {}
        for source, target in filtered_mapping.items():
            if target not in target_to_sources:
                target_to_sources[target] = []
            target_to_sources[target].append(source)
        
        # 检查是否有重复映射
        duplicate_mappings = []
        for target, sources in target_to_sources.items():
            if len(sources) > 1:
                duplicate_mappings.append((target, sources))
        
        if duplicate_mappings:
            error_msg = "发现重复的产品映射：\n"
            for target, sources in duplicate_mappings:
                error_msg += f"目标产品 '{target}' 被以下参考产品映射：{', '.join(sources)}\n"
            messagebox.showerror("错误", error_msg + "请修改产品映射，确保每个目标产品只被一个参考产品映射")
            return
        
        # 收集选中的月份
        selected_months = []
        for i, var in enumerate(self.month_vars):
            if var.get():
                selected_months.append(i+1)
        self.config['month_config']['selected_months'] = selected_months
        
        # 验证配置
        if not self.config['source_file'] or not self.config['target_file']:
            messagebox.showerror("错误", "请选择参考文件和目标文件")
            return
        
        if not self.config['month_config']['selected_months']:
            messagebox.showerror("错误", "请至少选择一个月份")
            return
        
        # 准备配置数据
        config_data = {
            'source_file': self.config['source_file'],
            'target_file': self.config['target_file'],
            'source_config': {
                'indicator_position': 'column',
                'indicator_start_row': self.config['source_config']['indicator_start_row'],
                'category_row': self.config['source_config']['category_row'],
                'b_category_start_col': 1  # 从第2列开始，包含所有产品
            },
            'target_config': {
                'indicator_position': 'column',
                'indicator_start_row': self.config['target_config']['indicator_start_row'],
                'month_row': 2
            },
            'check_marks': {
                'source': [self.source_mark_var.get()],
                'target': self.target_mark_var.get()
            },
            'category_mapping': filtered_mapping,
            'month_config': self.config['month_config']
        }
        
        # 创建临时配置文件
        temp_config_path = self.config_manager.create_temp_config(config_data)
        
        # 执行映射
        try:
            selector = ExcelSelector(config_file=temp_config_path)
            selector.run()
            # 成功弹窗，包含打开结果的按钮
            if messagebox.askyesno("成功", "映射操作执行完成！\n是否打开结果文件？"):
                self.view_result()
        except Exception as e:
            messagebox.showerror("错误", f"执行失败: {str(e)}")
        finally:
            # 恢复原始映射，确保执行后不会清空
            self.config['category_mapping'] = original_mapping
            # 删除临时文件
            self.config_manager.delete_temp_config(temp_config_path)
    
    def refresh_config_list(self):
        """刷新配置列表"""
        # 获取配置文件列表
        config_files = self.config_manager.get_config_files()
        
        # 更新下拉框
        self.config_combobox['values'] = config_files
        if config_files:
            self.config_var.set(config_files[0])
    
    def save_config(self):
        """保存配置"""
        # 收集配置
        self.config['source_file'] = self.source_file_var.get()
        self.config['target_file'] = self.target_file_var.get()
        self.config['source_config']['category_row'] = int(self.category_row_var.get())
        self.config['source_config']['indicator_start_row'] = int(self.source_indicator_row_var.get())
        self.config['target_config']['indicator_start_row'] = int(self.target_indicator_row_var.get())
        
        # 收集标记配置
        self.config['check_marks'] = {
            'source': [self.source_mark_var.get()],
            'target': self.target_mark_var.get()
        }
        
        # 收集产品映射
        filtered_mapping = {}
        for product, var in self.config['category_mapping'].items():
            if hasattr(var, 'get'):
                target_product = var.get()
                if target_product:
                    filtered_mapping[product] = target_product
        self.config['category_mapping'] = filtered_mapping
        
        # 收集选中的月份
        selected_months = []
        for i, var in enumerate(self.month_vars):
            if var.get():
                selected_months.append(i+1)
        self.config['month_config']['selected_months'] = selected_months
        
        # 配置目录
        config_dir = os.path.join(os.path.dirname(__file__), "..", "configs")
        if not os.path.exists(config_dir):
            os.makedirs(config_dir)
        
        # 保存配置文件
        config_name = filedialog.asksaveasfilename(
            title="保存配置",
            initialdir=config_dir,
            defaultextension=".json",
            filetypes=[("JSON文件", "*.json")]
        )
        
        if config_name:
            try:
                with open(config_name, 'w', encoding='utf-8') as f:
                    json.dump(self.config, f, ensure_ascii=False, indent=2)
                messagebox.showinfo("成功", "配置保存成功！")
                # 刷新配置列表
                self.refresh_config_list()
            except Exception as e:
                messagebox.showerror("错误", f"保存配置失败: {str(e)}")
    
    def auto_load_last_config(self):
        """自动加载上次配置"""
        # 获取最后修改的配置文件
        last_config = self.config_manager.get_last_config()
        if last_config:
            self.config_var.set(last_config)
            # 加载最新配置
            self.load_config_from_combobox()
    
    def load_config_from_combobox(self):
        """从下拉框加载配置"""
        config_name = self.config_var.get()
        if not config_name:
            self.config_status_var.set("")
            return
        
        # 加载配置文件
        loaded_config = self.config_manager.load_config(config_name)
        if loaded_config:
            try:
                # 填充配置
                self.source_file_var.set(loaded_config.get('source_file', ''))
                self.target_file_var.set(loaded_config.get('target_file', ''))
                self.category_row_var.set(str(loaded_config.get('source_config', {}).get('category_row', 1)))
                self.source_indicator_row_var.set(str(loaded_config.get('source_config', {}).get('indicator_start_row', 3)))
                self.target_indicator_row_var.set(str(loaded_config.get('target_config', {}).get('indicator_start_row', 3)))
                
                # 加载标记配置
                check_marks = loaded_config.get('check_marks', {})
                self.source_mark_var.set(check_marks.get('source', ['✓'])[0])
                self.target_mark_var.set(check_marks.get('target', '✓'))
                
                # 加载月份选择
                selected_months = loaded_config.get('month_config', {}).get('selected_months', [])
                for i, var in enumerate(self.month_vars):
                    var.set(1 if (i+1) in selected_months else 0)
                
                # 加载产品映射
                self.config['category_mapping'] = loaded_config.get('category_mapping', {})
                
                # 自动加载产品
                self.load_products()
                
                # 显示加载成功标记
                self.config_status_var.set("✅")
            except Exception as e:
                print(f"加载配置失败: {str(e)}")
                self.config_status_var.set("")
        else:
            print(f"配置文件不存在: {config_name}")
            self.config_status_var.set("")

if __name__ == "__main__":
    import sys
    root = tk.Tk()
    app = ExcelSelectorGUI(root)
    root.mainloop()
