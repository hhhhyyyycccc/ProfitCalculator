import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import chardet
import numpy as np

class OrderAnalysisApp:
    def __init__(self, root):
        self.root = root
        self.root.title("订单盈亏分析工具 (支持CSV/Excel)")
        self.root.geometry("1100x750")
        
        # 存储数据
        self.order_df = None
        self.operator_df = None
        self.cost_df = None
        self.result_df = None
        
        # 创建界面
        self.create_widgets()
    
    def create_widgets(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重，使界面可调整大小
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # 标题
        title_label = ttk.Label(main_frame, text="订单盈亏分析工具 (支持CSV和Excel格式)", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="数据文件 (支持.csv, .xlsx, .xls格式)", padding="10")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        
        # 订单表
        ttk.Label(file_frame, text="订单表:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.order_file_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.order_file_var, state='readonly').grid(row=0, column=1, sticky=(tk.W, tk.E))
        ttk.Button(file_frame, text="选择文件", command=self.select_order_file).grid(row=0, column=2, padx=(10, 0))
        
        # 运营对照表
        ttk.Label(file_frame, text="运营对照表:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 0))
        self.operator_file_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.operator_file_var, state='readonly').grid(row=1, column=1, sticky=(tk.W, tk.E), pady=(10, 0))
        ttk.Button(file_frame, text="选择文件", command=self.select_operator_file).grid(row=1, column=2, padx=(10, 0), pady=(10, 0))
        
        # 成本对照表
        ttk.Label(file_frame, text="成本对照表:").grid(row=2, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 0))
        self.cost_file_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.cost_file_var, state='readonly').grid(row=2, column=1, sticky=(tk.W, tk.E), pady=(10, 0))
        ttk.Button(file_frame, text="选择文件", command=self.select_cost_file).grid(row=2, column=2, padx=(10, 0), pady=(10, 0))
        
        # 编码选择区域
        encoding_frame = ttk.LabelFrame(main_frame, text="CSV文件编码设置 (如遇乱码请尝试不同编码)", padding="10")
        encoding_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(encoding_frame, text="CSV文件编码:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.encoding_var = tk.StringVar(value="auto")
        encoding_combo = ttk.Combobox(encoding_frame, textvariable=self.encoding_var, 
                                     values=["auto", "utf-8", "gbk", "gb2312", "latin1", "iso-8859-1"],
                                     state="readonly", width=15)
        encoding_combo.grid(row=0, column=1, sticky=tk.W)
        ttk.Label(encoding_frame, text="(auto: 自动检测编码)").grid(row=0, column=2, sticky=tk.W, padx=(10, 0))
        
        # 数据预览区域
        preview_frame = ttk.LabelFrame(main_frame, text="数据预览", padding="10")
        preview_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        preview_frame.columnconfigure(0, weight=1)
        
        # 创建Notebook用于切换预览不同文件
        self.preview_notebook = ttk.Notebook(preview_frame)
        self.preview_notebook.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        # 为每个文件类型创建预览框架
        self.order_preview_frame = ttk.Frame(self.preview_notebook)
        self.operator_preview_frame = ttk.Frame(self.preview_notebook)
        self.cost_preview_frame = ttk.Frame(self.preview_notebook)
        
        self.preview_notebook.add(self.order_preview_frame, text="订单表预览")
        self.preview_notebook.add(self.operator_preview_frame, text="运营对照表预览")
        self.preview_notebook.add(self.cost_preview_frame, text="成本对照表预览")
        
        # 在每个预览框架中创建树形视图
        self.order_preview_tree = self.create_preview_tree(self.order_preview_frame)
        self.operator_preview_tree = self.create_preview_tree(self.operator_preview_frame)
        self.cost_preview_tree = self.create_preview_tree(self.cost_preview_frame)
        
        # 订单状态计算规则说明
        rule_frame = ttk.LabelFrame(main_frame, text="订单状态计算规则 (已支持商品数量)", padding="10")
        rule_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        rules_text = """
        • 已收货/已完成：盈亏 = 实收金额 - (商品成本 × 商品数量)
        • 退货/退款：盈亏 = -(商品成本 × 商品数量) - 运费(如有)
        • 已发货待收货：预计盈利 = 实收金额 - (商品成本 × 商品数量) (单独显示为待确认盈利)
        • 待发货/待处理：盈亏 = 0 (不计算盈亏)
        • 其他状态：盈亏 = 实收金额 - (商品成本 × 商品数量) (可根据需要调整)
        • 如果订单表没有商品数量列，默认数量为1
        • 如果成本表中未找到商品编码，该订单盈利记为0
        """
        rules_label = ttk.Label(rule_frame, text=rules_text, justify=tk.LEFT)
        rules_label.grid(row=0, column=0, sticky=tk.W)
        
        # 操作按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, columnspan=3, pady=(0, 10))
        
        ttk.Button(button_frame, text="分析数据", command=self.analyze_data).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="导出结果", command=self.export_results).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="清空数据", command=self.clear_data).pack(side=tk.LEFT)
        
        # 结果显示区域
        result_frame = ttk.LabelFrame(main_frame, text="分析结果", padding="10")
        result_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(6, weight=1)
        
        # 创建树形视图用于显示结果
        columns = ("运营", "订单总数", "已收货订单", "退货订单", "已发货待收货", "总盈亏", "待确认盈利", "平均每单盈亏")
        self.result_tree = ttk.Treeview(result_frame, columns=columns, show="headings", height=15)
        
        # 设置列标题
        for col in columns:
            self.result_tree.heading(col, text=col)
            # 调整列宽
            if col in ["运营"]:
                self.result_tree.column(col, width=120, anchor=tk.CENTER)
            elif col in ["订单总数", "已收货订单", "退货订单", "已发货待收货"]:
                self.result_tree.column(col, width=80, anchor=tk.CENTER)
            else:
                self.result_tree.column(col, width=100, anchor=tk.CENTER)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.result_tree.yview)
        self.result_tree.configure(yscrollcommand=scrollbar.set)
        
        self.result_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 状态栏
        self.status_var = tk.StringVar(value="准备就绪 - 请选择数据文件")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E))
    
    def create_preview_tree(self, parent):
        """创建用于预览数据的树形视图"""
        tree = ttk.Treeview(parent, height=6, show="headings")
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        return tree
    
    def detect_encoding(self, file_path):
        """检测文件编码"""
        try:
            with open(file_path, 'rb') as f:
                raw_data = f.read(10000)  # 读取前10000字节来检测编码
                result = chardet.detect(raw_data)
                encoding = result.get('encoding', 'utf-8')
                confidence = result.get('confidence', 0)
                
                # 如果置信度低于阈值，尝试常见的中文编码
                if confidence < 0.7:
                    # 尝试常见的中文编码
                    for enc in ['gbk', 'gb2312', 'utf-8']:
                        try:
                            with open(file_path, 'r', encoding=enc) as test_file:
                                test_file.read(1024)
                                return enc
                        except:
                            continue
                
                return encoding if encoding else 'utf-8'
        except Exception as e:
            # 如果检测失败，返回常见的中文编码
            return 'gbk'
    
    def read_csv_with_encoding(self, file_path, encoding_setting='auto'):
        """使用指定编码读取CSV文件"""
        if encoding_setting == 'auto':
            encoding = self.detect_encoding(file_path)
        else:
            encoding = encoding_setting
        
        # 尝试使用检测到的编码读取
        try:
            return pd.read_csv(file_path, encoding=encoding)
        except UnicodeDecodeError:
            # 如果失败，尝试其他常见编码
            for enc in ['gbk', 'gb2312', 'utf-8', 'latin1', 'iso-8859-1']:
                if enc != encoding:  # 跳过已经尝试过的编码
                    try:
                        return pd.read_csv(file_path, encoding=enc)
                    except:
                        continue
            # 如果所有编码都失败，抛出异常
            raise Exception(f"无法读取文件 {file_path}，尝试了多种编码均失败")
    
    def find_column(self, df, possible_names):
        """在数据框中查找可能的列名"""
        for name in possible_names:
            if name in df.columns:
                return name
        return None
    
    def load_data(self):
        """加载数据文件"""
        try:
            encoding_setting = self.encoding_var.get()
            
            # 加载订单表
            if self.order_file_var.get():
                if self.order_file_var.get().endswith('.csv'):
                    self.order_df = self.read_csv_with_encoding(self.order_file_var.get(), encoding_setting)
                    self.preview_data(self.order_df, self.order_preview_tree, "订单表")
                else:
                    self.order_df = pd.read_excel(self.order_file_var.get())
                    self.preview_data(self.order_df, self.order_preview_tree, "订单表")
            
            # 加载运营对照表
            if self.operator_file_var.get():
                if self.operator_file_var.get().endswith('.csv'):
                    self.operator_df = self.read_csv_with_encoding(self.operator_file_var.get(), encoding_setting)
                    self.preview_data(self.operator_df, self.operator_preview_tree, "运营对照表")
                else:
                    self.operator_df = pd.read_excel(self.operator_file_var.get())
                    self.preview_data(self.operator_df, self.operator_preview_tree, "运营对照表")
            
            # 加载成本对照表
            if self.cost_file_var.get():
                if self.cost_file_var.get().endswith('.csv'):
                    self.cost_df = self.read_csv_with_encoding(self.cost_file_var.get(), encoding_setting)
                    self.preview_data(self.cost_df, self.cost_preview_tree, "成本对照表")
                else:
                    self.cost_df = pd.read_excel(self.cost_file_var.get())
                    self.preview_data(self.cost_df, self.cost_preview_tree, "成本对照表")
                    
            return True
        except Exception as e:
            messagebox.showerror("错误", f"加载数据时出错: {str(e)}\n\n请尝试更改CSV文件编码设置。")
            return False
    
    def preview_data(self, df, tree, title):
        """在预览区域显示数据"""
        # 清空现有数据
        for item in tree.get_children():
            tree.delete(item)
        
        # 设置列
        if df is not None and not df.empty:
            tree["columns"] = list(df.columns)
            for col in df.columns:
                tree.heading(col, text=col)
                tree.column(col, width=100, anchor=tk.CENTER)
            
            # 添加数据（只显示前20行）
            for _, row in df.head(20).iterrows():
                tree.insert("", tk.END, values=list(row))
    
    def select_order_file(self):
        filename = filedialog.askopenfilename(
            title="选择订单表文件",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.order_file_var.set(filename)
            self.status_var.set(f"已选择订单表: {os.path.basename(filename)}")
            # 自动加载并预览数据
            self.load_and_preview_order_file(filename)
    
    def select_operator_file(self):
        filename = filedialog.askopenfilename(
            title="选择运营对照表文件",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.operator_file_var.set(filename)
            self.status_var.set(f"已选择运营对照表: {os.path.basename(filename)}")
            # 自动加载并预览数据
            self.load_and_preview_operator_file(filename)
    
    def select_cost_file(self):
        filename = filedialog.askopenfilename(
            title="选择成本对照表文件",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.cost_file_var.set(filename)
            self.status_var.set(f"已选择成本对照表: {os.path.basename(filename)}")
            # 自动加载并预览数据
            self.load_and_preview_cost_file(filename)
    
    def load_and_preview_order_file(self, filename):
        try:
            if filename.endswith('.csv'):
                encoding_setting = self.encoding_var.get()
                self.order_df = self.read_csv_with_encoding(filename, encoding_setting)
            else:
                self.order_df = pd.read_excel(filename)
            
            self.preview_data(self.order_df, self.order_preview_tree, "订单表")
            self.status_var.set(f"已加载订单表: {len(self.order_df)} 行数据")
        except Exception as e:
            messagebox.showerror("错误", f"加载订单表时出错: {str(e)}\n\n请尝试更改CSV文件编码设置。")
    
    def load_and_preview_operator_file(self, filename):
        try:
            if filename.endswith('.csv'):
                encoding_setting = self.encoding_var.get()
                self.operator_df = self.read_csv_with_encoding(filename, encoding_setting)
            else:
                self.operator_df = pd.read_excel(filename)
            
            self.preview_data(self.operator_df, self.operator_preview_tree, "运营对照表")
            self.status_var.set(f"已加载运营对照表: {len(self.operator_df)} 行数据")
        except Exception as e:
            messagebox.showerror("错误", f"加载运营对照表时出错: {str(e)}\n\n请尝试更改CSV文件编码设置。")
    
    def load_and_preview_cost_file(self, filename):
        try:
            if filename.endswith('.csv'):
                encoding_setting = self.encoding_var.get()
                self.cost_df = self.read_csv_with_encoding(filename, encoding_setting)
            else:
                self.cost_df = pd.read_excel(filename)
            
            self.preview_data(self.cost_df, self.cost_preview_tree, "成本对照表")
            self.status_var.set(f"已加载成本对照表: {len(self.cost_df)} 行数据")
        except Exception as e:
            messagebox.showerror("错误", f"加载成本对照表时出错: {str(e)}\n\n请尝试更改CSV文件编码设置。")
    
    def analyze_data(self):
        """分析数据并计算盈亏"""
        if not all([self.order_file_var.get(), self.operator_file_var.get(), self.cost_file_var.get()]):
            messagebox.showwarning("警告", "请先选择所有必需的数据文件")
            return
        
        self.status_var.set("正在分析数据...")
        self.root.update()
        
        # 确保数据已加载
        if self.order_df is None or self.operator_df is None or self.cost_df is None:
            if not self.load_data():
                return
        
        try:
            # 查找必要的列（支持多种可能的列名）
            product_id_col = self.find_column(self.order_df, ['商品ID', '商品id', '产品ID', '产品id', '商品编号'])
            product_code_col = self.find_column(self.order_df, ['商品编码', '商品编码', '产品编码', '产品编码', 'SKU'])
            status_col = self.find_column(self.order_df, ['订单状态', '状态', '订单状态', '状态'])
            amount_col = self.find_column(self.order_df, ['实收金额', '金额', '实收', '收入'])
            quantity_col = self.find_column(self.order_df, ['商品数量(件)', '商品数量', '数量', '数量(件)', '件数'])
            
            operator_product_id_col = self.find_column(self.operator_df, ['商品ID', '商品id', '产品ID', '产品id', '商品编号'])
            operator_product_code_col = self.find_column(self.operator_df, ['商品编码', '商品编码', '产品编码', '产品编码', 'SKU'])
            operator_col = self.find_column(self.operator_df, ['运营人员', '运营', '负责人', '运营人员', '负责人'])
            
            cost_product_code_col = self.find_column(self.cost_df, ['商品编码', '商品编码', '产品编码', '产品编码', 'SKU'])
            cost_col = self.find_column(self.cost_df, ['商品成本', '成本', '商品成本', '成本价'])
            
            # 检查必要的列是否存在
            missing_columns = []
            if not product_id_col:
                missing_columns.append("订单表中缺少商品ID列")
            if not product_code_col:
                missing_columns.append("订单表中缺少商品编码列")
            if not status_col:
                missing_columns.append("订单表中缺少订单状态列")
            if not amount_col:
                missing_columns.append("订单表中缺少实收金额列")
            if not operator_product_id_col:
                missing_columns.append("运营对照表中缺少商品ID列")
            if not operator_col:
                missing_columns.append("运营对照表中缺少运营人员列")
            if not cost_product_code_col:
                missing_columns.append("成本对照表中缺少商品编码列")
            if not cost_col:
                missing_columns.append("成本对照表中缺少商品成本列")
                
            if missing_columns:
                messagebox.showerror("错误", f"数据表中缺少必要的列:\n\n" + "\n".join(missing_columns) + 
                                    f"\n\n订单表列名: {', '.join(self.order_df.columns)}" +
                                    f"\n运营对照表列名: {', '.join(self.operator_df.columns)}" +
                                    f"\n成本对照表列名: {', '.join(self.cost_df.columns)}")
                return
            
            # 重命名列以统一处理
            self.order_df = self.order_df.rename(columns={
                product_id_col: '商品ID',
                product_code_col: '商品编码',
                status_col: '订单状态',
                amount_col: '实收金额'
            })
            
            # 如果订单表有商品数量列，则重命名；如果没有，则添加一列，值为1
            if quantity_col:
                self.order_df = self.order_df.rename(columns={quantity_col: '商品数量'})
                # 确保商品数量是数值类型
                self.order_df['商品数量'] = pd.to_numeric(self.order_df['商品数量'], errors='coerce').fillna(1)
            else:
                # 如果没有商品数量列，添加一列，默认值为1
                self.order_df['商品数量'] = 1
                self.status_var.set("订单表中未找到商品数量列，默认数量为1")
            
            # 检查运营对照表是否有商品编码列
            if operator_product_code_col:
                self.operator_df = self.operator_df.rename(columns={
                    operator_product_id_col: '商品ID',
                    operator_product_code_col: '商品编码',
                    operator_col: '运营人员'
                })
                # 如果运营对照表有商品编码，则使用商品ID+商品编码进行合并
                merge_on_operator = ['商品ID', '商品编码']
            else:
                self.operator_df = self.operator_df.rename(columns={
                    operator_product_id_col: '商品ID',
                    operator_col: '运营人员'
                })
                # 如果运营对照表没有商品编码，则只使用商品ID进行合并
                merge_on_operator = ['商品ID']
            
            self.cost_df = self.cost_df.rename(columns={
                cost_product_code_col: '商品编码',
                cost_col: '商品成本'
            })
            
            # 检查运营对照表中的重复记录
            if len(merge_on_operator) == 2:
                # 使用商品ID+商品编码合并
                duplicate_mask = self.operator_df.duplicated(subset=merge_on_operator, keep=False)
            else:
                # 仅使用商品ID合并
                duplicate_mask = self.operator_df.duplicated(subset=merge_on_operator, keep=False)
            
            if duplicate_mask.any():
                duplicate_count = duplicate_mask.sum()
                # 显示警告但不阻止分析
                messagebox.showwarning("数据警告", 
                    f"运营对照表中发现 {duplicate_count} 条重复记录\n\n"
                    f"重复记录可能会导致分析结果不准确。\n"
                    f"建议清理运营对照表中的重复数据。\n\n"
                    f"程序将继续分析，但会使用第一条匹配记录。")
                
                # 去除运营对照表中的重复记录，保留第一条
                self.operator_df = self.operator_df.drop_duplicates(subset=merge_on_operator, keep='first')
            
            # 合并数据
            # 首先将订单表与运营对照表合并（按照商品ID+商品编码）
            merged_df = pd.merge(
                self.order_df, 
                self.operator_df, 
                on=merge_on_operator, 
                how='left'
            )
            
            # 检查合并后的数据量
            original_order_count = len(self.order_df)
            merged_count = len(merged_df)
            
            if merged_count > original_order_count:
                # 合并后数据变多，说明存在一对多关系
                duplicate_info = f"警告: 合并后数据从 {original_order_count} 条增加到 {merged_count} 条\n"
                duplicate_info += "这可能是因为运营对照表中存在重复记录或一对多关系\n"
                duplicate_info += "程序已尝试处理重复，但建议检查数据质量"
                
                self.status_var.set(duplicate_info)
                # 不显示消息框，只在状态栏提示，避免打断分析流程
            
            # 将找不到运营人员的订单归类为"其他"
            merged_df['运营人员'] = merged_df['运营人员'].fillna('其他')
            
            # 然后将结果与成本对照表合并
            merged_df = pd.merge(
                merged_df, 
                self.cost_df, 
                on='商品编码', 
                how='left'
            )
            
            # 根据订单状态计算盈亏和待确认盈利
            def calculate_profit_loss(row):
                status = str(row['订单状态']).strip().lower()
                
                # 处理实收金额，确保是数字
                try:
                    received_amount = float(row['实收金额']) if pd.notna(row['实收金额']) else 0
                except (ValueError, TypeError):
                    received_amount = 0
                
                # 处理商品成本，确保是数字
                # 如果成本表中未找到商品编码，商品成本将为NaN，此时将盈利记为0
                if pd.isna(row['商品成本']):
                    # 成本表中未找到商品编码，盈亏记为0
                    return pd.Series([0, 0])
                
                try:
                    cost = float(row['商品成本']) if pd.notna(row['商品成本']) else 0
                except (ValueError, TypeError):
                    cost = 0
                
                # 处理商品数量，确保是数字
                try:
                    quantity = float(row['商品数量']) if pd.notna(row['商品数量']) else 1
                except (ValueError, TypeError):
                    quantity = 1
                
                # 计算总成本 = 商品成本 × 商品数量
                total_cost = cost * quantity
                
                # 初始化盈亏和待确认盈利
                profit_loss = 0
                pending_profit = 0
                
                # 根据状态计算盈亏
                if any(s in status for s in ['已收货', '已完成', '完成', '收货']):
                    profit_loss = received_amount - total_cost
                elif any(s in status for s in ['退货', '退款', '退回']):
                    # 假设退货订单有运费字段，如果没有，可以设置为0
                    try:
                        shipping_cost = float(row.get('运费', 0)) if pd.notna(row.get('运费')) else 0
                    except (ValueError, TypeError):
                        shipping_cost = 0
                    profit_loss = -total_cost - shipping_cost
                elif any(s in status for s in ['已发货待收货', '已发货', '待收货']):
                    # 已发货待收货状态，计算待确认盈利
                    pending_profit = received_amount - total_cost
                    # 这种状态的订单不计入已实现盈亏
                    profit_loss = 0
                elif any(s in status for s in ['待发货', '待处理', '待确认']):
                    profit_loss = 0  # 不计算盈亏
                else:
                    # 其他状态默认计算方式
                    profit_loss = received_amount - total_cost
                
                return pd.Series([profit_loss, pending_profit])
            
            # 应用计算函数
            merged_df[['盈亏', '待确认盈利']] = merged_df.apply(calculate_profit_loss, axis=1)
            
            # 按运营人员汇总
            if '运营人员' in merged_df.columns:
                # 分组计算各种指标
                result = merged_df.groupby('运营人员').agg(
                    订单总数=('商品ID', 'count'),
                    已收货订单=('订单状态', lambda x: x.str.contains('已收货|已完成|完成|收货', case=False, na=False).sum()),
                    退货订单=('订单状态', lambda x: x.str.contains('退货|退款|退回', case=False, na=False).sum()),
                    已发货待收货=('订单状态', lambda x: x.str.contains('已发货待收货|已发货|待收货', case=False, na=False).sum()),
                    总盈亏=('盈亏', 'sum'),
                    待确认盈利总额=('待确认盈利', 'sum'),
                    平均每单盈亏=('盈亏', 'mean')
                ).reset_index()
                
                # 重命名列以匹配显示
                result = result.rename(columns={
                    '运营人员': '运营',
                    '待确认盈利总额': '待确认盈利'
                })
                
                # 对数值列进行格式化
                result['总盈亏'] = result['总盈亏'].round(2)
                result['待确认盈利'] = result['待确认盈利'].round(2)
                result['平均每单盈亏'] = result['平均每单盈亏'].round(2)
                
                # 按总盈亏排序，但将"其他"放在最后
                result['排序权重'] = result['运营'].apply(lambda x: 0 if x == '其他' else 1)
                result = result.sort_values(['排序权重', '总盈亏'], ascending=[False, False]).drop('排序权重', axis=1)
                
                self.result_df = result
                
                # 清空现有结果
                for item in self.result_tree.get_children():
                    self.result_tree.delete(item)
                
                # 添加新结果
                for _, row in result.iterrows():
                    self.result_tree.insert("", tk.END, values=(
                        row['运营'],
                        int(row['订单总数']),
                        int(row['已收货订单']),
                        int(row['退货订单']),
                        int(row['已发货待收货']),
                        f"¥{row['总盈亏']:.2f}",
                        f"¥{row['待确认盈利']:.2f}",
                        f"¥{row['平均每单盈亏']:.2f}"
                    ))
                
                # 计算未匹配到运营人员的订单数量
                other_count = len(merged_df[merged_df['运营人员'] == '其他'])
                
                # 计算未匹配到成本的订单数量
                missing_cost_count = len(merged_df[pd.isna(merged_df['商品成本'])])
                
                # 显示合并方式信息
                merge_info = f"合并方式: 订单表与运营对照表按 {merge_on_operator} 合并"
                if len(merge_on_operator) == 2:
                    merge_info += " (商品ID+商品编码)"
                else:
                    merge_info += " (仅商品ID)"
                
                # 检查是否使用了商品数量
                quantity_info = ""
                if quantity_col:
                    quantity_info = f" - 已使用商品数量列: {quantity_col}"
                else:
                    quantity_info = " - 未找到商品数量列，默认数量为1"
                
                self.status_var.set(f"分析完成 - {merge_info} - 未匹配运营的订单: {other_count} 条 - 未匹配成本的订单: {missing_cost_count} 条{quantity_info}")
            else:
                messagebox.showerror("错误", "数据中未找到运营人员列，请检查数据格式")
                
        except Exception as e:
            messagebox.showerror("错误", f"分析数据时出错: {str(e)}")
            self.status_var.set("分析出错")
    
    def export_results(self):
        """导出结果到Excel文件"""
        if self.result_df is None:
            messagebox.showwarning("警告", "没有可导出的结果数据")
            return
        
        filename = filedialog.asksaveasfilename(
            title="保存分析结果",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
        )
        
        if filename:
            try:
                if filename.endswith('.csv'):
                    self.result_df.to_csv(filename, index=False, encoding='utf-8-sig')
                else:
                    self.result_df.to_excel(filename, index=False)
                
                messagebox.showinfo("成功", f"结果已导出到: {filename}")
                self.status_var.set(f"结果已导出: {os.path.basename(filename)}")
            except Exception as e:
                messagebox.showerror("错误", f"导出结果时出错: {str(e)}")
    
    def clear_data(self):
        """清空所有数据"""
        self.order_df = None
        self.operator_df = None
        self.cost_df = None
        self.result_df = None
        
        self.order_file_var.set("")
        self.operator_file_var.set("")
        self.cost_file_var.set("")
        
        # 清空预览数据
        for tree in [self.order_preview_tree, self.operator_preview_tree, self.cost_preview_tree]:
            for item in tree.get_children():
                tree.delete(item)
        
        # 清空结果数据
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        
        self.status_var.set("数据已清空")

def main():
    root = tk.Tk()
    app = OrderAnalysisApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()