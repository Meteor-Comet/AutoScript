import customtkinter as ctk
from tkinter import filedialog, messagebox, END
import pandas as pd
import os
from pathlib import Path

# 设置CustomTkinter主题
ctk.set_appearance_mode("System")  # "System" (default), "Dark", "Light"
ctk.set_default_color_theme("blue")  # "blue" (default), "green", "dark-blue"


class ModernExcelProcessorGUI:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("数据处理系统")
        self.root.geometry("900x700")
        
        # 文件路径变量
        self.disbursement_files = []
        self.shipment_file = None
        self.supplier_order_file = None
        self.sanze_file = None
        self.direct_mail_files = []
        
        self.setup_ui()
        
    def setup_ui(self):
        # 主框架
        self.main_frame = ctk.CTkFrame(self.root)
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # 标题
        title_label = ctk.CTkLabel(
            self.main_frame, 
            text="Excel文件处理系统", 
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title_label.pack(pady=(0, 20))
        
        # 创建选项卡视图
        self.tabview = ctk.CTkTabview(self.main_frame)
        self.tabview.pack(fill="both", expand=True, pady=10)
        
        # 添加选项卡
        self.tabview.add("文件选择")
        self.tabview.add("处理功能")
        self.tabview.add("处理结果")
        self.tabview.add("直邮明细导入")
        
        # 配置选项卡
        self.setup_file_selection_tab()
        self.setup_processing_tab()
        self.setup_results_tab()
        self.setup_direct_mail_tab()
        
        # 状态栏
        self.status_var = ctk.StringVar(value="就绪")
        status_bar = ctk.CTkLabel(
            self.main_frame, 
            textvariable=self.status_var,
            font=ctk.CTkFont(size=12),
            anchor="w"
        )
        status_bar.pack(fill="x", pady=(10, 0))
        
    def setup_file_selection_tab(self):
        tab = self.tabview.tab("文件选择")
        
        # 发放明细查询文件选择
        disbursement_frame = ctk.CTkFrame(tab)
        disbursement_frame.pack(fill="x", padx=20, pady=10)
        
        disbursement_label = ctk.CTkLabel(
            disbursement_frame, 
            text="发放明细查询文件:", 
            font=ctk.CTkFont(size=14, weight="bold")
        )
        disbursement_label.pack(anchor="w", padx=10, pady=(10, 5))
        
        # 文件列表框
        self.disbursement_listbox = ctk.CTkTextbox(disbursement_frame, height=100)
        self.disbursement_listbox.pack(fill="x", padx=10, pady=5)
        
        # 按钮框架
        disbursement_btn_frame = ctk.CTkFrame(disbursement_frame)
        disbursement_btn_frame.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkButton(
            disbursement_btn_frame, 
            text="添加文件", 
            command=self.add_disbursement_files
        ).pack(side="left", padx=5)
        
        ctk.CTkButton(
            disbursement_btn_frame, 
            text="清空列表", 
            command=self.clear_disbursement_files
        ).pack(side="left", padx=5)
        
        # 发货明细文件选择
        shipment_frame = ctk.CTkFrame(tab)
        shipment_frame.pack(fill="x", padx=20, pady=10)
        
        shipment_label = ctk.CTkLabel(
            shipment_frame, 
            text="发货明细文件:", 
            font=ctk.CTkFont(size=14, weight="bold")
        )
        shipment_label.pack(anchor="w", padx=10, pady=(10, 5))
        
        self.shipment_entry = ctk.CTkEntry(shipment_frame)
        self.shipment_entry.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkButton(
            shipment_frame, 
            text="浏览", 
            command=self.select_shipment_file
        ).pack(anchor="e", padx=10, pady=5)
        
        # 供应商订单查询文件选择
        supplier_frame = ctk.CTkFrame(tab)
        supplier_frame.pack(fill="x", padx=20, pady=10)
        
        supplier_label = ctk.CTkLabel(
            supplier_frame, 
            text="供应商订单查询文件:", 
            font=ctk.CTkFont(size=14, weight="bold")
        )
        supplier_label.pack(anchor="w", padx=10, pady=(10, 5))
        
        self.supplier_order_entry = ctk.CTkEntry(supplier_frame)
        self.supplier_order_entry.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkButton(
            supplier_frame, 
            text="浏览", 
            command=self.select_supplier_order_file
        ).pack(anchor="e", padx=10, pady=5)
        
    def setup_processing_tab(self):
        tab = self.tabview.tab("处理功能")
        
        # 功能按钮框架
        button_frame = ctk.CTkFrame(tab)
        button_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # 标题
        functions_label = ctk.CTkLabel(
            button_frame, 
            text="可用功能", 
            font=ctk.CTkFont(size=18, weight="bold")
        )
        functions_label.pack(pady=10)
        
        # 功能按钮
        ctk.CTkButton(
            button_frame,
            text="批量处理发放明细",
            font=ctk.CTkFont(size=16),
            height=40,
            command=self.process_disbursement_files
        ).pack(fill="x", padx=50, pady=10)
        
        ctk.CTkButton(
            button_frame,
            text="核对发货明细与供应商订单",
            font=ctk.CTkFont(size=16),
            height=40,
            command=self.verify_shipment_supplier
        ).pack(fill="x", padx=50, pady=10)
        
        ctk.CTkButton(
            button_frame,
            text="标记集采信息",
            font=ctk.CTkFont(size=16),
            height=40,
            command=self.mark_procurement_info
        ).pack(fill="x", padx=50, pady=10)
        
    def setup_results_tab(self):
        tab = self.tabview.tab("处理结果")
        
        # 结果文本框
        self.result_text = ctk.CTkTextbox(tab, wrap="word")
        self.result_text.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 清空按钮
        ctk.CTkButton(
            tab,
            text="清空结果",
            command=self.clear_results
        ).pack(pady=10)
        
    def setup_direct_mail_tab(self):
        tab = self.tabview.tab("直邮明细导入")
        
        # 三择导单文件选择
        sanze_frame = ctk.CTkFrame(tab)
        sanze_frame.pack(fill="x", padx=20, pady=10)
        
        sanze_label = ctk.CTkLabel(
            sanze_frame, 
            text="三择导单文件:", 
            font=ctk.CTkFont(size=14, weight="bold")
        )
        sanze_label.pack(anchor="w", padx=10, pady=(10, 5))
        
        self.sanze_entry = ctk.CTkEntry(sanze_frame)
        self.sanze_entry.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkButton(
            sanze_frame, 
            text="浏览", 
            command=self.select_sanze_file
        ).pack(anchor="e", padx=10, pady=5)
        
        # 直邮明细文件选择
        direct_mail_frame = ctk.CTkFrame(tab)
        direct_mail_frame.pack(fill="x", padx=20, pady=10)
        
        direct_mail_label = ctk.CTkLabel(
            direct_mail_frame, 
            text="直邮明细文件:", 
            font=ctk.CTkFont(size=14, weight="bold")
        )
        direct_mail_label.pack(anchor="w", padx=10, pady=(10, 5))
        
        # 文件列表框
        self.direct_mail_listbox = ctk.CTkTextbox(direct_mail_frame, height=100)
        self.direct_mail_listbox.pack(fill="x", padx=10, pady=5)
        
        # 按钮框架
        direct_mail_btn_frame = ctk.CTkFrame(direct_mail_frame)
        direct_mail_btn_frame.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkButton(
            direct_mail_btn_frame, 
            text="添加文件", 
            command=self.add_direct_mail_files
        ).pack(side="left", padx=5)
        
        ctk.CTkButton(
            direct_mail_btn_frame, 
            text="清空列表", 
            command=self.clear_direct_mail_files
        ).pack(side="left", padx=5)
        
        # 导入按钮
        ctk.CTkButton(
            tab,
            text="开始导入",
            font=ctk.CTkFont(size=16),
            height=40,
            command=self.import_direct_mail_to_sanze
        ).pack(fill="x", padx=50, pady=20)
        
    def add_disbursement_files(self):
        files = filedialog.askopenfilenames(
            title="选择发放明细查询文件",
            filetypes=[("Excel文件", "*.xlsx *.xls")]
        )
        for file in files:
            if file not in self.disbursement_files:
                self.disbursement_files.append(file)
                filename = Path(file).name
                self.disbursement_listbox.insert("end", f"{filename}\n")
        self.status_var.set(f"已添加 {len(self.disbursement_files)} 个发放明细查询文件")
        
    def clear_disbursement_files(self):
        self.disbursement_files.clear()
        self.disbursement_listbox.delete("0.0", "end")
        self.status_var.set("已清空发放明细查询文件列表")
        
    def select_shipment_file(self):
        file = filedialog.askopenfilename(
            title="选择发货明细文件",
            filetypes=[("Excel文件", "*.xlsx *.xls")]
        )
        if file:
            self.shipment_file = file
            self.shipment_entry.delete(0, END)
            self.shipment_entry.insert(0, file)
            self.status_var.set("已选择发货明细文件")
            
    def select_supplier_order_file(self):
        file = filedialog.askopenfilename(
            title="选择供应商订单查询文件",
            filetypes=[("Excel文件", "*.xlsx *.xls")]
        )
        if file:
            self.supplier_order_file = file
            self.supplier_order_entry.delete(0, END)
            self.supplier_order_entry.insert(0, file)
            self.status_var.set("已选择供应商订单查询文件")
            
    def select_sanze_file(self):
        file = filedialog.askopenfilename(
            title="选择三择导单文件",
            filetypes=[("Excel文件", "*.xlsx *.xls")]
        )
        if file:
            self.sanze_file = file
            self.sanze_entry.delete(0, END)
            self.sanze_entry.insert(0, file)
            self.status_var.set("已选择三择导单文件")
            
    def add_direct_mail_files(self):
        files = filedialog.askopenfilenames(
            title="选择直邮明细文件",
            filetypes=[("Excel文件", "*.xlsx *.xls")]
        )
        for file in files:
            if file not in self.direct_mail_files:
                self.direct_mail_files.append(file)
                filename = Path(file).name
                self.direct_mail_listbox.insert("end", f"{filename}\n")
        self.status_var.set(f"已添加 {len(self.direct_mail_files)} 个直邮明细文件")
        
    def clear_direct_mail_files(self):
        self.direct_mail_files.clear()
        self.direct_mail_listbox.delete("0.0", "end")
        self.status_var.set("已清空直邮明细文件列表")
            
    def clear_results(self):
        self.result_text.delete("0.0", "end")
        self.status_var.set("已清空结果")
        
    def process_disbursement_files(self):
        if not self.disbursement_files:
            messagebox.showwarning("警告", "请先选择发放明细查询文件")
            return
            
        try:
            self.result_text.delete("0.0", "end")
            self.result_text.insert("end", "开始处理发放明细查询文件...\n\n")
            self.tabview.set("处理结果")  # 切换到结果选项卡
            
            all_results = []
            for file_path in self.disbursement_files:
                df = self.process_disbursement_file(file_path)
                if not df.empty:
                    all_results.append(df)
                    self.result_text.insert("end", f"处理完成: {Path(file_path).name}\n")
                else:
                    self.result_text.insert("end", f"处理完成，但无数据: {Path(file_path).name}\n")
                    
            if all_results:
                combined_df = pd.concat(all_results, ignore_index=True)
                
                # 保存结果
                output_file = "发货明细.xlsx"
                combined_df.to_excel(output_file, index=False)
                
                self.result_text.insert("end", f"\n合并完成，总行数: {len(combined_df)}\n")
                self.result_text.insert("end", f"结果已保存为: {output_file}\n")
                self.status_var.set(f"处理完成，生成文件: {output_file}")
            else:
                self.result_text.insert("end", "\n未生成任何数据\n")
                self.status_var.set("处理完成，但未生成数据")
                
        except Exception as e:
            error_msg = f"处理过程中出错: {str(e)}"
            self.result_text.insert("end", f"\n{error_msg}\n")
            self.status_var.set("处理失败")
            messagebox.showerror("错误", error_msg)
            
    def process_disbursement_file(self, file_path):
        """
        处理单个发放明细查询文件，提取数据
        """
        try:
            # 读取Excel文件，不指定header以便手动处理
            df = pd.read_excel(file_path, header=None)
            
            if df.empty:
                return pd.DataFrame()
            
            # 获取标题行（第一行是文件名说明，第二行开始是实际标题）
            header_row2 = df.iloc[1]  # 第二行（产品名称）
            header_row3 = df.iloc[2]  # 第三行（公司名称）
            
            # 找到包含特定公司的列
            target_cols = []
            for i in range(len(header_row3)):
                if pd.notna(header_row3[i]) and '怡亚通' in str(header_row3[i]):
                    product_name = header_row2[i] if pd.notna(header_row2[i]) else ""
                    target_cols.append((i, product_name))
            
            if not target_cols:
                return pd.DataFrame()
            
            # 提取前14列的通用信息
            info_columns = list(range(14))
            
            # 构建结果数据
            result_data = []
            
            # 从第4行开始是数据行（索引3）
            for row_idx in range(3, len(df)):
                row = df.iloc[row_idx]
                
                # 跳过合计行
                if pd.notna(row.iloc[0]) and '合计' in str(row.iloc[0]):
                    continue
                
                # 为每个目标产品创建一行
                for col_idx, product_name in target_cols:
                    # 检查该列的数量是否不为0
                    if col_idx < len(row) and pd.notna(row.iloc[col_idx]) and row.iloc[col_idx] != 0:
                        quantity = row.iloc[col_idx]
                        # 只有数量不为0的数据才添加
                        if quantity != 0:
                            # 获取通用信息
                            base_info = []
                            for col_idx_info in info_columns:
                                if col_idx_info < len(row):
                                    base_info.append(row.iloc[col_idx_info])
                                else:
                                    base_info.append(None)
                            
                            # 创建一行数据：通用信息(14列) + 产品名称 + 数量
                            new_row = base_info + [product_name, quantity]
                            result_data.append(new_row)
            
            # 更新列定义
            columns = ['制单日期', '打印时间', '领用单号', '领用类型', '方案类型', '领用部门', 
                      '区域', '方案提报人', '方案提报人电话', '发货仓库', '收货人', 
                      '收货人电话', '收货地址', '领用说明', '产品名称', '数量']
            
            result_df = pd.DataFrame(result_data, columns=columns)
            return result_df
            
        except Exception as e:
            self.result_text.insert("end", f"处理文件 {Path(file_path).name} 时出错: {e}\n")
            return pd.DataFrame()
            
    def verify_shipment_supplier(self):
        if not self.shipment_file or not self.supplier_order_file:
            messagebox.showwarning("警告", "请先选择发货明细文件和供应商订单查询文件")
            return
            
        try:
            self.result_text.delete("0.0", "end")
            self.result_text.insert("end", "开始核对发货明细与供应商订单...\n\n")
            self.tabview.set("处理结果")  # 切换到结果选项卡
            
            # 读取文件
            shipment_df = pd.read_excel(self.shipment_file)
            supplier_df = pd.read_excel(self.supplier_order_file)
            
            # 筛选供应商为怡亚通的订单
            yiyatong_orders = supplier_df[supplier_df['供应商'].str.contains('怡亚通', na=False)].copy()
            
            # 重命名供应商订单列以便后续处理
            order_summary = yiyatong_orders[['方案编号', '收货人', '商品名称', '数量']].copy()
            order_summary.rename(columns={'数量': '订单数量'}, inplace=True)
            
            # 重命名发货明细列以便后续处理
            delivery_summary = shipment_df[['领用说明', '收货人', '产品名称', '数量']].copy()
            delivery_summary.rename(columns={
                '领用说明': '方案编号',
                '产品名称': '商品名称',
                '数量': '发货数量'
            }, inplace=True)
            
            # 按方案编号、收货人、商品名称分组汇总数量
            order_grouped = order_summary.groupby(['方案编号', '收货人', '商品名称'])['订单数量'].sum().reset_index()
            delivery_grouped = delivery_summary.groupby(['方案编号', '收货人', '商品名称'])['发货数量'].sum().reset_index()
            
            # 以供应商订单为准进行核对
            # 左连接确保所有供应商订单记录都被包含
            merged_data = pd.merge(order_grouped, delivery_grouped, on=['方案编号', '收货人', '商品名称'], how='left')
            
            # 填充NaN值（发货明细中没有对应记录的情况）
            merged_data['发货数量'] = merged_data['发货数量'].fillna(0)
            
            # 计算差异
            merged_data['数量差异'] = merged_data['订单数量'] - merged_data['发货数量']
            
            # 判断是否一致
            merged_data['是否一致'] = merged_data['数量差异'] == 0
            
            # 分类结果
            consistent_records = merged_data[merged_data['是否一致']]
            inconsistent_records = merged_data[(~merged_data['是否一致']) & (merged_data['发货数量'] > 0)]  # 数量不一致
            not_found_records = merged_data[merged_data['发货数量'] == 0]  # 发货明细中没找到
            
            # 显示核对结果
            self.result_text.insert("end", "核对结果:\n")
            self.result_text.insert("end", "========================\n")
            
            if not consistent_records.empty:
                self.result_text.insert("end", f"\n数量一致的记录 ({len(consistent_records)} 条):\n")
                for _, row in consistent_records.iterrows():
                    self.result_text.insert("end", 
                        f"方案编号: {row['方案编号']}, "
                        f"收货人: {row['收货人']}, "
                        f"商品名称: {row['商品名称']}, "
                        f"订单数量: {row['订单数量']}, "
                        f"发货数量: {row['发货数量']}\n")
            
            if not inconsistent_records.empty:
                self.result_text.insert("end", f"\n数量不一致的记录 ({len(inconsistent_records)} 条):\n")
                for _, row in inconsistent_records.iterrows():
                    self.result_text.insert("end", 
                        f"方案编号: {row['方案编号']}, "
                        f"收货人: {row['收货人']}, "
                        f"商品名称: {row['商品名称']}, "
                        f"订单数量: {row['订单数量']}, "
                        f"发货数量: {row['发货数量']}, "
                        f"数量差异: {row['数量差异']}\n")
            
            if not not_found_records.empty:
                self.result_text.insert("end", f"\n发货明细中未找到的记录 ({len(not_found_records)} 条):\n")
                for _, row in not_found_records.iterrows():
                    self.result_text.insert("end", 
                        f"方案编号: {row['方案编号']}, "
                        f"收货人: {row['收货人']}, "
                        f"商品名称: {row['商品名称']}, "
                        f"订单数量: {row['订单数量']}\n")
            
            # 统计信息
            total_records = len(merged_data)
            consistent_count = len(consistent_records)
            inconsistent_count = len(inconsistent_records)
            not_found_count = len(not_found_records)
            self.result_text.insert("end", f"\n核对完成，共核对 {total_records} 条供应商订单记录\n")
            self.result_text.insert("end", f"  - 数量一致: {consistent_count} 条\n")
            self.result_text.insert("end", f"  - 数量不一致: {inconsistent_count} 条\n")
            self.result_text.insert("end", f"  - 发货明细中未找到: {not_found_count} 条\n")
            self.status_var.set("核对完成")
            
        except Exception as e:
            error_msg = f"核对过程中出错: {str(e)}"
            self.result_text.insert("end", f"\n{error_msg}\n")
            self.status_var.set("核对失败")
            messagebox.showerror("错误", error_msg)
            
    def mark_procurement_info(self):
        if not self.shipment_file or not self.supplier_order_file:
            messagebox.showwarning("警告", "请先选择发货明细文件和供应商订单查询文件")
            return
            
        try:
            self.result_text.delete("0.0", "end")
            self.result_text.insert("end", "开始标记集采信息...\n\n")
            self.tabview.set("处理结果")  # 切换到结果选项卡
            
            # 读取文件
            shipment_df = pd.read_excel(self.shipment_file)
            supplier_df = pd.read_excel(self.supplier_order_file)
            
            # 筛选供应商为怡亚通的订单
            yiyatong_orders = supplier_df[supplier_df['供应商'].str.contains('怡亚通', na=False)].copy()
            
            # 根据'一件代发'列判断是否集采（一件代发为'是'表示非集采）
            yiyatong_orders['是否集采'] = yiyatong_orders['一件代发'].apply(
                lambda x: '非集采' if str(x).strip() == '是' else '集采'
            )
            
            # 按方案编号、收货人、联系方式和商品名称分组，确定每个组合的集采状态
            procurement_status = yiyatong_orders.groupby(['方案编号', '收货人', '联系方式', '商品名称'])['是否集采'].apply(
                lambda x: '集采' if '集采' in x.values else '非集采'
            ).reset_index()
            
            # 创建一个用于标记的发货明细副本
            shipment_marked = shipment_df.copy()
            
            # 初始化'是否集采'列为空字符串
            shipment_marked['是否集采'] = ''
            
            # 根据匹配条件更新'是否集采'列
            marked_count = 0
            for _, row in procurement_status.iterrows():
                mask = (
                    (shipment_marked['领用说明'] == row['方案编号']) &
                    (shipment_marked['收货人'] == row['收货人']) &
                    (shipment_marked['收货人电话'] == row['联系方式']) &
                    (shipment_marked['产品名称'] == row['商品名称'])
                )
                matched_rows = shipment_marked[mask]
                if not matched_rows.empty:
                    shipment_marked.loc[mask, '是否集采'] = row['是否集采']
                    marked_count += len(matched_rows)
            
            # 保存结果
            output_file = "发货明细_集采标记.xlsx"
            shipment_marked.to_excel(output_file, index=False)
            
            # 统计信息
            total_rows = len(shipment_marked)
            empty_rows = len(shipment_marked[shipment_marked['是否集采'] == ''])
            
            self.result_text.insert("end", f"标记完成:\n")
            self.result_text.insert("end", f"- 总记录数: {total_rows}\n")
            self.result_text.insert("end", f"- 已标记集采信息的记录数: {marked_count}\n")
            self.result_text.insert("end", f"- 未匹配到供应商订单的记录数: {empty_rows} (保持空白)\n")
            self.result_text.insert("end", f"- 结果已保存为: {output_file}\n")
            
            # 显示前几行标记结果作为示例
            self.result_text.insert("end", f"\n标记结果示例:\n")
            self.result_text.insert("end", "========================\n")
            self.result_text.insert("end", shipment_marked[['领用说明', '收货人', '产品名称', '是否集采']].head(10).to_string(index=False))
            self.result_text.insert("end", "\n")
            
            self.status_var.set(f"标记完成，生成文件: {output_file}")
            
        except Exception as e:
            error_msg = f"标记过程中出错: {str(e)}"
            self.result_text.insert("end", f"\n{error_msg}\n")
            self.status_var.set("标记失败")
            messagebox.showerror("错误", error_msg)
            
    def import_direct_mail_to_sanze(self):
        if not self.sanze_file or not self.direct_mail_files:
            messagebox.showwarning("警告", "请先选择三择导单文件和直邮明细文件")
            return
            
        try:
            self.result_text.delete("0.0", "end")
            self.result_text.insert("end", "开始导入直邮明细到三择导单...\n\n")
            self.tabview.set("处理结果")  # 切换到结果选项卡
            
            # 读取三择导单文件
            sanze_df = pd.read_excel(self.sanze_file)
            self.result_text.insert("end", f"读取三择导单文件，共 {len(sanze_df)} 行数据\n")
            
            # 处理所有直邮明细文件
            all_direct_mail_data = []
            for file_path in self.direct_mail_files:
                try:
                    # 读取直邮明细文件
                    direct_mail_df = pd.read_excel(file_path)
                    filename = Path(file_path).name
                    self.result_text.insert("end", f"读取直邮明细文件 {filename}，共 {len(direct_mail_df)} 行数据\n")
                    
                    # 处理直邮明细数据，提取需要的列
                    processed_data = self.extract_direct_mail_info(direct_mail_df, filename)
                    if processed_data:
                        all_direct_mail_data.extend(processed_data)
                        self.result_text.insert("end", f"处理完成 {filename}，提取到 {len(processed_data)} 行数据\n")
                except Exception as e:
                    error_msg = f"处理文件 {filename} 时出错: {str(e)}"
                    self.result_text.insert("end", f"{error_msg}\n")
                    
            # 将处理后的数据添加到三择导单
            if all_direct_mail_data:
                self.result_text.insert("end", f"\n总共提取了 {len(all_direct_mail_data)} 行直邮数据\n")
                
                # 创建新的数据行并添加到三择导单
                new_rows = []
                for item in all_direct_mail_data:
                    # 创建一个新行，保持三择导单的列结构
                    new_row = {}
                    for col in sanze_df.columns:
                        new_row[col] = ''  # 默认为空字符串
                    
                    # 填充相关信息
                    new_row['收货人'] = item.get('收货人', '')
                    new_row['手机'] = item.get('手机', '')
                    new_row['收货地址'] = item.get('收货地址', '')
                    new_row['客户名称'] = item.get('客户名称', '')
                    new_row['客户备注'] = f"导入自 {item.get('源文件', '')}"
                    
                    new_rows.append(new_row)
                
                # 将新行转换为DataFrame并添加到三择导单
                if new_rows:
                    new_rows_df = pd.DataFrame(new_rows)
                    # 确保新行的列与三择导单一致
                    new_rows_df = new_rows_df.reindex(columns=sanze_df.columns, fill_value='')
                    # 合并到三择导单（添加到末尾）
                    result_df = pd.concat([sanze_df, new_rows_df], ignore_index=True)
                    
                    # 保存结果
                    output_file = f"三择导单_导入直邮明细_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    result_df.to_excel(output_file, index=False)
                    
                    self.result_text.insert("end", f"导入完成，最终数据共 {len(result_df)} 行\n")
                    self.result_text.insert("end", f"结果已保存为: {output_file}\n")
                    self.status_var.set(f"导入完成，生成文件: {output_file}")
                    
                    # 显示前几行结果作为示例
                    self.result_text.insert("end", f"\n导入结果示例:\n")
                    self.result_text.insert("end", "========================\n")
                    preview_data = result_df[['客户名称', '收货人', '手机', '收货地址', '客户备注']].tail(10)
                    self.result_text.insert("end", preview_data.to_string(index=False))
                    self.result_text.insert("end", "\n")
                else:
                    self.result_text.insert("end", "合并数据时出错\n")
                    self.status_var.set("导入失败")
            else:
                self.result_text.insert("end", "没有成功提取任何直邮数据\n")
                self.status_var.set("导入完成，但未生成数据")
                
        except Exception as e:
            error_msg = f"导入过程中出错: {str(e)}"
            self.result_text.insert("end", f"\n{error_msg}\n")
            self.status_var.set("导入失败")
            messagebox.showerror("错误", error_msg)
            
    def extract_direct_mail_info(self, df, filename):
        """
        从直邮明细文件中提取姓名、联系方式、收货地址等信息
        """
        try:
            extracted_data = []
            
            # 根据文件名确定市场类型
            market = ""
            if "兰州" in filename:
                market = "兰州市场"
            elif "酒泉" in filename:
                market = "酒泉市场"
            
            # 遍历数据行并提取信息
            for index, row in df.iterrows():
                try:
                    # 根据不同市场的文件结构处理数据
                    if market == "兰州市场":
                        # 兰州市场文件结构: 序号, 客户名称, 地址, 联系人姓名, 电话
                        item = {
                            '客户名称': row.get('Unnamed: 1', ''),  # 客户名称在第二列
                            '收货地址': row.get('Unnamed: 2', ''),  # 地址在第三列
                            '收货人': row.get('Unnamed: 3', ''),    # 联系人姓名在第四列
                            '手机': row.get('Unnamed: 4', ''),      # 电话在第五列
                            '源文件': filename
                        }
                    elif market == "酒泉市场":
                        # 酒泉市场文件结构: 序号, 区域, NaN, NaN, 电话, 详细邮寄地址, 访销周期
                        item = {
                            '客户名称': '',  # 酒泉市场文件中没有明确的客户名称列
                            '收货地址': row.get('Unnamed: 6', ''),  # 详细邮寄地址在第七列
                            '收货人': '',   # 酒泉市场文件中没有明确的收货人列
                            '手机': row.get('Unnamed: 5', ''),      # 电话在第六列
                            '源文件': filename
                        }
                    else:
                        # 默认处理方式
                        item = {
                            '客户名称': row.get('客户名称', row.get('Unnamed: 1', '')),
                            '收货地址': row.get('地址', row.get('Unnamed: 2', '')),
                            '收货人': row.get('联系人姓名', row.get('Unnamed: 3', '')),
                            '手机': row.get('电话', row.get('Unnamed: 4', '')),
                            '源文件': filename
                        }
                    
                    extracted_data.append(item)
                except Exception as e:
                    self.result_text.insert("end", f"处理第 {index+1} 行时出错: {e}\n")
            
            return extracted_data
            
        except Exception as e:
            self.result_text.insert("end", f"处理直邮明细文件 {filename} 时出错: {e}\n")
            return []
            
    def run(self):
        self.root.mainloop()


def main():
    app = ModernExcelProcessorGUI()
    app.run()


if __name__ == "__main__":
    main()