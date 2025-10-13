# 主应用入口文件
import streamlit as st
import pandas as pd
import os
import glob
from datetime import datetime
import tempfile
import shutil

# 导入模块
from modules.column_detection import smart_column_detection, find_best_header_row
from modules.data_processing import (
    batch_process_files, process_supplier_order_file, 
    compare_data, mark_procurement_info, extract_direct_mail_info
)
from modules.ui_components import setup_page, setup_sidebar, show_footer, download_button
from modules.product_functions import standardize_product_names, convert_product_quantities, convert_product_quantities_manual
from modules.enhanced_vlookup import enhanced_vlookup
from modules.logistics_matching import match_logistics_info

# 设置页面
setup_page()

# 设置侧边栏
app_mode = setup_sidebar()

# 商品列表
product_list = [
    "奥克斯（AUX） 除螨仪 90W （计价单位：台）",
    "国产定制 黄金葉  盒装抽纸  （计价单位：盒）",
    "国产定制 黄金葉 环保塑料袋  50个/捆 300个/箱 （计价单位：个）",
    "国产定制 黄金葉 两盒装翻盖式礼盒 30个/箱 （计价单位：个）",
    "国产定制 黄金葉 湿纸巾 10片/包 （计价单位：包）",
    "国产定制 黄金葉 四盒装翻盖式礼盒 30个/箱 （计价单位：个）",
    "国产定制 黄金葉 四盒装简易封套（天叶品系） 50个/箱 （计价单位：个）",
    "国产定制 黄金葉 天叶叁毫克   两条装纸袋 （计价单位：个）",
    "国产定制 黄金葉 五盒装简易封套（常规款） 50个/箱 （计价单位：个）",
    "国产定制 黄金葉 五盒装简易封套（细支款） 50个/箱 （计价单位：个）",
    "国产定制 黄金葉 五盒装简易封套（中支款） 50个/箱 （计价单位：个）",
    "品胜（PISEN） 数据线三合一充电线100W  一拖三 （计价单位：条）",
    "有色 剃须刀便携合金电动刮胡刀男士  MINI 2.0 （计价单位：个） 颜色随机"
]

if app_mode == "批量处理发放明细":
    st.header("批量处理发放明细")

    # 文件上传
    uploaded_files = st.file_uploader(
        "上传发放明细查询文件（支持多个文件）",
        type=["xlsx"],
        accept_multiple_files=True
    )

    if uploaded_files:
        st.info(f"已选择 {len(uploaded_files)} 个文件")
        
        # 显示上传的文件名
        file_names = [f.name for f in uploaded_files]
        st.write("上传的文件:")
        st.write(file_names)
        
        # 处理按钮
        if st.button("开始处理"):
            with st.spinner("正在处理文件..."):
                result_df = batch_process_files(uploaded_files)
                
                if result_df is not None and not result_df.empty:
                    st.success(f"处理完成！共汇总 {len(result_df)} 行数据")
                    
                    # 显示结果预览
                    st.subheader("处理结果预览")
                    st.dataframe(result_df.head(20))
                    
                    # 提供下载
                    download_button(result_df, "发货明细")

elif app_mode == "核对发放明细与供应商订单":
    st.header("核对发放明细与供应商订单")
    st.info("此功能用于核对供应商订单中怡亚通的数据与发放明细是否一致，并标记集采信息")
    
    col1, col2 = st.columns(2)
    
    with col1:
        delivery_file = st.file_uploader(
            "上传发放明细文件",
            type=["xlsx"],
            key="delivery_file"
        )
    
    with col2:
        order_files = st.file_uploader(
            "上传供应商订单文件（支持多个）",
            type=["xlsx"],
            accept_multiple_files=True,
            key="order_files"
        )
    
    if delivery_file and order_files:
        # 添加选项让用户选择要执行的操作
        st.subheader("选择操作")
        perform_verification = st.checkbox("核对发放明细与供应商订单", value=True)
        perform_procurement_marking = st.checkbox("标记集采信息", value=True)
        
        if st.button("开始处理"):
            with st.spinner("正在处理数据..."):
                # 保存上传的文件到临时目录
                with tempfile.TemporaryDirectory() as temp_dir:
                    # 保存发货明细文件
                    delivery_path = os.path.join(temp_dir, delivery_file.name)
                    with open(delivery_path, "wb") as f:
                        f.write(delivery_file.getbuffer())
                    
                    # 读取发货明细
                    try:
                        发货明细_df = pd.read_excel(delivery_path)
                        st.info(f"读取发货明细文件，共 {len(发货明细_df)} 行数据")
                    except Exception as e:
                        st.error(f"读取发货明细文件时出错: {e}")
                        st.stop()
                    
                    # 保存供应商订单文件
                    order_file_paths = []
                    for uploaded_file in order_files:
                        file_path = os.path.join(temp_dir, uploaded_file.name)
                        with open(file_path, "wb") as f:
                            f.write(uploaded_file.getbuffer())
                        order_file_paths.append(file_path)
                    
                    # 处理所有供应商订单文件
                    all_orders = []
                    for file_path in order_file_paths:
                        with st.spinner(f"正在处理 {os.path.basename(file_path)}..."):
                            df = process_supplier_order_file(file_path)
                            if df is not None:
                                all_orders.append(df)
                                st.success(f"成功处理 {os.path.basename(file_path)}")
                    
                    if all_orders:
                        # 合并所有订单数据
                        供应商订单_df = pd.concat(all_orders, ignore_index=True)
                        st.info(f"合并订单数据，共 {len(供应商订单_df)} 行")
                        
                        # 初始化结果DataFrame
                        final_result = 发货明细_df.copy()
                        
                        # 执行核对操作
                        if perform_verification:
                            # 进行数据核对
                            consistent_records, inconsistent_records, not_found_records, _ = compare_data(发货明细_df, 供应商订单_df)
                            
                            # 显示核对结果
                            st.success("数据核对完成！")
                            
                            # 显示一致的记录
                            if not consistent_records.empty:
                                st.subheader("数量一致的记录")
                                st.dataframe(consistent_records)
                                st.success(f"发现 {len(consistent_records)} 条数量一致的记录")
                            
                            # 显示数量不一致的记录
                            if not inconsistent_records.empty:
                                st.subheader("数量不一致的记录")
                                st.dataframe(inconsistent_records)
                                st.warning(f"发现 {len(inconsistent_records)} 条数量不一致的记录")
                            
                            # 显示发放明细中未找到的记录
                            if not not_found_records.empty:
                                st.subheader("发放明细中未找到的记录")
                                st.dataframe(not_found_records)
                                st.error(f"发现 {len(not_found_records)} 条发放明细中未找到的记录")
                            
                            # 统计信息
                            total_records = len(consistent_records) + len(inconsistent_records) + len(not_found_records)
                            consistent_count = len(consistent_records)
                            st.info(f"总共核对 {total_records} 条供应商订单记录")
                            st.info(f"  - 数量一致: {consistent_count} 条")
                            st.info(f"  - 数量不一致: {len(inconsistent_records)} 条")
                            st.info(f"  - 发放明细中未找到: {len(not_found_records)} 条")
                        
                        # 执行集采标记操作
                        if perform_procurement_marking:
                            # 标记集采信息
                            marked_df = mark_procurement_info(发货明细_df, 供应商订单_df)
                            
                            if marked_df is not None:
                                final_result = marked_df
                                st.success("集采信息标记完成！")
                                
                                # 显示标记结果统计
                                procurement_stats = marked_df['是否集采'].value_counts()
                                st.subheader("集采标记统计")
                                st.write("注意：空白表示发放明细中有但供应商订单中没有的记录")
                                st.write(procurement_stats)
                                
                                # 单独统计空白数量
                                blank_count = len(marked_df[marked_df['是否集采'] == ''])
                                if blank_count > 0:
                                    st.info(f"共有 {blank_count} 行记录在供应商订单中未找到对应信息（显示为空白）")
                                
                                # 显示详细的集采和非集采数据预览
                                st.subheader("集采标记详情")
                                
                                # 显示集采记录
                                procurement_records = marked_df[marked_df['是否集采'] == '集采']
                                if not procurement_records.empty:
                                    st.write("集采记录:")
                                    st.dataframe(procurement_records[['领用说明', '收货人', '收货人电话', '产品名称', '数量', '是否集采']].head(20))
                                
                                # 显示非集采记录
                                non_procurement_records = marked_df[marked_df['是否集采'] == '非集采']
                                if not non_procurement_records.empty:
                                    st.write("非集采记录:")
                                    st.dataframe(non_procurement_records[['领用说明', '收货人', '收货人电话', '产品名称', '数量', '是否集采']].head(20))
                                
                                # 显示未找到对应信息的记录
                                blank_records = marked_df[marked_df['是否集采'] == '']
                                if not blank_records.empty:
                                    st.write("未在供应商订单中找到的记录:")
                                    st.dataframe(blank_records[['领用说明', '收货人', '收货人电话', '产品名称', '数量', '是否集采']].head(20))
                                
                                # 显示部分标记结果
                                st.subheader("标记结果预览")
                                preview_df = marked_df[['领用说明', '收货人', '收货人电话', '产品名称', '数量', '是否集采']].head(20)
                                st.dataframe(preview_df)
                        
                        # 提供下载
                        download_button(final_result, "处理结果")

elif app_mode == "导入明细到导单模板":
    st.header("导入明细到导单模板")
    st.info("请按步骤操作：1.上传导单模板 2.上传明细文件并选择列映射")
    
    # 初始化session state
    if 'processed_data_list' not in st.session_state:
        st.session_state.processed_data_list = []
    if 'show_import_success' not in st.session_state:
        st.session_state.show_import_success = False
    if 'selected_product' not in st.session_state:
        st.session_state.selected_product = ""
    if 'product_quantity' not in st.session_state:
        st.session_state.product_quantity = 1
    if 'download_triggered' not in st.session_state:
        st.session_state.download_triggered = False
    
    # 步骤1：上传三择导单
    sanze_file = st.file_uploader("1. 上传导单模板文件", type=["xlsx"], key="sanze_file")
    
    if sanze_file:
        try:
            # 创建临时文件以确保可写权限
            with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_file:
                tmp_file.write(sanze_file.getvalue())
                tmp_file_path = tmp_file.name
            
            # 从临时文件读取三择导单
            sanze_df = pd.read_excel(tmp_file_path)
            st.success(f"导单模板已加载（共{len(sanze_df)}行）")
            
            # 确保有客户备注列和货品名称、数量列
            if '客户备注' not in sanze_df.columns:
                sanze_df['客户备注'] = ''
            if '货品名称' not in sanze_df.columns:
                sanze_df['货品名称'] = ''
            if '数量' not in sanze_df.columns:
                sanze_df['数量'] = ''
            if '规格' not in sanze_df.columns:
                sanze_df['规格'] = ''
            
            # 步骤2：处理直邮文件
            st.subheader("2. 处理直邮明细文件")
            uploaded_file = st.file_uploader(
                "上传直邮明细文件", 
                type=["xlsx"],
                key="direct_mail_current"
            )
            
            if uploaded_file:
                # 添加列名行选择功能
                st.subheader("3. 选择列名所在行")
                header_option = st.radio(
                    "请选择列名所在的行（查看下面的预览来确定）",
                    options=[0, 1],
                    format_func=lambda x: f"第{x+1}行",
                    key="header_option"
                )
                
                # 读取前3行数据用于预览
                preview_df = pd.read_excel(uploaded_file, header=None, nrows=3)
                st.write("前三行数据预览（用于确定列名所在行）:")
                st.dataframe(preview_df)
                
                # 根据用户选择的列名行重新读取数据
                df = pd.read_excel(uploaded_file, header=header_option)
                
                st.write(f"使用第{header_option+1}行作为列名后的数据预览:")
                st.dataframe(df.head(5))
                
                # 步骤4：自动识别商品信息
                st.subheader("4. 商品信息识别")
                
                # 使用优化的方法识别商品列
                # 这里需要导入智能识别函数
                from modules.column_detection import identify_product_column_by_content
                
                product_col = identify_product_column_by_content(df, product_list)
                
                # 如果无法通过内容识别，回退到原来的列名匹配方法
                if not product_col:
                    product_columns = [col for col in df.columns if '商品' in col or '货品' in col or '产品' in col or '名称' in col]
                    if product_columns:
                        product_col = product_columns[0]  # 选择第一个找到的商品列
                
                # 获取所有唯一的商品名称
                auto_detected_products = []
                standardized_mapping = {}  # 初始化标准化映射
                
                if product_col:
                    # 对识别出的商品名称进行标准化处理
                    raw_products = df[product_col].dropna().unique()
                    raw_products = [str(p) for p in raw_products if str(p).strip()]
                    
                    # 应用商品名称标准化逻辑
                    for product in raw_products:
                        # 使用与"商品名称标准化"功能相同的逻辑
                        matched = False
                        # 检查特殊映射规则
                        special_mappings = {
                            "硬盒纸抽（130抽）（250506）": "黄金葉  盒装抽纸  （计价单位：盒） 国产定制",
                            "湿巾（10片/包）（250506）": "黄金葉 湿纸巾 10片/包 （计价单位：包） 国产定制"
                        }
                        
                        if product in special_mappings:
                            standardized_mapping[product] = special_mappings[product]
                            matched = True
                        
                        # 检查关键词匹配
                        if not matched:
                            keyword_mappings = {
                                "除螨仪": "奥克斯（AUX） 除螨仪 90W （计价单位：台）",
                                "抽": "国产定制 黄金葉  盒装抽纸  （计价单位：盒）",
                                "塑料袋": "国产定制 黄金葉 环保塑料袋  50个/捆 300个/箱 （计价单位：个）",
                                "两盒装翻盖式礼盒": "国产定制 黄金葉 两盒装翻盖式礼盒 30个/箱 （计价单位：个）",
                                "湿纸巾": "国产定制 黄金葉 湿纸巾 10片/包 （计价单位：包）",
                                "四盒装翻盖式礼盒": "国产定制 黄金葉 四盒装翻盖式礼盒 30个/箱 （计价单位：个）",
                                "四盒装简易封套（天叶品系）": "国产定制 黄金葉 四盒装简易封套（天叶品系） 50个/箱 （计价单位：个）",
                                "天叶叁毫克": "国产定制 黄金葉 天叶叁毫克   两条装纸袋 （计价单位：个）",
                                "五盒装简易封套（常规款）": "国产定制 黄金葉 五盒装简易封套（常规款） 50个/箱 （计价单位：个）",
                                "五盒装简易封套（细支款）": "国产定制 黄金葉 五盒装简易封套（细支款） 50个/箱 （计价单位：个）",
                                "五盒装简易封套（中支款）": "国产定制 黄金葉 五盒装简易封套（中支款） 50个/箱 （计价单位：个）",
                                "数据线三合一充电线100W": "品胜（PISEN） 数据线三合一充电线100W  一拖三 （计价单位：条）",
                                "剃须刀": "有色 剃须刀便携合金电动刮胡刀男士  MINI 2.0 （计价单位：个） 颜色随机"
                            }
                            
                            for keyword, standard_name in keyword_mappings.items():
                                if keyword in product:
                                    standardized_mapping[product] = standard_name
                                    matched = True
                                    break
                        
                        # 使用模糊匹配作为最后手段
                        if not matched:
                            import difflib
                            matches = difflib.get_close_matches(product, product_list, n=1, cutoff=0.6)
                            if matches:
                                standardized_mapping[product] = matches[0]
                            else:
                                # 保持原名称
                                standardized_mapping[product] = product
                    
                    # 应用标准化映射
                    auto_detected_products = list(standardized_mapping.values())
                
                # 显示自动识别结果
                if auto_detected_products:
                    st.success(f"已自动识别到 {len(auto_detected_products)} 种商品:")
                    for i, product in enumerate(auto_detected_products):
                        st.write(f"{i+1}. {product}")
                    
                    # 提供选项让用户选择是否使用自动识别的商品信息
                    use_auto_detection = st.checkbox("使用自动识别的商品信息", value=True)
                    
                    if use_auto_detection:
                        # 使用自动识别的商品信息
                        selected_products = auto_detected_products
                        # 尝试查找规格和数量列
                        spec_columns = [col for col in df.columns if '规格' in col]
                        quantity_columns = [col for col in df.columns if '数量' in col]
                        spec_col = spec_columns[0] if spec_columns else None
                        quantity_col = quantity_columns[0] if quantity_columns else None
                        
                        # 为每种商品获取规格和数量信息
                        product_info = {}
                        for product in selected_products:
                            # 找到该商品的第一行数据
                            # 使用标准化前的名称进行查找
                            original_product_name = next((k for k, v in standardized_mapping.items() if v == product), product)
                            product_rows = df[df[product_col].astype(str) == original_product_name]
                            if not product_rows.empty:
                                first_row = product_rows.iloc[0]
                                # 获取规格信息
                                spec = ""
                                if spec_col and spec_col in df.columns:
                                    spec_value = first_row[spec_col]
                                    spec = str(spec_value) if pd.notna(spec_value) else ""
                                    # 如果规格为空，默认为"个"
                                    if not spec.strip():
                                        spec = "个"
                                
                                # 获取数量信息
                                quantity = 1
                                if quantity_col and quantity_col in df.columns:
                                    try:
                                        quantity_value = first_row[quantity_col]
                                        quantity = int(quantity_value) if pd.notna(quantity_value) else 1
                                    except:
                                        quantity = 1
                                
                                product_info[product] = {
                                    'spec': spec,
                                    'quantity': quantity
                                }
                            else:
                                product_info[product] = {
                                    'spec': '个',
                                    'quantity': 1
                                }
                        
                        st.info("商品规格和数量信息（用于导单模板）:")
                        for product in selected_products:
                            info = product_info[product]
                            st.write(f"- {product}: 规格={info['spec']}, 数量={info['quantity']}")
                    else:
                        # 用户选择手动输入商品信息
                        selected_product = st.selectbox("选择商品", [""] + product_list, 
                                                      index=product_list.index(st.session_state.selected_product) + 1 
                                                      if st.session_state.selected_product in product_list else 0,
                                                      key="product_selector")
                        st.session_state.selected_product = selected_product
                        
                        product_quantity = st.number_input("数量", min_value=1, value=st.session_state.product_quantity, key="quantity_selector")
                        st.session_state.product_quantity = product_quantity
                else:
                    st.warning("未检测到商品信息，请手动选择")
                    # 商品选择界面（只能选择一个商品）
                    selected_product = st.selectbox("选择商品", [""] + product_list, 
                                                  index=product_list.index(st.session_state.selected_product) + 1 
                                                  if st.session_state.selected_product in product_list else 0,
                                                  key="product_selector")
                    st.session_state.selected_product = selected_product
                    
                    product_quantity = st.number_input("数量", min_value=1, value=st.session_state.product_quantity, key="quantity_selector")
                    st.session_state.product_quantity = product_quantity
                
                # 步骤5：选择对应列（保持原版）
                st.subheader("5. 请选择对应列")
                cols = st.columns(2)
                with cols[0]:
                    name_col = st.selectbox("收货人列", df.columns, key="name_col")
                    phone_col = st.selectbox("电话列", df.columns, key="phone_col")
                with cols[1]:
                    address_col = st.selectbox("地址列", df.columns, key="address_col")
                    shop_col = st.selectbox("店名列", df.columns, key="shop_col")
                
                # 处理按钮
                if st.button("处理当前文件"):
                    processed = []
                    # 判断是否使用自动识别的商品信息
                    if 'use_auto_detection' in locals() and use_auto_detection and auto_detected_products:
                        # 使用自动识别的商品信息处理每一行
                        for _, row in df.iterrows():
                            product_name = str(row[product_col]) if product_col in df.columns and pd.notna(row[product_col]) else ""
                            # 使用标准化后的商品名称
                            standardized_product_name = standardized_mapping.get(product_name, product_name)
                            if standardized_product_name in product_info:
                                info = product_info[product_name]
                                item = {
                                    '收货人': str(row[name_col]).strip() if pd.notna(row[name_col]) else '',
                                    '手机': str(row[phone_col]).strip() if pd.notna(row[phone_col]) else '',
                                    '收货地址': str(row[address_col]).strip() if pd.notna(row[address_col]) else '',
                                    '客户备注': str(row[shop_col]).strip() if pd.notna(row[shop_col]) else '',
                                    '商品名称': standardized_product_name,
                                    '商品数量': info['quantity'],
                                    '规格': info['spec']
                                }
                                if item['收货人'] or item['手机']:
                                    processed.append(item)
                    else:
                        # 使用手动选择的商品信息
                        for _, row in df.iterrows():
                            item = {
                                '收货人': str(row[name_col]).strip() if pd.notna(row[name_col]) else '',
                                '手机': str(row[phone_col]).strip() if pd.notna(row[phone_col]) else '',
                                '收货地址': str(row[address_col]).strip() if pd.notna(row[address_col]) else '',
                                '客户备注': str(row[shop_col]).strip() if pd.notna(row[shop_col]) else '',
                                '商品名称': selected_product if 'selected_product' in locals() else st.session_state.selected_product,
                                '商品数量': product_quantity if 'product_quantity' in locals() else st.session_state.product_quantity
                            }
                            if item['收货人'] or item['手机']:
                                processed.append(item)
                    
                    # 将当前处理的文件数据添加到列表中
                    st.session_state.processed_data_list.append({
                        'filename': uploaded_file.name,
                        'data': processed
                    })
                    st.session_state.show_import_success = False
                    st.session_state.download_triggered = False
                    st.success(f"成功处理{len(processed)}行数据")
                    
                    # 创建一个安全的DataFrame用于显示
                    display_data = []
                    for item in processed[:10]:  # 只显示前10行
                        display_item = {
                            '收货人': item['收货人'],
                            '手机': item['手机'],
                            '收货地址': item['收货地址'],
                            '客户备注': item['客户备注'],
                            '商品名称': item['商品名称'],
                            '商品数量': item['商品数量']
                        }
                        if '规格' in item:
                            display_item['规格'] = item['规格']
                        display_data.append(display_item)
                    
                    display_df = pd.DataFrame(display_data)
                    # 确保所有列都是字符串类型，避免PyArrow错误
                    for col in display_df.columns:
                        display_df[col] = display_df[col].astype(str)
                    st.dataframe(display_df)
                
                # 显示已处理的文件列表和导入按钮
                if st.session_state.processed_data_list:
                    st.subheader("已处理的文件列表")
                    for i, item in enumerate(st.session_state.processed_data_list):
                        st.write(f"{i+1}. {item['filename']} ({len(item['data'])} 行数据)")
                    
                    # 添加到三择导单按钮放在已处理文件列表下方
                    if st.button("确认导入到导单模板"):
                        total_imported = 0
                        for processed_item in st.session_state.processed_data_list:
                            processed_data = processed_item['data']
                            for item in processed_data:
                                new_row = {col: '' for col in sanze_df.columns}
                                new_row.update({
                                    '收货人': item['收货人'],
                                    '手机': item['手机'],
                                    '收货地址': item['收货地址'],
                                    '客户备注': item['客户备注'],
                                    '货品名称': item['商品名称'],
                                    '数量': item['商品数量']
                                })
                                # 如果有规格信息，也添加到导单中
                                if '规格' in item:
                                    new_row['规格'] = item['规格']
                                sanze_df.loc[len(sanze_df)] = new_row
                                total_imported += 1
                        
                        st.session_state.show_import_success = True
                        st.session_state.download_triggered = False
                
                if st.session_state.show_import_success:
                    st.success(f"导入完成，导单现有{len(sanze_df)}行")
                    
                    # 下载更新后的文件，文件名精确到秒
                    output = f"导单_更新_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    sanze_df.to_excel(output, index=False)
                    
                    with open(output, "rb") as f:
                        st.download_button("下载更新后的导单", f, file_name=output)
                        
                    # 导入完成后清空已处理数据列表
                    if st.button("清空已处理文件列表并开始下一轮导入"):
                        st.session_state.processed_data_list = []
                        st.session_state.show_import_success = False
                        st.session_state.selected_product = ""
                        st.session_state.product_quantity = 1
                        st.session_state.download_triggered = False
                        st.rerun()
                        
        except Exception as e:
            st.error(f"处理导单时出错: {str(e)}")
            import traceback
            st.error(traceback.format_exc())

elif app_mode == "商品名称标准化":
    st.header("商品名称标准化")
    st.info("此功能用于将发货明细中的不规范商品名称标准化为统一格式")
    
    # 文件上传
    uploaded_file = st.file_uploader("上传发货明细文件", type=["xlsx"], key="standardize_products")
    
    if uploaded_file:
        try:
            # 读取文件
            df = pd.read_excel(uploaded_file)
            
            # 执行标准化
            if st.button("执行标准化"):
                with st.spinner("正在执行商品名称标准化..."):
                    df, mapping = standardize_product_names(df, product_list)
                    
                    # 显示标准化结果
                    st.subheader("标准化结果")
                    st.success("商品名称标准化完成！")
                    
                    # 显示映射关系
                    mapping_df = pd.DataFrame(list(mapping.items()), columns=['原始名称', '标准化名称'])
                    st.write("名称映射关系:")
                    st.dataframe(mapping_df)
                    
                    # 显示标准化后的数据预览
                    st.subheader("标准化后数据预览")
                    st.dataframe(df[['标准化产品名称'] + [col for col in df.columns if col != '标准化产品名称']].head(10))
                    
                    # 统计信息
                    unchanged_count = sum(1 for k, v in mapping.items() if k == v)
                    changed_count = len(mapping) - unchanged_count
                    st.info(f"标准化统计: {changed_count} 个名称已修改, {unchanged_count} 个名称保持不变")
                    
                    # 提供下载
                    download_button(df, "标准化发货明细")
                        
        except Exception as e:
            st.error(f"处理文件时出错: {str(e)}")
            import traceback
            st.error(traceback.format_exc())

elif app_mode == "增强版VLOOKUP":
    st.header("增强版VLOOKUP")
    st.info("此功能允许您从参考表中批量匹配多个列到主表中")
    
    # 初始化session state
    if 'vlookup_processed' not in st.session_state:
        st.session_state.vlookup_processed = False
    if 'vlookup_result' not in st.session_state:
        st.session_state.vlookup_result = None
    
    col1, col2 = st.columns(2)
    
    with col1:
        main_file = st.file_uploader("上传主表文件", type=["xlsx"], key="main_file")
    
    with col2:
        reference_file = st.file_uploader("上传参考表文件", type=["xlsx"], key="reference_file")
    
    # 添加自动检测表头行的变量
    main_header_option = None
    reference_header_option = None
    
    if main_file:
        # 读取前3行数据用于预览
        main_preview_df = pd.read_excel(main_file, header=None, nrows=3)
        st.subheader("主表数据预览")
        st.write("前三行数据（用于确定列名所在行）:")
        st.dataframe(main_preview_df)
        
        # 自动检测最佳表头行
        best_main_header = find_best_header_row(main_preview_df)
        
        # 选择表头行
        main_header_option = st.radio(
            "请选择主表列名所在的行（查看上面的预览来确定）",
            options=[0, 1, 2],
            index=best_main_header if best_main_header is not None else 0,
            format_func=lambda x: f"第{x+1}行",
            key="main_header_option"
        )
    
    if reference_file:
        # 读取前3行数据用于预览
        reference_preview_df = pd.read_excel(reference_file, header=None, nrows=3)
        st.subheader("参考表数据预览")
        st.write("前三行数据（用于确定列名所在行）:")
        st.dataframe(reference_preview_df)
        
        # 自动检测最佳表头行
        best_reference_header = find_best_header_row(reference_preview_df)
        
        # 选择表头行
        reference_header_option = st.radio(
            "请选择参考表列名所在的行（查看上面的预览来确定）",
            options=[0, 1, 2],
            index=best_reference_header if best_reference_header is not None else 0,
            format_func=lambda x: f"第{x+1}行",
            key="reference_header_option"
        )
    
    if main_file and reference_file:
        try:
            # 根据用户选择的表头行读取数据
            main_df = pd.read_excel(main_file, header=main_header_option)
            reference_df = pd.read_excel(reference_file, header=reference_header_option)
            
            st.subheader("主表数据预览")
            st.dataframe(main_df.head(10))
            
            st.subheader("参考表数据预览")
            st.dataframe(reference_df.head(10))
            
            # 选择匹配列
            st.subheader("配置匹配参数")
            
            # 获取所有列名
            main_columns = main_df.columns.tolist()
            reference_columns = reference_df.columns.tolist()
            
            # 选择用于匹配的列（支持多选）
            st.write("选择用于匹配的列（支持多列组合匹配）:")
            match_cols_main = st.multiselect("主表中的匹配列", main_columns, key="match_cols_main", max_selections=len(main_columns))
            match_cols_ref = st.multiselect("参考表中的匹配列（请按与主表相同的顺序选择）", reference_columns, key="match_cols_ref", max_selections=len(reference_columns))
            
            # 验证匹配列选择
            if len(match_cols_main) != len(match_cols_ref):
                st.warning("主表和参考表的匹配列数量必须相同")
            elif len(match_cols_main) == 0:
                st.warning("请至少选择一列用于匹配")
            else:
                # 显示匹配列对应关系
                st.write("匹配列对应关系:")
                match_df = pd.DataFrame({
                    '主表列名': match_cols_main,
                    '参考表列名': match_cols_ref
                })
                st.table(match_df)
                
                # 选择要从参考表添加的列
                st.write("选择要从参考表添加到主表的列:")
                # 排除已用于匹配的列
                available_columns = [col for col in reference_columns if col not in match_cols_ref]
                columns_to_add = st.multiselect("选择列", available_columns, key="columns_to_add")
                
                # 选择匹配方式
                st.write("匹配方式:")
                join_type = st.radio("选择连接方式", 
                                    ["LEFT JOIN (保留主表所有行)", 
                                     "INNER JOIN (只保留两表匹配的行)"], 
                                    key="join_type")
                
                # 处理重复项选项
                st.subheader("重复项处理")
                handle_duplicates = st.radio(
                    "如何处理参考表中的重复匹配记录",
                    ["保留第一条记录", "合并所有记录（可能导致行数增加）"],
                    index=0,
                    key="vlookup_handle_duplicates"
                )
                
                if st.button("执行增强VLOOKUP"):
                    if columns_to_add:
                        with st.spinner("正在执行增强VLOOKUP..."):
                            from modules.enhanced_vlookup import enhanced_vlookup
                            result_df = enhanced_vlookup(
                                main_df, reference_df, match_cols_main, match_cols_ref, 
                                columns_to_add, join_type, handle_duplicates)
                            
                            if result_df is not None:
                                st.success("增强VLOOKUP执行完成！")
                                st.subheader("结果统计")
                                st.write(f"原始主表行数: {len(main_df)}")
                                st.write(f"匹配后结果行数: {len(result_df)}")
                                
                                if len(result_df) > len(main_df):
                                    st.warning("注意：结果行数增加，这是因为某些记录在参考表中有多条匹配记录")
                                
                                st.subheader("结果预览")
                                st.dataframe(result_df.head(20))
                                
                                # 提供下载
                                download_button(result_df, "增强VLOOKUP结果")
                    else:
                        st.warning("请至少选择一列添加到主表中")
        except Exception as e:
            st.error(f"文件读取或处理过程中出现错误: {str(e)}")
            import traceback
            st.error(traceback.format_exc())

elif app_mode == "物流单号匹配":
    st.header("物流单号匹配")
    st.info("此功能用于将物流单号表中的快递公司、单号和额外单号信息匹配到待发货明细表中")
    
    col1, col2 = st.columns(2)
    
    with col1:
        pending_shipment_file = st.file_uploader("上传待发货明细表", type=["xlsx"], key="pending_shipment_file")
    
    with col2:
        logistics_file = st.file_uploader("上传物流单号表", type=["xlsx"], key="logistics_file")
    
    # 添加自动检测表头行的变量
    pending_header_option = None
    logistics_header_option = None
    
    if pending_shipment_file:
        # 读取前3行数据用于预览
        pending_preview_df = pd.read_excel(pending_shipment_file, header=None, nrows=3)
        st.subheader("待发货明细表数据预览")
        st.write("前三行数据（用于确定列名所在行）:")
        st.dataframe(pending_preview_df)
        
        # 自动检测最佳表头行
        best_pending_header = find_best_header_row(pending_preview_df)
        
        # 选择表头行
        pending_header_option = st.radio(
            "请选择待发货明细表列名所在的行（查看上面的预览来确定）",
            options=[0, 1, 2],
            index=best_pending_header if best_pending_header is not None else 0,
            format_func=lambda x: f"第{x+1}行",
            key="pending_header_option"
        )
    
    if logistics_file:
        # 读取前3行数据用于预览
        logistics_preview_df = pd.read_excel(logistics_file, header=None, nrows=3)
        st.subheader("物流单号表数据预览")
        st.write("前三行数据（用于确定列名所在行）:")
        st.dataframe(logistics_preview_df)
        
        # 自动检测最佳表头行
        best_logistics_header = find_best_header_row(logistics_preview_df)
        
        # 选择表头行
        logistics_header_option = st.radio(
            "请选择物流单号表列名所在的行（查看上面的预览来确定）",
            options=[0, 1, 2],
            index=best_logistics_header if best_logistics_header is not None else 0,
            format_func=lambda x: f"第{x+1}行",
            key="logistics_header_option"
        )
    
    if pending_shipment_file and logistics_file:
        try:
            # 根据用户选择的表头行读取数据
            pending_shipment_df = pd.read_excel(pending_shipment_file, header=pending_header_option)
            logistics_df = pd.read_excel(logistics_file, header=logistics_header_option)
            
            st.subheader("待发货明细表数据预览")
            st.dataframe(pending_shipment_df.head(10))
            
            st.subheader("物流单号表数据预览")
            st.dataframe(logistics_df.head(10))
            
            # 显示列名
            st.subheader("列名信息")
            with st.expander("点击展开/收起列名详情"):
                col1, col2 = st.columns(2)
                with col1:
                    st.write("待发货明细表列名:")
                    st.write(list(pending_shipment_df.columns))
                with col2:
                    st.write("物流单号表列名:")
                    st.write(list(logistics_df.columns))
            
            # 自动匹配关键列
            st.subheader("关键列匹配")
            
            # 定义可能的列名
            pending_shipment_name_cols = ['收货人', '收件人', '客户名称', '姓名']
            logistics_name_cols = ['收件人', '收货人', '客户名称', '姓名']
            
            # 自动检测匹配列，优先选择"收货人"
            pending_name_col = None
            logistics_name_col = None
            
            # 优先检查"收货人"列
            if '收货人' in pending_shipment_df.columns:
                pending_name_col = '收货人'
            else:
                for col in pending_shipment_df.columns:
                    if col in pending_shipment_name_cols:
                        pending_name_col = col
                        break
            
            if '收货人' in logistics_df.columns:
                logistics_name_col = '收货人'
            else:
                for col in logistics_df.columns:
                    if col in logistics_name_cols:
                        logistics_name_col = col
                        break
            
            # 显示检测到的列
            st.write("自动检测到的匹配列:")
            with st.expander("点击选择匹配列", expanded=True):
                col1, col2 = st.columns(2)
                with col1:
                    pending_name_select = st.selectbox(
                        "待发货明细表中的收件人列", 
                        pending_shipment_df.columns, 
                        index=pending_shipment_df.columns.tolist().index(pending_name_col) if pending_name_col else 0,
                        key="pending_name_select"
                    )
                with col2:
                    logistics_name_select = st.selectbox(
                        "物流单号表中的收件人列", 
                        logistics_df.columns, 
                        index=logistics_df.columns.tolist().index(logistics_name_col) if logistics_name_col else 0,
                        key="logistics_name_select"
                    )
            
            # 检测重名情况
            from modules.logistics_matching import check_duplicate_names
            has_duplicates, pending_duplicates, logistics_duplicates = check_duplicate_names(
                pending_shipment_df, logistics_df, pending_name_select, logistics_name_select)
            
            # 显示重名检测结果
            st.subheader("重名检测结果")
            if has_duplicates:
                st.warning("检测到存在重名情况，请选择匹配方式")
                if pending_duplicates:
                    st.write(f"待发货明细表中的重名: {', '.join(pending_duplicates[:10])}{'...' if len(pending_duplicates) > 10 else ''}")
                if logistics_duplicates:
                    st.write(f"物流单号表中的重名: {', '.join(logistics_duplicates[:10])}{'...' if len(logistics_duplicates) > 10 else ''}")
                
                # 让用户选择匹配方式
                match_method = st.radio(
                    "请选择匹配方式",
                    ["使用姓名匹配", "使用电话模糊匹配"],
                    index=0
                )
            else:
                st.success("未检测到重名情况，将使用姓名匹配")
                match_method = "使用姓名匹配"
            
            # 如果选择电话匹配，显示电话列选择
            pending_phone_select = None
            logistics_phone_select = None
            
            if match_method == "使用电话模糊匹配":
                st.subheader("电话列匹配")
                # 在待发货明细表中查找电话列
                phone_keywords = ['手机', '电话', '联系方式', '联系电话']
                pending_phone_col = None
                logistics_phone_col = None
                
                for col in pending_shipment_df.columns:
                    if any(keyword in col for keyword in phone_keywords):
                        pending_phone_col = col
                        break
                
                # 在物流单号表中查找电话列
                for col in logistics_df.columns:
                    if any(keyword in col for keyword in phone_keywords):
                        logistics_phone_col = col
                        break
                
                col1, col2 = st.columns(2)
                with col1:
                    pending_phone_select = st.selectbox(
                        "待发货明细表中的电话列",
                        pending_shipment_df.columns,
                        index=pending_shipment_df.columns.tolist().index(pending_phone_col) if pending_phone_col else 0
                    )
                with col2:
                    logistics_phone_select = st.selectbox(
                        "物流单号表中的电话列",
                        logistics_df.columns,
                        index=logistics_df.columns.tolist().index(logistics_phone_col) if logistics_phone_col else 0
                    )
            
            # 选择要添加的列
            st.subheader("选择要添加的列")
            available_columns = [col for col in logistics_df.columns if col not in [logistics_name_select]]
            
            # 设置默认选中的列
            default_columns = []
            common_logistics_columns = ['物流公司', '物流单号', '额外物流单号']
            for col in common_logistics_columns:
                if col in available_columns:
                    default_columns.append(col)
            
            with st.expander("点击选择要添加的列", expanded=True):
                columns_to_add = st.multiselect(
                    "选择要从物流单号表添加到待发货明细表的列",
                    available_columns,
                    default=default_columns,
                    key="columns_to_add"
                )
                st.write("已选择要添加的列:", columns_to_add)
            
            # 处理重复项选项
            st.subheader("重复项处理")
            handle_duplicates = st.radio(
                "如何处理物流单号表中的重复收件人记录",
                ["保留第一条记录", "合并所有记录（可能导致行数增加）"],
                index=0,
                key="handle_duplicates"
            )
            
            # 选择输出列
            st.subheader("输出列选择")
            with st.expander("点击选择输出列", expanded=True):
                # 定义关键列
                key_columns = ['网店订单号', '收货人', '手机', '收货地址', '货品名称', '规格', '数量', '物流公司', '物流单号', '额外物流单号', '发货时间']
                
                # 合并后的所有可能列
                all_possible_columns = list(pending_shipment_df.columns) + columns_to_add
                
                # 确保关键列优先显示
                ordered_columns = []
                # 先添加关键列中存在于数据中的列
                for col in key_columns:
                    if col in all_possible_columns and col not in ordered_columns:
                        ordered_columns.append(col)
                
                # 再添加其他列
                for col in all_possible_columns:
                    if col not in ordered_columns:
                        ordered_columns.append(col)
                
                # 默认选中所有关键列（如果存在）加上从物流表添加的列
                default_output_columns = []
                for col in key_columns:
                    if col in all_possible_columns:
                        default_output_columns.append(col)
                
                # 添加从物流单号表中选择的列（如果尚未包含）
                for col in columns_to_add:
                    if col not in default_output_columns:
                        default_output_columns.append(col)
                
                output_columns = st.multiselect(
                    "选择最终输出的列（可自定义）",
                    ordered_columns,
                    default=default_output_columns,
                    key="output_columns"
                )
            
            if st.button("执行匹配"):
                if columns_to_add:
                    with st.spinner("正在执行匹配..."):
                        try:
                            # 根据用户选择的匹配方式执行匹配
                            if match_method == "使用电话模糊匹配" and pending_phone_select and logistics_phone_select:
                                from modules.logistics_matching import match_logistics_info_fuzzy_phone
                                result_df = match_logistics_info_fuzzy_phone(
                                    pending_shipment_df, logistics_df, pending_name_select,
                                    logistics_name_select, columns_to_add, handle_duplicates,
                                    pending_phone_select=pending_phone_select,
                                    logistics_phone_select=logistics_phone_select)
                            else:
                                # 使用常规匹配（姓名匹配）
                                from modules.logistics_matching import match_logistics_info
                                result_df = match_logistics_info(
                                    pending_shipment_df, logistics_df, pending_name_select,
                                    logistics_name_select, columns_to_add, handle_duplicates,
                                    phone_matching_enabled=False)
                            
                            if result_df is not None:
                                # 只保留指定的输出列
                                if output_columns:
                                    # 检查选定的列是否存在于结果中
                                    existing_output_columns = [col for col in output_columns if col in result_df.columns]
                                    result_df = result_df[existing_output_columns]
                                
                                st.success("匹配完成！")
                                st.subheader("匹配结果统计")
                                st.write(f"原始待发货明细表行数: {len(pending_shipment_df)}")
                                st.write(f"匹配后结果行数: {len(result_df)}")
                                
                                if len(result_df) > len(pending_shipment_df):
                                    st.warning("注意：结果行数增加，这是因为某些收件人在物流单号表中有多条记录")
                                
                                st.subheader("匹配结果预览")
                                st.dataframe(result_df.head(20))
                                
                                # 提供下载
                                download_button(result_df, "发货明细")
                        except Exception as e:
                            st.error(f"匹配过程中出现错误: {str(e)}")
                            import traceback
                            st.error(traceback.format_exc())
                else:
                    st.warning("请至少选择一列添加到待发货明细表中")
        except Exception as e:
            st.error(f"文件读取或处理过程中出现错误: {str(e)}")
            import traceback
            st.error(traceback.format_exc())

elif app_mode == "表合并":
    st.header("表合并")
    st.info("此功能将纵向合并多个表格，并自动移除全空的列")
    
    # 文件上传
    uploaded_files = st.file_uploader(
        "上传需要合并的Excel文件（支持多个文件）",
        type=["xlsx"],
        accept_multiple_files=True,
        key="merge_files"
    )
    
    if uploaded_files and len(uploaded_files) >= 2:
        st.info(f"已选择 {len(uploaded_files)} 个文件")
        
        if st.button("执行合并"):
            with st.spinner("正在执行合并..."):
                try:
                    # 读取所有文件
                    dataframes = []
                    for uploaded_file in uploaded_files:
                        df = pd.read_excel(uploaded_file)
                        # 确保所有列都是字符串类型，避免PyArrow错误
                        for col in df.columns:
                            df[col] = df[col].astype(str)
                        dataframes.append(df)
                    
                    # 纵向合并所有表格
                    result_df = pd.concat(dataframes, ignore_index=True)
                    
                    # 确保所有列都是字符串类型
                    for col in result_df.columns:
                        result_df[col] = result_df[col].astype(str)
                    
                    # 自动过滤全空的列
                    columns_to_keep = []
                    for col in result_df.columns:
                        # 检查列是否全为空值（除了标题）
                        non_empty_values = result_df[col][result_df[col].notna() & (result_df[col] != '') & (result_df[col] != 'nan')]
                        if len(non_empty_values) > 0:
                            columns_to_keep.append(col)
                    
                    # 只保留非空列
                    result_df = result_df[columns_to_keep]
                    
                    st.success("合并完成！")
                    st.write(f"合并后总行数: {len(result_df)}，总列数: {len(result_df.columns)}")
                    
                    st.subheader("合并结果预览")
                    st.dataframe(result_df.head(20))
                    
                    # 提供下载
                    download_button(result_df, "合并结果")
                except Exception as e:
                    st.error(f"合并过程中出现错误: {str(e)}")
                    import traceback
                    st.error(traceback.format_exc())
    else:
        st.info("请至少上传两个Excel文件进行合并")

elif app_mode == "供应商订单分析":
    st.header("供应商订单分析")
    st.info("此功能用于分析供应商订单数据，包括产品、价格、数量、总金额等维度的统计分析")

    # 文件上传
    uploaded_file = st.file_uploader("上传供应商订单文件", type=["xlsx"], key="supplier_analysis")

    if uploaded_file:
        try:
            # 读取文件，正确处理表头
            df = pd.read_excel(uploaded_file, header=1)  # 假设第二行是真正的列名

            # 数据预处理
            # 将数值列转换为正确的数据类型
            numeric_columns = ['数量', '单价(元)', '含税总金额(元)']
            for col in numeric_columns:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

            # 确保非数值列都是字符串类型
            for col in df.columns:
                if col not in numeric_columns:
                    df[col] = df[col].astype(str)

            # 分析部分
            st.subheader("数据分析")

            # 1. 每个供应商的详细分析
            required_columns = ['供应商', '商品名称', '数量', '含税总金额(元)']
            if all(col in df.columns for col in required_columns):
                st.write("### 各供应商详细分析")
                suppliers = df['供应商'].unique()

                # 限制显示的供应商数量，避免界面过于复杂
                max_suppliers = min(10, len(suppliers))
                if len(suppliers) > 10:
                    st.warning(f"供应商数量较多，仅显示前{max_suppliers}个供应商的详细信息")

                # 为每个供应商创建一个可折叠区域
                for i, supplier in enumerate(suppliers[:max_suppliers]):
                    with st.expander(f"{supplier} 详细信息"):
                        st.write(f"##### {supplier}")
                        supplier_data = df[df['供应商'] == supplier]

                        # 该供应商的商品数量和总金额
                        if '商品名称' in df.columns and '数量' in df.columns and '含税总金额(元)' in df.columns:
                            supplier_product_summary = supplier_data.groupby('商品名称').agg({
                                '数量': 'sum',
                                '含税总金额(元)': 'sum'
                            }).sort_values('含税总金额(元)', ascending=False).head(15)  # 限制显示前15个商品

                            st.write("商品数量和总金额:")
                            st.dataframe(supplier_product_summary)

                            # 可视化商品数量和总金额
                            if not supplier_product_summary.empty:
                                # 确保数据类型正确
                                quantity_data = supplier_product_summary['数量'].astype(float)
                                amount_data = supplier_product_summary['含税总金额(元)'].astype(float)

                                col1, col2 = st.columns(2)
                                with col1:
                                    st.write("商品数量分布:")
                                    st.bar_chart(quantity_data)
                                with col2:
                                    st.write("商品总金额分布:")
                                    st.bar_chart(amount_data)

                        # 该供应商的地区分布
                        if '省份' in df.columns and '含税总金额(元)' in df.columns:
                            supplier_province = supplier_data.groupby('省份')['含税总金额(元)'].sum().sort_values(
                                ascending=False)
                            if not supplier_province.empty:
                                st.write("地区订单总金额分布:")
                                # 确保数据类型正确
                                province_data = supplier_province.astype(float)
                                st.bar_chart(province_data)
                                st.dataframe(supplier_province)

            # 2. 综合统计表
            st.write("### 综合统计")
            if all(col in df.columns for col in ['供应商', '商品名称', '数量', '含税总金额(元)']):
                summary_stats = df.groupby('供应商').agg({
                    '含税总金额(元)': ['sum', 'mean', 'count'],
                    '数量': 'sum'
                }).round(2)
                summary_stats.columns = ['总金额', '平均金额', '订单数', '总数量']
                st.dataframe(summary_stats)

                # 添加综合图表
                st.write("### 综合分析图表")
                # 确保数据类型正确
                total_amount_by_supplier = df.groupby('供应商')['含税总金额(元)'].sum().astype(float)
                total_quantity_by_supplier = df.groupby('供应商')['数量'].sum().astype(float)

                col1, col2 = st.columns(2)
                with col1:
                    st.write("各供应商总金额")
                    st.bar_chart(total_amount_by_supplier)
                with col2:
                    st.write("各供应商总数量")
                    st.bar_chart(total_quantity_by_supplier)

            st.success("分析完成！")

        except Exception as e:
            st.error(f"数据分析过程中出现错误: {str(e)}")
            import traceback
            st.error(traceback.format_exc())
    else:
        st.info("请上传供应商订单文件进行分析")

elif app_mode == "商品数量转换":
    st.header("商品数量转换")
    st.info("此功能支持多种单位之间的转换，您可以手动选择转换方向和目标单位")

    # 文件上传
    uploaded_file = st.file_uploader("上传发货明细文件", type=["xlsx"], key="convert_quantities")

    if uploaded_file:
        try:
            # 读取文件
            df = pd.read_excel(uploaded_file)

            st.subheader("原始数据预览")
            st.dataframe(df.head(10))

            # 定义转换规则 - 支持多种单位转换
            conversion_rules = {
                "四盒装翻盖式礼盒": 30,  # 30个/箱
                "盒装抽纸": 20,  # 20个（盒）/箱
                "天叶叁毫克": 100,  # 100个/箱
                "抽纸": 20,  # 20个（盒）/箱
                "环保塑料袋": 300,  # 300个/箱
                "两盒装翻盖式礼盒": 30,  # 30个/箱
                "湿纸巾": 50,  # 50个（包）/箱
                "四盒装简易封套（天叶品系）": 50,  # 50个/箱
                "五盒装简易封套（常规款）": 50,  # 50个/箱
                "五盒装简易封套（细支款）": 50,  # 50个/箱
                "五盒装简易封套（中支款）": 50,  # 50个/箱
                "除螨仪": 1,  # 单个商品，不需要转换倍数
                "数据线三合一充电线100W": 1,  # 单个商品
                "剃须刀便携合金电动刮胡刀男士": 1  # 单个商品
            }

            # 获取当前文件中的所有规格
            unique_units = df['规格'].unique() if '规格' in df.columns else []
            unique_units = [str(unit) for unit in unique_units if pd.notna(unit) and str(unit).strip()]
            
            # 转换设置
            st.subheader("转换设置")
            
            # 让用户选择源单位和目标单位
            col1, col2, col3 = st.columns(3)
            
            with col1:
                source_unit = st.selectbox("选择源单位", [""] + unique_units, key="source_unit")
            
            with col2:
                target_unit = st.selectbox("选择目标单位", ["个", "箱", "件", "条", "包", "台"], key="target_unit")
                
            with col3:
                # 显示当前文件中的规格信息
                st.write("当前文件中的规格:")
                st.write(unique_units)

            # 商品选择功能
            st.subheader("商品选择")
            
            # 获取当前文件中所有商品
            product_col = '货品名称' if '货品名称' in df.columns else '产品名称' if '产品名称' in df.columns else None
            if product_col:
                unique_products = df[product_col].unique()
                unique_products = [str(p) for p in unique_products if pd.notna(p) and str(p).strip()]
                
                # 筛选出有转换规则的商品
                products_with_rules = []
                for product in unique_products:
                    for keyword in conversion_rules.items():
                        if (keyword[0] == product or
                            product.startswith(keyword[0]) or
                            product.endswith(keyword[0]) or
                            keyword[0] in product):
                            products_with_rules.append(product)
                            break
                
                products_with_rules = list(set(products_with_rules))  # 去重
                
                if products_with_rules:
                    # 商品多选
                    selected_products = st.multiselect(
                        f"选择要转换的商品（共{len(products_with_rules)}个可转换商品）",
                        products_with_rules,
                        default=products_with_rules,  # 默认全选
                        key="selected_products"
                    )
                    
                    st.info(f"已选择 {len(selected_products)} 个商品进行转换")
                    
                    # 显示选中商品的转换规则
                    if selected_products:
                        st.subheader("选中商品的转换规则")
                        rule_info = []
                        for product in selected_products:
                            for keyword, multiplier in conversion_rules.items():
                                if (keyword == product or
                                    product.startswith(keyword) or
                                    product.endswith(keyword) or
                                    keyword in product):
                                    # 根据用户选择的单位显示规则
                                    if source_unit and target_unit:
                                        if target_unit == "个":
                                            rule_info.append(f"{product}: 1{source_unit} = {multiplier}{target_unit}")
                                        elif target_unit == "箱":
                                            rule_info.append(f"{product}: {multiplier}{source_unit} = 1{target_unit}")
                                        else:
                                            rule_info.append(f"{product}: 1{source_unit} = {multiplier}{target_unit} (参考)")
                                    break
                        
                        for info in rule_info:
                            st.write(f"• {info}")
                else:
                    st.warning("未找到可转换的商品，请检查商品名称是否包含以下关键词：")
                    st.write(list(conversion_rules.keys()))
                    selected_products = []

                # 转换预览
                if selected_products and source_unit and target_unit:
                    st.subheader("转换预览")
                    
                    # 筛选出选中商品的数据
                    preview_df = df[df[product_col].isin(selected_products)].copy()
                    
                    # 只显示具有指定源单位的行
                    if '规格' in df.columns:
                        preview_df = preview_df[preview_df['规格'].astype(str).str.contains(source_unit, na=False)]
                    
                    if not preview_df.empty:
                        # 计算转换后的数量
                        preview_df['转换后数量'] = preview_df['数量'].copy() if '数量' in df.columns else preview_df.iloc[:, 1].copy()
                        
                        for index, row in preview_df.iterrows():
                            product_name = str(row[product_col])
                            original_quantity = float(row['数量']) if '数量' in df.columns and pd.notna(row['数量']) else 0
                            
                            # 查找匹配的转换规则
                            multiplier = None
                            for keyword, mult in conversion_rules.items():
                                if (keyword == product_name or
                                    product_name.startswith(keyword) or
                                    product_name.endswith(keyword) or
                                    keyword in product_name):
                                    multiplier = mult
                                    break
                            
                            if multiplier and multiplier > 1:
                                # 根据用户选择的转换方向计算
                                if target_unit == "个":  # 箱/件等 → 个
                                    converted_quantity = original_quantity * multiplier
                                else:  # 个 → 箱/件等
                                    converted_quantity = original_quantity / multiplier
                                preview_df.at[index, '转换后数量'] = converted_quantity
                    
                    # 显示预览（只显示相关列）
                    preview_columns = [product_col, '数量' if '数量' in df.columns else df.columns[1], '规格' if '规格' in df.columns else df.columns[2], '转换后数量']
                    if '收货人' in preview_df.columns:
                        preview_columns.insert(0, '收货人')
                    
                    st.dataframe(preview_df[preview_columns].head(10))
                    
                    # 统计信息
                    total_original = preview_df['数量' if '数量' in df.columns else df.columns[1]].sum()
                    total_converted = preview_df['转换后数量'].sum()
                    st.info(f"转换统计：原始总数 {total_original}，转换后总数 {total_converted:.2f}")
                    
                    # 执行转换按钮
                    if st.button("执行数量转换"):
                        with st.spinner("正在执行商品数量转换..."):
                            # 创建转换规则字典以传递给转换函数
                            conversion_config = {
                                'source_unit': source_unit,
                                'target_unit': target_unit,
                                'conversion_rules': conversion_rules
                            }
                            
                            result_df, converted_count = convert_product_quantities_manual(df, selected_products, conversion_config)
                            
                            # 显示转换结果
                            st.success(f"商品数量转换完成！共转换了 {converted_count} 条记录")

                            # 显示转换后的数据预览
                            st.subheader("转换后数据预览")

                            # 选择要显示的列
                            preview_columns = [product_col, '数量' if '数量' in df.columns else df.columns[1], '规格' if '规格' in df.columns else df.columns[2]]
                            new_col_name = f'数量（{target_unit}）'
                            if new_col_name in result_df.columns:
                                preview_columns.append(new_col_name)
                                preview_columns.append('规格（转换后）')
                            if '收货人' in result_df.columns:
                                preview_columns.insert(0, '收货人')
                            
                            # 只显示有转换的记录
                            # 使用numpy数组比较来避免索引对齐问题
                            if new_col_name in result_df.columns and '数量' in result_df.columns:
                                # 检查列的形状以调试维度问题
                                left_series = result_df[new_col_name].astype(float)
                                right_series = result_df['数量'].astype(float)
                                
                                # 确保比较的Series是一维的
                                if len(left_series.shape) > 1:
                                    left_series = left_series.iloc[:, 0]  # 取第一列
                                if len(right_series.shape) > 1:
                                    right_series = right_series.iloc[:, 0]  # 取第一列
                                
                                # 使用pandas的ne方法进行逐元素比较，避免维度问题
                                comparison_series = left_series.ne(right_series)
                                converted_df = result_df[comparison_series]
                                if not converted_df.empty:
                                    st.write("已转换的记录：")
                                    # 修复重复列名问题 - 处理DataFrame中的重复列
                                    converted_df = converted_df.loc[:, ~converted_df.columns.duplicated()]
                                    # 修复preview_columns中的重复列名
                                    unique_preview_columns = []
                                    for col in preview_columns:
                                        if col not in unique_preview_columns:
                                            unique_preview_columns.append(col)
                                    st.dataframe(converted_df[unique_preview_columns].head(20))
                                else:
                                    st.info("没有记录被转换，请检查选择的条件")

                            # 提供下载
                            download_button(result_df, "转换后发货明细")
                elif source_unit and target_unit:
                    st.info("请选择要转换的商品")
                else:
                    st.info("请选择源单位和目标单位")
            else:
                st.warning("未找到合适的商品列，请确保数据中包含'货品名称'或'产品名称'列")
        except Exception as e:
            st.error(f"处理文件时出错: {str(e)}")
            import traceback
            st.error(traceback.format_exc())

# 显示页脚
show_footer()