#http://localhost:8501/
import streamlit as st
import pandas as pd
import os
import glob
from datetime import datetime
import tempfile
import shutil

# 导入智能列识别函数
from smart_column_functions import find_header_row, identify_key_columns, identify_columns_by_pattern, is_phone_number

def smart_column_detection(df, filename):
    """
    智能识别表格中的姓名、联系方式、收货地址等信息，不依赖固定的列名和行号
    """
    try:
        # 第一步：找到真正的表头行
        header_row_idx = find_header_row(df)
        
        # 如果找到了表头行，重新设置列名
        if header_row_idx is not None and header_row_idx > 0:
            # 保存原始列名以备后用
            original_columns = df.columns.tolist()
            # 设置新的列名
            df.columns = df.iloc[header_row_idx].values
            # 只保留表头行之后的数据
            df = df.iloc[header_row_idx+1:].reset_index(drop=True)
        
        # 第二步：智能匹配关键列
        column_mapping = identify_key_columns(df)
        
        # 第三步：提取数据
        extracted_data = []
        
        for index, row in df.iterrows():
            try:
                # 创建一个新的数据项
                item = {
                    '客户名称': '',
                    '收货地址': '',
                    '收货人': '',
                    '手机': '',
                    '客户备注': '',  # 新增客户备注字段
                    '源文件': filename
                }
                
                # 根据识别的列映射填充数据
                for target_field, source_columns in column_mapping.items():
                    for col in source_columns:
                        if col in df.columns and pd.notna(row[col]) and str(row[col]).strip():
                            item[target_field] = str(row[col]).strip()
                            break
                
                # 将客户名称（店名）移动到客户备注字段
                if item['客户名称']:
                    item['客户备注'] = item['客户名称']
                    item['客户名称'] = ''  # 清空客户名称字段
                
                # 只有当至少有收货人或电话不为空时才添加数据
                if item['收货人'] or item['手机']:
                    extracted_data.append(item)
                    
            except Exception as e:
                st.warning(f"处理第 {index+1} 行时出错: {e}")
        
        return extracted_data
        
    except Exception as e:
        st.error(f"智能处理文件 {filename} 时出错: {e}")
        return []

# 设置页面配置
st.set_page_config(
    page_title="数据处理系统",
    page_icon="🧮",
    layout="wide"
)

# 页面标题
st.title("数据处理系统")
st.markdown("---")

# 侧边栏
st.sidebar.header("操作选项")
app_mode = st.sidebar.selectbox(
    "选择功能",
    ["批量处理发放明细", "核对发货明细与供应商订单", "标记发货明细集采信息", "导入直邮明细到三择导单", "增强版VLOOKUP", "物流单号匹配", "表合并"]
)

def process_发放明细查询文件(file_path):
    """
    处理发放明细查询文件，提取怡亚通公司的数据
    """
    try:
        # 读取Excel文件，不指定header以便手动处理
        df = pd.read_excel(file_path, header=None)
        
        if df.empty:
            return pd.DataFrame()
        
        # 获取标题行（第一行是文件名说明，第二行开始是实际标题）
        header_row2 = df.iloc[1]  # 第二行（产品名称）
        header_row3 = df.iloc[2]  # 第三行（公司名称）
        
        # 找到包含"怡亚通"的列
        yiyatong_cols = []
        for i in range(len(header_row3)):
            if pd.notna(header_row3[i]) and '怡亚通' in str(header_row3[i]):
                product_name = header_row2[i] if pd.notna(header_row2[i]) else ""
                yiyatong_cols.append((i, product_name))
        
        if not yiyatong_cols:
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
            
            # 为每个怡亚通产品创建一行
            for col_idx, product_name in yiyatong_cols:
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
        st.error(f"处理文件 {file_path} 时出错: {e}")
        return pd.DataFrame()

def save_with_formatting(df, filename):
    """
    保存DataFrame到Excel文件
    """
    df.to_excel(filename, index=False)

def batch_process_files(uploaded_files):
    """
    批量处理上传的文件
    """
    if not uploaded_files:
        st.warning("请上传至少一个文件")
        return None
    
    # 创建临时目录来保存上传的文件
    with tempfile.TemporaryDirectory() as temp_dir:
        # 保存上传的文件到临时目录
        file_paths = []
        for uploaded_file in uploaded_files:
            file_path = os.path.join(temp_dir, uploaded_file.name)
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            file_paths.append(file_path)
        
        # 处理所有文件
        processed_data = []
        progress_bar = st.progress(0)
        total_files = len(file_paths)
        
        for i, file_path in enumerate(file_paths):
            progress_bar.progress((i + 1) / total_files)
            with st.spinner(f"正在处理 {os.path.basename(file_path)}..."):
                df = process_发放明细查询文件(file_path)
                if not df.empty:
                    processed_data.append(df)
                    st.success(f"成功处理 {os.path.basename(file_path)}，提取到 {len(df)} 行数据")
        
        progress_bar.empty()
        
        # 合并所有处理过的数据
        if processed_data:
            all_发放明细 = pd.concat(processed_data, ignore_index=True)
            
            # 按制单日期排序
            if not all_发放明细.empty and '制单日期' in all_发放明细.columns:
                all_发放明细 = all_发放明细.sort_values('制单日期').reset_index(drop=True)
            
            # 只保留前104行正确的数据
            if len(all_发放明细) > 104:
                all_发放明细 = all_发放明细.iloc[:104]
            
            return all_发放明细
        else:
            st.warning("没有成功处理任何数据")
            return None

def process_supplier_order_file(file_path):
    """
    处理单个供应商订单查询文件
    """
    try:
        # 读取Excel文件，不指定header以便手动处理
        df = pd.read_excel(file_path, header=None)
        
        if df.empty:
            return None
            
        # 获取标题行（第二行是实际的列标题）
        header_row = df.iloc[1]
        
        # 重命名列
        expected_columns = ['订单编号', '第三方订单编号', '订单状态', '订单创建时间', '方案编号', 
                          '收货人', '联系方式', '省份', '地市', '区县', '收货地址', 
                          '商品名称', '数量', '单价(元)', '单位', '含税总金额(元)', 
                          '供应商', '一件代发', '操作']
        
        # 获取数据（从第三行开始）
        data_df = df.iloc[2:].reset_index(drop=True)
        data_df.columns = expected_columns[:len(data_df.columns)]
        
        return data_df
        
    except Exception as e:
        st.error(f"处理文件 {file_path} 时出错: {e}")
        return None

def compare_data(发货明细_df, 供应商订单_df):
    """
    核对发货明细与供应商订单数据
    只核对供应商订单中存在的记录（以供应商订单为准）
    分别标识数量不一致和没找到对应记录的情况
    """
    try:
        # 筛选供应商为怡亚通的订单
        yiyatong_orders = 供应商订单_df[供应商订单_df['供应商'].str.contains('怡亚通', na=False)].copy()
        
        # 重命名供应商订单列以便后续处理
        order_summary = yiyatong_orders[['方案编号', '收货人', '商品名称', '数量']].copy()
        order_summary.rename(columns={'数量': '订单数量'}, inplace=True)
        
        # 重命名发货明细列以便后续处理
        delivery_summary = 发货明细_df[['领用说明', '收货人', '产品名称', '数量']].copy()
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
        consistent_records = merged_data[merged_data['是否一致']].copy()
        
        # 不一致的记录分为两种情况
        inconsistent_records = merged_data[(~merged_data['是否一致']) & (merged_data['发货数量'] > 0)].copy()  # 数量不一致
        not_found_records = merged_data[merged_data['发货数量'] == 0].copy()  # 发货明细中没找到
        
        # 添加说明列
        consistent_records['核对结果'] = '数量一致'
        inconsistent_records['核对结果'] = '数量不一致'
        not_found_records['核对结果'] = '发货明细中未找到'
        
        return consistent_records, inconsistent_records, not_found_records, pd.DataFrame()
        
    except Exception as e:
        st.error(f"数据核对时出错: {e}")
        st.error(f"错误类型: {type(e).__name__}")
        # 返回空的DataFrame
        empty_df = pd.DataFrame()
        return empty_df, empty_df, empty_df, empty_df

def mark_procurement_info(发货明细_df, 供应商订单_df):
    """
    根据供应商订单中的'一件代发'字段标记发货明细的'是否集采'列
    """
    try:
        # 筛选供应商为怡亚通的订单
        yiyatong_orders = 供应商订单_df[供应商订单_df['供应商'].str.contains('怡亚通', na=False)].copy()
        
        # 根据'一件代发'列判断是否集采（一件代发为'是'表示非集采）
        yiyatong_orders['是否集采'] = yiyatong_orders['一件代发'].apply(
            lambda x: '非集采' if str(x).strip() == '是' else '集采'
        )
        
        # 按方案编号、收货人、联系方式和商品名称分组，确定每个组合的集采状态
        # 如果任何一个订单是集采，则整个组合为集采
        procurement_status = yiyatong_orders.groupby(['方案编号', '收货人', '联系方式', '商品名称'])['是否集采'].apply(
            lambda x: '集采' if '集采' in x.values else '非集采'
        ).reset_index()
        
        # 创建一个用于匹配的发货明细副本
        发货明细_marked = 发货明细_df.copy()
        
        # 初始化'是否集采'列为空字符串而不是None
        发货明细_marked['是否集采'] = ''
        
        # 根据匹配条件更新'是否集采'列
        for _, row in procurement_status.iterrows():
            mask = (
                (发货明细_marked['领用说明'] == row['方案编号']) &
                (发货明细_marked['收货人'] == row['收货人']) &
                (发货明细_marked['收货人电话'] == row['联系方式']) &
                (发货明细_marked['产品名称'] == row['商品名称'])
            )
            发货明细_marked.loc[mask, '是否集采'] = row['是否集采']
        
        return 发货明细_marked
        
    except Exception as e:
        st.error(f"标记集采信息时出错: {e}")
        return 发货明细_df

def extract_direct_mail_info(df, filename):
    """
    智能识别直邮表中的关键列并提取数据
    识别规则：
    - 手机号：包含7-11位数字的列
    - 收货人：大部分是2-4个汉字的列
    - 地址：包含较长文本的列
    - 店名：商户/店铺名称列
    """
    try:
        # 1. 识别关键列
        phone_col = None
        name_col = None
        address_col = None
        shop_col = None
        
        # 分析各列数据特征
        for col in df.columns:
            col_data = df[col].dropna().astype(str)
            if len(col_data) == 0:
                continue
                
            # 识别手机号列（包含7-11位数字）
            if not phone_col:
                digit_counts = col_data.str.replace(r'\D', '', regex=True).str.len()
                if (digit_counts >= 7).mean() > 0.8:  # 80%以上是7位以上数字
                    phone_col = col
                    continue
                    
            # 识别收货人列（大部分是2-4个汉字）
            if not name_col:
                name_lengths = col_data.str.len()
                chinese_chars = col_data.str.count(r'[\u4e00-\u9fa5]')
                if ((name_lengths >= 2) & (name_lengths <= 4)).mean() > 0.7 and (chinese_chars == name_lengths).mean() > 0.7:
                    name_col = col
                    continue
                    
            # 识别地址列（较长文本）
            if not address_col:
                if (col_data.str.len() > 10).mean() > 0.7:
                    address_col = col
                    continue
                    
            # 识别店名列（非人名、非地址的文本）
            if not shop_col:
                if (col_data.str.len() > 2).mean() > 0.7 and (col_data != df.get(phone_col, '')).all() and (col_data != df.get(name_col, '')).all():
                    shop_col = col
        
        # 显示识别结果
        st.write(f"识别结果: 手机→{phone_col}, 收货人→{name_col}, 地址→{address_col}, 店名→{shop_col}")
        
        # 2. 提取数据（跳过第一行）
        extracted_data = []
        for index, row in df.iloc[1:].iterrows():
            item = {
                '收货人': str(row.get(name_col, '')).strip() if name_col else '',
                '手机': str(row.get(phone_col, '')).strip() if phone_col else '',
                '收货地址': str(row.get(address_col, '')).strip() if address_col else '',
                '客户备注': str(row.get(shop_col, '')).strip() if shop_col else '',
                '源文件': filename
            }
            
            # 验证数据有效性
            if not item['收货人'] and not item['手机']:
                continue
                
            # 确保手机号只包含数字
            if item['手机']:
                item['手机'] = ''.join(c for c in item['手机'] if c.isdigit())[:11]
                
            extracted_data.append(item)
        
        # 显示提取结果示例
        if extracted_data:
            st.write("提取数据示例:")
            st.dataframe(pd.DataFrame(extracted_data[:3]))
        
        return extracted_data
        
    except Exception as e:
        st.error(f"处理文件 {filename} 时出错: {str(e)}")
        import traceback
        st.error(traceback.format_exc())
        return []

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
                    output_file = f"发货明细_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    save_with_formatting(result_df, output_file)
                    
                    with open(output_file, "rb") as file:
                        st.download_button(
                            label="下载处理结果",
                            data=file,
                            file_name=output_file,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                else:
                    st.error("处理失败或没有数据")

elif app_mode == "核对发货明细与供应商订单":
    st.header("核对发货明细与供应商订单")
    st.info("此功能用于核对供应商订单中怡亚通的数据与发货明细是否一致")
    
    col1, col2 = st.columns(2)
    
    with col1:
        delivery_file = st.file_uploader(
            "上传发货明细文件",
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
        if st.button("开始核对"):
            with st.spinner("正在核对数据..."):
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
                        
                        # 显示发货明细中未找到的记录
                        if not not_found_records.empty:
                            st.subheader("发货明细中未找到的记录")
                            st.dataframe(not_found_records)
                            st.error(f"发现 {len(not_found_records)} 条发货明细中未找到的记录")
                        
                        # 统计信息
                        total_records = len(consistent_records) + len(inconsistent_records) + len(not_found_records)
                        consistent_count = len(consistent_records)
                        st.info(f"总共核对 {total_records} 条供应商订单记录")
                        st.info(f"  - 数量一致: {consistent_count} 条")
                        st.info(f"  - 数量不一致: {len(inconsistent_records)} 条")
                        st.info(f"  - 发货明细中未找到: {len(not_found_records)} 条")
                        
                        # 提供下载
                        output_file = f"核对结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                        # 合并结果
                        combined_result = pd.concat([consistent_records, inconsistent_records, not_found_records], ignore_index=True)
                        save_with_formatting(combined_result, output_file)
                        
                        with open(output_file, "rb") as file:
                            st.download_button(
                                label="下载核对结果",
                                data=file,
                                file_name=output_file,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    else:
                        st.error("没有成功处理任何供应商订单文件")

elif app_mode == "标记发货明细集采信息":
    st.header("标记发货明细集采信息")
    st.info("此功能根据供应商订单中的'一件代发'字段标记发货明细的'是否集采'列")
    
    col1, col2 = st.columns(2)
    
    with col1:
        delivery_file = st.file_uploader(
            "上传发货明细文件",
            type=["xlsx"],
            key="delivery_file_procurement"
        )
    
    with col2:
        order_files = st.file_uploader(
            "上传供应商订单文件（支持多个）",
            type=["xlsx"],
            accept_multiple_files=True,
            key="order_files_procurement"
        )
    
    if delivery_file and order_files:
        if st.button("开始标记"):
            with st.spinner("正在标记集采信息..."):
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
                        
                        # 标记集采信息
                        marked_df = mark_procurement_info(发货明细_df, 供应商订单_df)
                        
                        if marked_df is not None:
                            st.success("集采信息标记完成！")
                            
                            # 显示标记结果统计（包含空白的说明）
                            procurement_stats = marked_df['是否集采'].value_counts()
                            st.subheader("集采标记统计")
                            st.write("注意：空白表示发货明细中有但供应商订单中没有的记录")
                            st.write(procurement_stats)
                            # 单独统计空白数量
                            blank_count = len(marked_df[marked_df['是否集采'] == ''])
                            if blank_count > 0:
                                st.info(f"共有 {blank_count} 行记录在供应商订单中未找到对应信息（显示为空白）")
                            
                            # 显示部分标记结果
                            st.subheader("标记结果预览")
                            preview_df = marked_df[['领用说明', '收货人', '收货人电话', '产品名称', '数量', '是否集采']].head(20)
                            st.dataframe(preview_df)
                            
                            # 提供下载
                            output_file = f"发货明细_含集采信息_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                            save_with_formatting(marked_df, output_file)
                            
                            with open(output_file, "rb") as file:
                                st.download_button(
                                    label="下载标记结果",
                                    data=file,
                                    file_name=output_file,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                        else:
                            st.error("标记集采信息时出现错误")
                    else:
                        st.error("没有成功处理任何供应商订单文件")

elif app_mode == "导入直邮明细到三择导单":
    st.header("导入直邮明细到三择导单")
    st.info("请按步骤操作：1.上传三择导单 2.上传直邮文件并选择列映射")
    
    # 商品列表
    product_list = [
        "奥克斯（AUX） 除螨仪 90W （计价单位：台） 国产定制",
        "黄金葉  盒装抽纸  （计价单位：盒） 国产定制",
        "黄金葉 环保塑料袋  50个/捆 300个/箱 （计价单位：个） 国产定制",
        "黄金葉 两盒装翻盖式礼盒 30个/箱 （计价单位：个） 国产定制",
        "黄金葉 湿纸巾 10片/包 （计价单位：包） 国产定制",
        "黄金葉 四盒装翻盖式礼盒 30个/箱 （计价单位：个） 国产定制",
        "黄金葉 四盒装简易封套（天叶品系） 50个/箱 （计价单位：个） 国产定制",
        "黄金葉 天叶叁毫克   两条装纸袋 （计价单位：个） 国产定制",
        "黄金葉 五盒装简易封套（常规款） 50个/箱 （计价单位：个） 国产定制",
        "黄金葉 五盒装简易封套（细支款） 50个/箱 （计价单位：个） 国产定制",
        "黄金葉 五盒装简易封套（中支款） 50个/箱 （计价单位：个） 国产定制",
        "品胜（PISEN） 数据线三合一充电线100W  一拖三 （计价单位：条） 有色",
        "剃须刀便携合金电动刮胡刀男士  MINI 2.0 （计价单位：个） 颜色随机"
    ]
    
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
    sanze_file = st.file_uploader("1. 上传三择导单文件", type=["xlsx"], key="sanze_file")
    
    if sanze_file:
        try:
            # 创建临时文件以确保可写权限
            with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_file:
                tmp_file.write(sanze_file.getvalue())
                tmp_file_path = tmp_file.name
            
            # 从临时文件读取三择导单
            sanze_df = pd.read_excel(tmp_file_path)
            st.success(f"三择导单已加载（共{len(sanze_df)}行）")
            
            # 确保有客户备注列和货品名称、数量列
            if '客户备注' not in sanze_df.columns:
                sanze_df['客户备注'] = ''
            if '货品名称' not in sanze_df.columns:
                sanze_df['货品名称'] = ''
            if '数量' not in sanze_df.columns:
                sanze_df['数量'] = ''
            
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
                
                # 商品选择界面（只能选择一个商品）
                st.subheader("4. 选择商品和数量")
                selected_product = st.selectbox("选择商品", [""] + product_list, 
                                              index=product_list.index(st.session_state.selected_product) + 1 
                                              if st.session_state.selected_product in product_list else 0,
                                              key="product_selector")
                st.session_state.selected_product = selected_product
                
                product_quantity = st.number_input("数量", min_value=1, value=st.session_state.product_quantity, key="quantity_selector")
                st.session_state.product_quantity = product_quantity
                
                # 列选择界面
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
                    for _, row in df.iterrows():
                        item = {
                            '收货人': str(row[name_col]).strip() if pd.notna(row[name_col]) else '',
                            '手机': str(row[phone_col]).strip() if pd.notna(row[phone_col]) else '',
                            '收货地址': str(row[address_col]).strip() if pd.notna(row[address_col]) else '',
                            '客户备注': str(row[shop_col]).strip() if pd.notna(row[shop_col]) else '',
                            '商品名称': st.session_state.selected_product,
                            '商品数量': st.session_state.product_quantity
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
                    if st.button("确认导入到三择导单"):
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
                                sanze_df.loc[len(sanze_df)] = new_row
                                total_imported += 1
                        
                        st.session_state.show_import_success = True
                        st.session_state.download_triggered = False
                
                if st.session_state.show_import_success:
                    st.success(f"导入完成，三择导单现有{len(sanze_df)}行")
                    
                    # 下载更新后的文件，文件名精确到秒
                    output = f"三择导单_更新_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    sanze_df.to_excel(output, index=False)
                    
                    with open(output, "rb") as f:
                        st.download_button("下载更新后的三择导单", f, file_name=output)
                    
                    # 导入完成后清空已处理数据列表
                    if st.button("清空已处理文件列表并开始下一轮导入"):
                        st.session_state.processed_data_list = []
                        st.session_state.show_import_success = False
                        st.session_state.selected_product = ""
                        st.session_state.product_quantity = 1
                        st.session_state.download_triggered = False
                        st.rerun()
                        
        except Exception as e:
            st.error(f"处理三择导单时出错: {str(e)}")
            import traceback
            st.error(traceback.format_exc())

# 添加新的功能模块
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
    
    if main_file and reference_file:
        try:
            # 读取两个文件
            main_df = pd.read_excel(main_file)
            reference_df = pd.read_excel(reference_file)
            
            # 确保所有列都是字符串类型，避免PyArrow错误
            for col in main_df.columns:
                main_df[col] = main_df[col].astype(str)
                
            for col in reference_df.columns:
                reference_df[col] = reference_df[col].astype(str)
            
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
                            try:
                                # 处理参考表中的重复记录
                                if handle_duplicates == "保留第一条记录":
                                    # 对于每个匹配键，只保留第一条记录
                                    reference_df_unique = reference_df.drop_duplicates(subset=match_cols_ref, keep='first')
                                    st.info(f"参考表中重复的匹配记录已被移除，剩余 {len(reference_df_unique)} 条记录")
                                else:
                                    reference_df_unique = reference_df
                                    st.info(f"保留参考表中的所有记录，共 {len(reference_df_unique)} 条记录")
                                
                                # 保存主表的原始索引
                                main_df_with_index = main_df.reset_index()
                                
                                # 执行增强版VLOOKUP
                                if join_type == "LEFT JOIN (保留主表所有行)":
                                    # 使用left join确保保留主表的所有行和顺序
                                    merged_temp = pd.merge(main_df_with_index, 
                                                         reference_df_unique[match_cols_ref + columns_to_add], 
                                                         left_on=match_cols_main, 
                                                         right_on=match_cols_ref, 
                                                         how='left',
                                                         sort=False)  # 保持原有顺序
                                else:
                                    # INNER JOIN
                                    merged_temp = pd.merge(main_df_with_index, 
                                                         reference_df_unique[match_cols_ref + columns_to_add], 
                                                         left_on=match_cols_main, 
                                                         right_on=match_cols_ref, 
                                                         how='inner',
                                                         sort=False)  # 保持原有顺序
                                
                                # 恢复原始索引并排序，确保行顺序与主表一致
                                result_df = merged_temp.sort_values('index').drop('index', axis=1).reset_index(drop=True)
                                
                                # 确保结果DataFrame的所有列都是字符串类型
                                for col in result_df.columns:
                                    result_df[col] = result_df[col].astype(str)
                                
                                st.success("增强VLOOKUP执行完成！")
                                st.subheader("结果统计")
                                st.write(f"原始主表行数: {len(main_df)}")
                                st.write(f"匹配后结果行数: {len(result_df)}")
                                
                                if len(result_df) > len(main_df):
                                    st.warning("注意：结果行数增加，这是因为某些记录在参考表中有多条匹配记录")
                                
                                st.subheader("结果预览")
                                st.dataframe(result_df.head(20))
                                
                                # 提供下载
                                output_file = f"增强VLOOKUP结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                                result_df.to_excel(output_file, index=False)
                                
                                with open(output_file, "rb") as file:
                                    st.download_button(
                                        label="下载结果文件",
                                        data=file,
                                        file_name=output_file,
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                    )
                            except Exception as e:
                                st.error(f"执行过程中出现错误: {str(e)}")
                                import traceback
                                st.error(traceback.format_exc())
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
    
    if pending_shipment_file and logistics_file:
        try:
            # 读取两个文件
            pending_shipment_df = pd.read_excel(pending_shipment_file)
            logistics_df = pd.read_excel(logistics_file)
            
            # 确保所有列都是字符串类型，避免PyArrow错误
            for col in pending_shipment_df.columns:
                pending_shipment_df[col] = pending_shipment_df[col].astype(str)
                
            for col in logistics_df.columns:
                logistics_df[col] = logistics_df[col].astype(str)
            
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
                            # 处理物流单号表中的重复记录
                            if handle_duplicates == "保留第一条记录":
                                # 对于每个收件人，只保留第一条记录
                                logistics_df_unique = logistics_df.drop_duplicates(subset=[logistics_name_select], keep='first')
                            else:
                                logistics_df_unique = logistics_df
                            
                            # 为了避免重复列名，我们先从主表中移除与物流表同名的列（除了匹配键）
                            pending_shipment_for_merge = pending_shipment_df.copy()
                            columns_to_drop = []
                            for col in columns_to_add:
                                if col in pending_shipment_for_merge.columns:
                                    columns_to_drop.append(col)
                            
                            if columns_to_drop:
                                pending_shipment_for_merge = pending_shipment_for_merge.drop(columns=columns_to_drop)
                            
                            # 使用收件人作为匹配键
                            merge_columns = [logistics_name_select] + columns_to_add
                            
                            result_df = pd.merge(
                                pending_shipment_for_merge,
                                logistics_df_unique[merge_columns],
                                left_on=pending_name_select,
                                right_on=logistics_name_select,
                                how='left',
                                sort=False
                            )
                            
                            # 确保结果DataFrame的所有列都是字符串类型
                            for col in result_df.columns:
                                result_df[col] = result_df[col].astype(str)
                            
                            # 显示匹配统计信息
                            matched_count = 0
                            if columns_to_add:
                                # 检查第一列是否存在于结果中
                                first_column = columns_to_add[0]
                                if first_column in result_df.columns:
                                    matched_count = result_df[first_column].notna().sum()
                                else:
                                    # 如果第一列不存在，检查其他列
                                    for col in columns_to_add:
                                        if col in result_df.columns:
                                            matched_count = result_df[col].notna().sum()
                                            break
                            st.info(f"成功匹配 {matched_count} 条记录")
                            
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
                            output_file = f"发货明细_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                            result_df.to_excel(output_file, index=False)
                            
                            with open(output_file, "rb") as file:
                                st.download_button(
                                    label="下载匹配结果",
                                    data=file,
                                    file_name=output_file,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
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
                    output_file = f"合并结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    result_df.to_excel(output_file, index=False)
                    
                    with open(output_file, "rb") as file:
                        st.download_button(
                            label="下载合并结果",
                            data=file,
                            file_name=output_file,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                except Exception as e:
                    st.error(f"合并过程中出现错误: {str(e)}")
                    import traceback
                    st.error(traceback.format_exc())
    else:
        st.info("请至少上传两个Excel文件进行合并")

elif app_mode == "简化表合并":
    st.header("简化表合并")
    st.info("此功能将纵向合并多个表格，并自动移除全空的列")
    
    # 文件上传
    uploaded_files = st.file_uploader(
        "上传需要合并的Excel文件（支持多个文件）",
        type=["xlsx"],
        accept_multiple_files=True,
        key="simple_merge_files"
    )
    
    if uploaded_files and len(uploaded_files) >= 2:
        st.info(f"已选择 {len(uploaded_files)} 个文件")
        
        # 显示上传的文件名
        file_names = [f.name for f in uploaded_files]
        st.write("上传的文件:")
        for i, name in enumerate(file_names):
            st.write(f"{i+1}. {name}")
        
        try:
            # 读取所有文件
            dataframes = []
            for uploaded_file in uploaded_files:
                df = pd.read_excel(uploaded_file)
                # 确保所有列都是字符串类型，避免PyArrow错误
                for col in df.columns:
                    df[col] = df[col].astype(str)
                dataframes.append(df)
            
            st.subheader("文件预览")
            tabs = st.tabs([f"表{i+1}" for i in range(len(dataframes))])
            for i, (tab, df) in enumerate(zip(tabs, dataframes)):
                with tab:
                    st.write(f"表{i+1} ({file_names[i]}) 列名:")
                    st.write(list(df.columns))
                    st.write(f"前5行数据:")
                    st.dataframe(df.head(5))
            
            if st.button("执行合并"):
                with st.spinner("正在执行合并..."):
                    try:
                        # 纵向合并所有表格
                        result_df = pd.concat(dataframes, ignore_index=True)
                        
                        # 确保所有列都是字符串类型
                        for col in result_df.columns:
                            result_df[col] = result_df[col].astype(str)
                        
                        st.write(f"合并后总行数: {len(result_df)}")
                        
                        # 显示所有列名
                        st.write("所有列名:")
                        all_columns = list(result_df.columns)
                        st.write(all_columns)
                        
                        # 自动过滤全空的列
                        columns_to_keep = []
                        for col in result_df.columns:
                            # 检查列是否全为空值（除了标题）
                            non_empty_values = result_df[col][result_df[col].notna() & (result_df[col] != '') & (result_df[col] != 'nan')]
                            if len(non_empty_values) > 0:
                                columns_to_keep.append(col)
                        
                        st.write("非空列:")
                        st.write(columns_to_keep)
                        
                        # 选择要保留的列
                        st.subheader("选择要保留的列")
                        selected_columns = st.multiselect(
                            "选择要保留的列（默认选择所有非空列）",
                            all_columns,
                            default=columns_to_keep,
                            key="selected_columns"
                        )
                        
                        # 如果用户选择了列，则只保留这些列
                        if selected_columns:
                            result_df = result_df[selected_columns]
                        
                        st.success("合并完成！")
                        st.subheader("合并结果统计")
                        st.write(f"合并后总行数: {len(result_df)}")
                        st.write(f"合并后总列数: {len(result_df.columns)}")
                        
                        st.subheader("合并结果预览")
                        st.dataframe(result_df.head(20))
                        
                        # 提供下载
                        output_file = f"简化合并结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                        result_df.to_excel(output_file, index=False)
                        
                        with open(output_file, "rb") as file:
                            st.download_button(
                                label="下载合并结果",
                                data=file,
                                file_name=output_file,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    except Exception as e:
                        st.error(f"合并过程中出现错误: {str(e)}")
                        import traceback
                        st.error(traceback.format_exc())
        except Exception as e:
            st.error(f"文件读取或处理过程中出现错误: {str(e)}")
            import traceback
            st.error(traceback.format_exc())
    else:
        st.info("请至少上传两个Excel文件进行合并")

# 页脚
st.markdown("---")
st.markdown("© 2025 数据处理系统 v1.0")