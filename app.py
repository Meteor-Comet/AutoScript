import streamlit as st
import pandas as pd
import os
import glob
from datetime import datetime
import tempfile
import shutil

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
    ["批量处理发放明细", "核对发货明细与供应商订单", "标记发货明细集采信息", "导入直邮明细到三择导单"]
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
        
        # 遍历数据行并提取信息（从第二行开始，因为第一行是标题）
        for index, row in df.iterrows():
            try:
                # 跳过标题行
                if index == 0:
                    continue
                    
                # 根据不同市场的文件结构处理数据
                if market == "兰州市场":
                    # 兰州市场文件结构: 
                    # 第一列: 序号
                    # 第二列 (Unnamed: 1): 客户名称
                    # 第三列 (Unnamed: 2): 地址
                    # 第四列 (Unnamed: 3): 联系人姓名
                    # 第五列 (Unnamed: 4): 电话
                    item = {
                        '客户名称': row.get('Unnamed: 1', ''),
                        '收货地址': row.get('Unnamed: 2', ''),
                        '收货人': row.get('Unnamed: 3', ''),
                        '手机': str(row.get('Unnamed: 4', '')),
                        '源文件': filename
                    }
                elif market == "酒泉市场":
                    # 酒泉市场文件结构: 序号, 区域, NaN, NaN, 电话, 详细邮寄地址, 访销周期
                    item = {
                        '客户名称': '',  # 酒泉市场文件中没有明确的客户名称列
                        '收货地址': row.get('Unnamed: 6', ''),  # 详细邮寄地址在第七列
                        '收货人': '',   # 酒泉市场文件中没有明确的收货人列
                        '手机': str(row.get('Unnamed: 5', '')),      # 电话在第六列
                        '源文件': filename
                    }
                else:
                    # 默认处理方式
                    item = {
                        '客户名称': row.get('客户名称', row.get('Unnamed: 1', '')),
                        '收货地址': row.get('地址', row.get('Unnamed: 2', '')),
                        '收货人': row.get('联系人姓名', row.get('Unnamed: 3', '')),
                        '手机': str(row.get('电话', row.get('Unnamed: 4', ''))),
                        '源文件': filename
                    }
                
                # 只有当收货人不为空时才添加数据
                if item['收货人'].strip() != '':
                    extracted_data.append(item)
            except Exception as e:
                st.warning(f"处理第 {index+1} 行时出错: {e}")
        
        return extracted_data
        
    except Exception as e:
        st.error(f"处理直邮明细文件 {filename} 时出错: {e}")
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
    st.info("此功能用于将直邮明细表中的姓名、联系方式、收货地址导入到三择导单中")
    
    # 文件上传
    sanze_file = st.file_uploader("上传三择导单文件", type=["xlsx"], key="sanze_file")
    direct_mail_files = st.file_uploader(
        "上传直邮明细文件（支持多个文件）",
        type=["xlsx"],
        accept_multiple_files=True,
        key="direct_mail_files"
    )
    
    if sanze_file and direct_mail_files:
        st.info(f"已选择 {len(direct_mail_files)} 个直邮明细文件")
        
        # 显示上传的文件名
        file_names = [f.name for f in direct_mail_files]
        st.write("上传的直邮明细文件:")
        st.write(file_names)
        
        # 处理按钮
        if st.button("开始导入"):
            with st.spinner("正在处理文件..."):
                try:
                    # 保存上传的文件到临时目录
                    with tempfile.TemporaryDirectory() as temp_dir:
                        # 保存三择导单文件
                        sanze_path = os.path.join(temp_dir, sanze_file.name)
                        with open(sanze_path, "wb") as f:
                            f.write(sanze_file.getbuffer())
                        
                        # 读取三择导单
                        sanze_df = pd.read_excel(sanze_path)
                        st.info(f"读取三择导单文件，共 {len(sanze_df)} 行数据")
                        
                        # 保存直邮明细文件
                        direct_mail_paths = []
                        for uploaded_file in direct_mail_files:
                            file_path = os.path.join(temp_dir, uploaded_file.name)
                            with open(file_path, "wb") as f:
                                f.write(uploaded_file.getbuffer())
                            direct_mail_paths.append(file_path)
                        
                        # 处理所有直邮明细文件
                        all_direct_mail_data = []
                        for i, file_path in enumerate(direct_mail_paths):
                            with st.spinner(f"正在处理 {os.path.basename(file_path)}..."):
                                try:
                                    # 读取直邮明细文件
                                    direct_mail_df = pd.read_excel(file_path)
                                    st.success(f"成功读取 {os.path.basename(file_path)}，共 {len(direct_mail_df)} 行数据")
                                    
                                    # 处理直邮明细数据，提取需要的列
                                    processed_data = extract_direct_mail_info(direct_mail_df, os.path.basename(file_path))
                                    if processed_data:
                                        all_direct_mail_data.extend(processed_data)
                                        st.success(f"成功处理 {os.path.basename(file_path)}，提取到 {len(processed_data)} 行数据")
                                except Exception as e:
                                    st.error(f"处理文件 {os.path.basename(file_path)} 时出错: {e}")
                        
                        # 将处理后的数据添加到三择导单
                        if all_direct_mail_data:
                            st.success(f"总共提取了 {len(all_direct_mail_data)} 行直邮数据")
                            
                            # 创建新的数据行并添加到三择导单
                            new_rows = []
                            for item in all_direct_mail_data:
                                # 创建一个新行，保持三择导单的列结构
                                new_row = {}
                                for col in sanze_df.columns:
                                    new_row[col] = ''  # 默认为空字符串
                                
                                # 填充相关信息，确保正确导入收货人等关键字段
                                new_row['收货人'] = item.get('收货人', '')
                                new_row['手机'] = item.get('手机', '')
                                new_row['收货地址'] = item.get('收货地址', '')
                                new_row['客户名称'] = item.get('客户名称', '')
                                
                                new_rows.append(new_row)
                            
                            # 将新行转换为DataFrame并添加到三择导单
                            if new_rows:
                                new_rows_df = pd.DataFrame(new_rows)
                                # 确保新行的列与三择导单一致
                                new_rows_df = new_rows_df.reindex(columns=sanze_df.columns, fill_value='')
                                # 合并到三择导单（添加到末尾）
                                result_df = pd.concat([sanze_df, new_rows_df], ignore_index=True)
                                
                                st.success(f"导入完成，最终数据共 {len(result_df)} 行")
                                
                                # 显示结果预览（显示完整的三泽导单格式）
                                st.subheader("导入结果预览")
                                st.dataframe(result_df.tail(10))
                                
                                # 提供下载
                                output_file = f"三择导单_导入直邮明细_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                                result_df.to_excel(output_file, index=False)
                                
                                with open(output_file, "rb") as file:
                                    st.download_button(
                                        label="下载导入结果",
                                        data=file,
                                        file_name=output_file,
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                    )
                        else:
                            st.warning("没有成功提取任何直邮数据")
                            
                except Exception as e:
                    st.error(f"处理过程中出错: {e}")

# 页脚
st.markdown("---")
st.markdown("© 2025 数据处理系统 v1.0")