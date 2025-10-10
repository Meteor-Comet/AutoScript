import pandas as pd
import streamlit as st
import os
import tempfile
from datetime import datetime

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

def compare_data(发货明细_df, 供应商订单_df):
    """
    核对发放明细与供应商订单数据
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