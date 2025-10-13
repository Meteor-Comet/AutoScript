import pandas as pd
import streamlit as st

def check_duplicate_names(pending_shipment_df, logistics_df, pending_name_col, logistics_name_col):
    """
    检测待发货明细表和物流单号表中的重名情况
    
    返回:
    - has_duplicates: 是否存在重名
    - pending_duplicates: 待发货明细表中的重名列表
    - logistics_duplicates: 物流单号表中的重名列表
    """
    # 检查待发货明细表中的重名
    pending_name_counts = pending_shipment_df[pending_name_col].value_counts()
    pending_duplicates = pending_name_counts[pending_name_counts > 1].index.tolist()
    
    # 检查物流单号表中的重名
    logistics_name_counts = logistics_df[logistics_name_col].value_counts()
    logistics_duplicates = logistics_name_counts[logistics_name_counts > 1].index.tolist()
    
    # 判断是否存在重名
    has_duplicates = len(pending_duplicates) > 0 or len(logistics_duplicates) > 0
    
    return has_duplicates, pending_duplicates, logistics_duplicates

def fuzzy_phone_match(pending_phone, logistics_phone):
    """
    模糊匹配电话号码
    支持部分隐藏的电话号码匹配
    
    Args:
        pending_phone: 待发货明细表中的完整电话号码
        logistics_phone: 物流单号表中的可能部分隐藏的电话号码
    
    Returns:
        bool: 是否匹配
    """
    # 清理电话号码，只保留数字
    pending_digits = ''.join(filter(str.isdigit, str(pending_phone)))
    logistics_digits = ''.join(filter(str.isdigit, str(logistics_phone)))
    
    # 如果任一电话号码为空或太短，不匹配
    if len(pending_digits) < 7 or len(logistics_digits) < 4:
        return False
    
    # 如果物流电话是完整号码，直接比较
    if len(logistics_digits) >= 10:
        return pending_digits == logistics_digits
    
    # 如果物流电话是部分号码（如1580*995），提取可见部分进行匹配
    # 在这种情况下，我们尝试匹配前缀和后缀
    if '*' in str(logistics_phone):
        # 分离前缀和后缀
        parts = str(logistics_phone).split('*')
        if len(parts) == 2:
            prefix = ''.join(filter(str.isdigit, parts[0]))
            suffix = ''.join(filter(str.isdigit, parts[1]))
            
            # 检查前缀和后缀是否匹配
            if (pending_digits.startswith(prefix) and 
                pending_digits.endswith(suffix) and
                len(prefix) + len(suffix) <= len(pending_digits)):
                return True
    
    # 否则使用之前的前缀后缀匹配方法
    logistics_prefix = logistics_digits[:len(logistics_digits)//2] if len(logistics_digits) >= 2 else logistics_digits
    logistics_suffix = logistics_digits[-(len(logistics_digits)//2):] if len(logistics_digits) >= 2 else logistics_digits
    
    pending_prefix = pending_digits[:len(logistics_prefix)] if len(logistics_prefix) <= len(pending_digits) else pending_digits
    pending_suffix = pending_digits[-len(logistics_suffix):] if len(logistics_suffix) <= len(pending_digits) else pending_digits
    
    return pending_prefix == logistics_prefix and pending_suffix == logistics_suffix

def match_logistics_info(pending_shipment_df, logistics_df, pending_name_select, 
                        logistics_name_select, columns_to_add, handle_duplicates, 
                        phone_matching_enabled=False):
    """
    匹配物流信息到待发货明细表
    """
    try:
        # 确保所有列都是字符串类型，避免PyArrow错误
        for col in pending_shipment_df.columns:
            pending_shipment_df[col] = pending_shipment_df[col].astype(str)
            
        for col in logistics_df.columns:
            logistics_df[col] = logistics_df[col].astype(str)
        
        # 处理物流单号表中的重复记录
        if handle_duplicates == "保留第一条记录":
            # 对于每个收件人，只保留第一条记录
            logistics_df_unique = logistics_df.drop_duplicates(subset=[logistics_name_select], keep='first')
        else:
            logistics_df_unique = logistics_df
        
        # 数据预处理：去除收货人姓名中的多余空格（包括前导、尾随和中间的多余空格）
        pending_shipment_df[pending_name_select] = pending_shipment_df[pending_name_select].astype(str).str.strip().str.replace(r'\s+', ' ', regex=True)
        logistics_df_unique[logistics_name_select] = logistics_df_unique[logistics_name_select].astype(str).str.strip().str.replace(r'\s+', ' ', regex=True)
        
        # 如果启用了电话匹配，对电话号码进行预处理
        pending_phone_col = None
        logistics_phone_col = None
        
        if phone_matching_enabled:
            # 在待发货明细表中查找电话列
            phone_keywords = ['手机', '电话', '联系方式', '联系电话']
            for col in pending_shipment_df.columns:
                if any(keyword in col for keyword in phone_keywords):
                    pending_phone_col = col
                    break
            
            # 在物流单号表中查找电话列
            for col in logistics_df_unique.columns:
                if any(keyword in col for keyword in phone_keywords):
                    logistics_phone_col = col
                    break
            
            if pending_phone_col:
                pending_shipment_df[pending_phone_col] = pending_shipment_df[pending_phone_col].astype(str).str.replace(r'[^0-9]', '', regex=True)
            if logistics_phone_col:
                logistics_df_unique[logistics_phone_col] = logistics_df_unique[logistics_phone_col].astype(str).str.replace(r'[^0-9*]', '', regex=True)
        
        # 为了避免重复列名，我们先从主表中移除与物流表同名的列（除了匹配键）
        pending_shipment_for_merge = pending_shipment_df.copy()
        for col in columns_to_add:
            if col in pending_shipment_for_merge.columns and col != pending_name_select:
                pending_shipment_for_merge = pending_shipment_for_merge.drop(columns=[col])
        
        # 使用收件人作为匹配键
        merge_columns = [logistics_name_select] + columns_to_add
        
        if phone_matching_enabled and pending_phone_col and logistics_phone_col:
            merge_columns.append(logistics_phone_col)
        
        if phone_matching_enabled and pending_phone_col and logistics_phone_col:
            # 使用姓名和电话进行匹配
            result_df = pd.merge(
                pending_shipment_for_merge,
                logistics_df_unique[merge_columns],
                left_on=[pending_name_select, pending_phone_col],
                right_on=[logistics_name_select, logistics_phone_col],
                how='left',
                sort=False
            )
        else:
            # 仅使用姓名匹配
            result_df = pd.merge(
                pending_shipment_for_merge,
                logistics_df_unique[merge_columns],
                left_on=pending_name_select,
                right_on=logistics_name_select,
                how='left',
                sort=False
            )
        
        # 处理匹配列重复的问题
        if pending_name_select == logistics_name_select and f"{logistics_name_select}_x" in result_df.columns:
            result_df = result_df.drop(columns=[f"{logistics_name_select}_y"]).rename(columns={f"{logistics_name_select}_x": logistics_name_select})
        
        # 确保结果DataFrame的所有列都是字符串类型
        for col in result_df.columns:
            result_df[col] = result_df[col].astype(str)
            
        return result_df
    except Exception as e:
        st.error(f"匹配过程中出现错误: {str(e)}")
        import traceback
        st.error(traceback.format_exc())
        return None

def match_logistics_info_fuzzy_phone(pending_shipment_df, logistics_df, pending_name_select, 
                                   logistics_name_select, columns_to_add, handle_duplicates,
                                   pending_phone_select=None, logistics_phone_select=None):
    """
    使用模糊电话匹配的物流信息匹配函数
    
    Args:
        pending_shipment_df: 待发货明细表
        logistics_df: 物流单号表
        pending_name_select: 待发货明细表中的姓名列
        logistics_name_select: 物流单号表中的姓名列
        columns_to_add: 要添加的列
        handle_duplicates: 重复项处理方式
        pending_phone_select: 待发货明细表中的电话列
        logistics_phone_select: 物流单号表中的电话列
    """
    try:
        # 确保所有列都是字符串类型，避免PyArrow错误
        for col in pending_shipment_df.columns:
            pending_shipment_df[col] = pending_shipment_df[col].astype(str)
            
        for col in logistics_df.columns:
            logistics_df[col] = logistics_df[col].astype(str)
        
        # 注意：对于电话匹配，我们不预先处理重复项，因为需要保留所有记录以供电话匹配
        logistics_df_unique = logistics_df
        
        # 数据预处理：去除收货人姓名中的多余空格（包括前导、尾随和中间的多余空格）
        pending_shipment_df[pending_name_select] = pending_shipment_df[pending_name_select].astype(str).str.strip().str.replace(r'\s+', ' ', regex=True)
        logistics_df_unique[logistics_name_select] = logistics_df_unique[logistics_name_select].astype(str).str.strip().str.replace(r'\s+', ' ', regex=True)
        
        # 如果提供了电话列，进行电话号码预处理
        if pending_phone_select and logistics_phone_select:
            pending_shipment_df[pending_phone_select] = pending_shipment_df[pending_phone_select].astype(str).str.replace(r'[^0-9]', '', regex=True)
            logistics_df_unique[logistics_phone_select] = logistics_df_unique[logistics_phone_select].astype(str).str.replace(r'[^0-9*]', '', regex=True)
        
        # 按索引顺序匹配，确保每个记录都能匹配到对应的物流信息
        matched_rows = []
        match_stats = {'phone_match': 0, 'name_match': 0, 'no_match': 0}
        
        # 为物流表中的每个姓名创建索引映射
        logistics_name_groups = {}
        for name in logistics_df_unique[logistics_name_select].unique():
            logistics_name_groups[name] = logistics_df_unique[logistics_df_unique[logistics_name_select] == name].reset_index(drop=True)
        
        # 为每个物流记录维护一个使用计数器，处理一个物流记录对应多个发货记录的情况
        logistics_usage_count = {}
        
        # 获取重名列表
        name_counts = pending_shipment_df[pending_name_select].value_counts()
        duplicate_names = name_counts[name_counts > 1].index.tolist()
        
        # 遍历待发货明细表的每一行
        for idx, pending_row in pending_shipment_df.iterrows():
            pending_name = pending_row[pending_name_select]
            pending_phone = pending_row[pending_phone_select] if pending_phone_select else None
            
            best_match = None
            
            # 检查物流表中是否存在该姓名的记录
            if pending_name in logistics_name_groups:
                name_group = logistics_name_groups[pending_name]
                
                # 判断是否是重名
                is_duplicate = pending_name in duplicate_names
                
                if is_duplicate and pending_phone and logistics_phone_select:
                    # 是重名且有电话信息，使用电话匹配
                    matched = False
                    for i, logistics_row in name_group.iterrows():
                        logistics_phone = logistics_row[logistics_phone_select]
                        if fuzzy_phone_match(pending_phone, logistics_phone):
                            best_match = logistics_row
                            match_stats['phone_match'] += 1
                            matched = True
                            
                            # 更新使用计数器而不是移除记录，以支持一个物流记录对应多个发货记录
                            logistics_record_key = (pending_name, str(logistics_phone), i)
                            if logistics_record_key not in logistics_usage_count:
                                logistics_usage_count[logistics_record_key] = 0
                            logistics_usage_count[logistics_record_key] += 1
                            break
                    
                    if not matched:
                        # 电话匹配失败，使用第一个可用的记录
                        if not name_group.empty:
                            best_match = name_group.iloc[0]
                            match_stats['name_match'] += 1
                            
                            # 更新使用计数器
                            logistics_record_key = (pending_name, str(name_group.iloc[0][logistics_phone_select]) if logistics_phone_select else '', 0)
                            if logistics_record_key not in logistics_usage_count:
                                logistics_usage_count[logistics_record_key] = 0
                            logistics_usage_count[logistics_record_key] += 1
                else:
                    # 非重名或没有电话信息，使用第一个可用的记录
                    if not name_group.empty:
                        best_match = name_group.iloc[0]
                        match_stats['name_match'] += 1
                        
                        # 更新使用计数器
                        logistics_record_key = (pending_name, str(name_group.iloc[0][logistics_phone_select]) if logistics_phone_select else '', 0)
                        if logistics_record_key not in logistics_usage_count:
                            logistics_usage_count[logistics_record_key] = 0
                        logistics_usage_count[logistics_record_key] += 1
                    else:
                        match_stats['no_match'] += 1
            else:
                # 物流表中没有该姓名的记录
                match_stats['no_match'] += 1
            
            # 合并行数据
            merged_row = pending_row.to_dict()
            if best_match is not None:
                for col in [logistics_name_select] + columns_to_add:
                    if logistics_phone_select and col == logistics_phone_select:
                        # 特别处理电话列
                        merged_row[col] = best_match[col]
                    elif col != pending_name_select:  # 不覆盖姓名列
                        merged_row[col] = best_match[col]
            else:
                # 确保所有需要的列都存在且为空
                for col in [logistics_name_select] + columns_to_add:
                    if col not in merged_row:
                        merged_row[col] = ''
            
            matched_rows.append(merged_row)
        
        # 显示匹配统计信息
        st.write("匹配统计信息:")
        st.write(f"- 电话匹配: {match_stats['phone_match']}")
        st.write(f"- 姓名匹配: {match_stats['name_match']}")
        st.write(f"- 无匹配: {match_stats['no_match']}")
        
        # 创建结果DataFrame
        result_df = pd.DataFrame(matched_rows)
        
        # 确保结果DataFrame的所有列都是字符串类型
        for col in result_df.columns:
            result_df[col] = result_df[col].astype(str)
            
        return result_df
    except Exception as e:
        st.error(f"匹配过程中出现错误: {str(e)}")
        import traceback
        st.error(traceback.format_exc())
        return None