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
        if len(parts) >= 2:
            prefix = ''.join(filter(str.isdigit, parts[0]))
            suffix = ''.join(filter(str.isdigit, parts[-1]))  # 使用最后一个部分作为后缀
            
            # 检查前缀和后缀是否匹配
            if (pending_digits.startswith(prefix) and 
                pending_digits.endswith(suffix) and
                len(prefix) + len(suffix) <= len(pending_digits)):
                return True
    
    # 使用更可靠的前缀后缀匹配方法
    # 提取前4位和后3位进行匹配（符合隐私保护电话匹配方法）
    if len(pending_digits) >= 7 and len(logistics_digits) >= 4:
        pending_prefix = pending_digits[:4] if len(pending_digits) >= 4 else pending_digits
        pending_suffix = pending_digits[-3:] if len(pending_digits) >= 3 else pending_digits
        
        logistics_prefix = logistics_digits[:min(4, len(logistics_digits))]
        logistics_suffix = logistics_digits[max(0, len(logistics_digits)-3):]
        
        # 如果物流号码较短，只需要部分匹配
        if len(logistics_digits) <= 4:
            return pending_digits.startswith(logistics_digits)
        elif len(logistics_digits) <= 7:
            return pending_digits.startswith(logistics_prefix)
        else:
            return pending_prefix == logistics_prefix and pending_suffix == logistics_suffix
    
    return False

def fuzzy_address_match(pending_address, logistics_address):
    """
    模糊匹配地址信息
    支持部分隐藏的地址匹配，也能匹配地址中的关键部分（如街道名）
    
    Args:
        pending_address: 待发货明细表中的完整地址
        logistics_address: 物流单号表中的可能部分隐藏的地址
    
    Returns:
        bool: 是否匹配
    """
    import re
    
    # 转换为字符串并去除首尾空格，统一处理空格
    pending_addr = str(pending_address).strip().replace(' ', '')
    logistics_addr = str(logistics_address).strip().replace(' ', '')
    
    # 如果物流地址是完整地址，直接比较（忽略空格）
    if '*' not in logistics_addr and '**' not in logistics_addr:
        return pending_addr == logistics_addr
    
    # 处理包含*号或**号的地址匹配
    try:
        # 将地址中的**替换为*，统一处理
        normalized_logistics_addr = logistics_addr.replace('**', '*')
        
        # 将地址按*号分割成多个部分
        parts = normalized_logistics_addr.split('*')
        
        # 过滤掉空的部分并去除首尾空格
        parts = [part.strip().replace(' ', '') for part in parts if part.strip()]
        
        # 构建正则表达式模式进行匹配
        pattern = ""
        for i, part in enumerate(parts):
            # 转义特殊字符
            escaped_part = re.escape(part)
            pattern += escaped_part
            # 如果不是最后一部分，添加数字和文字匹配
            if i < len(parts) - 1:
                pattern += r'[\d\u4e00-\u9fa5]*'
        
        # 使用正则表达式进行匹配
        match_result = bool(re.search(pattern, pending_addr))
        # 调试信息
        # st.write(f"DEBUG: 地址匹配 - 待发货: {pending_addr}, 物流: {logistics_addr}, 模式: {pattern}, 匹配结果: {match_result}")
        return match_result
    except re.error as e:
        # 如果正则表达式出错，使用备用方法
        # st.write(f"DEBUG: 地址匹配正则表达式出错: {str(e)}")
        
        # 将地址中的**替换为*，统一处理
        normalized_logistics_addr = logistics_addr.replace('**', '*')
        
        # 将地址按*号分割成多个部分
        parts = normalized_logistics_addr.split('*')
        
        # 过滤掉空的部分并去除首尾空格
        parts = [part.strip().replace(' ', '') for part in parts if part.strip()]
        
        # 检查每个部分是否都在待发货地址中，且按顺序出现
        last_pos = -1
        for part in parts:
            # 特殊处理：如果部分以常见地理标识结尾，使用包含匹配
            geographic_endings = ('省', '市', '区', '县', '镇', '乡', '村', '街', '路', '巷', '道', 
                                '小区', '花园', '公寓', '大厦', '广场', '商场', '写字楼', '园区', '号楼', '单元', '室')
            if part.endswith(geographic_endings):
                # 对于地理标识，使用包含匹配
                pos = pending_addr.find(part)
                if pos == -1:
                    # st.write(f"DEBUG: 地址匹配失败，找不到部分: {part}")
                    return False
            else:
                # 在待发货地址中查找当前部分
                pos = pending_addr.find(part, last_pos + 1)
            
            # 如果找不到或者不是按顺序出现，则不匹配
            if pos == -1 or pos <= last_pos:
                # st.write(f"DEBUG: 地址匹配失败，部分 {part} 未按顺序找到")
                return False
            last_pos = pos
        
        # st.write(f"DEBUG: 地址匹配成功，使用备用方法")
        return True

def extract_address_keywords(address):
    """
    从地址中提取关键词，如街道名、小区名等
    
    Args:
        address: 地址字符串
    
    Returns:
        list: 关键词列表
    """
    # 常见的地址关键词后缀
    suffixes = ['路', '街', '巷', '道', '大街', '小区', '花园', '公寓', '大厦', '广场', '商场', '写字楼', '园区', 
                '区', '县', '镇', '乡', '村', '组', '栋', '单元', '室', '楼', '号', '层', '座', '苑', '庭', '园']
    
    keywords = []
    
    # 提取包含关键词后缀的部分
    for suffix in suffixes:
        # 查找包含后缀的词
        import re
        pattern = r'[\u4e00-\u9fa50-9]+?' + suffix
        matches = re.findall(pattern, address)
        keywords.extend(matches)
    
    # 如果没有找到关键词，返回整个地址作为关键词
    if not keywords:
        keywords = [address]
    
    return keywords

def extract_multiple_products(product_string):
    """
    从物流表的一个单元格中提取多个商品关键词
    
    Args:
        product_string: 物流表中的商品信息字符串
        
    Returns:
        list: 提取出的商品关键词列表
    """
    if not product_string or product_string == 'nan':
        return []
    
    # 常见的分隔符
    separators = ['，', ',', '、', ';', '；', '和', '+', '以及', '及']
    
    # 尝试使用分隔符分割
    products = [product_string]
    for sep in separators:
        temp_products = []
        for product in products:
            temp_products.extend(product.split(sep))
        products = temp_products
    
    # 清理每个产品名称
    cleaned_products = []
    for product in products:
        cleaned = product.strip()
        if cleaned and cleaned != 'nan':
            # 去除常见的数量描述
            # 例如："商品A(2个)" -> "商品A"
            import re
            cleaned = re.sub(r'\(.*?\)', '', cleaned)
            cleaned = re.sub(r'（.*?）', '', cleaned)
            cleaned = re.sub(r'\[.*?\]', '', cleaned)
            cleaned = re.sub(r'【.*?】', '', cleaned)
            cleaned = cleaned.strip()
            if cleaned:
                cleaned_products.append(cleaned)
    
    return cleaned_products

def extract_multiple_products_with_logistics(product_string, logistics_row, logistics_product_select, columns_to_add):
    """
    从物流表的一个单元格中提取多个商品关键词，并保持与物流信息的对应关系
    
    Args:
        product_string: 物流表中的商品信息字符串
        logistics_row: 物流表行数据
        logistics_product_select: 物流表中的商品列名
        columns_to_add: 需要添加的列名列表
        
    Returns:
        list: 包含商品和对应物流信息的字典列表
    """
    if not product_string or product_string == 'nan':
        return []
    
    # 常见的分隔符
    separators = ['，', ',', '、', ';', '；', '和', '+', '以及', '及']
    
    # 尝试使用分隔符分割
    products = [product_string]
    for sep in separators:
        temp_products = []
        for product in products:
            temp_products.extend(product.split(sep))
        products = temp_products
    
    # 清理每个产品名称
    cleaned_products = []
    for product in products:
        cleaned = product.strip()
        if cleaned and cleaned != 'nan':
            # 去除常见的数量描述
            # 例如："商品A(2个)" -> "商品A"
            import re
            cleaned = re.sub(r'\(.*?\)', '', cleaned)
            cleaned = re.sub(r'（.*?）', '', cleaned)
            cleaned = re.sub(r'\[.*?\]', '', cleaned)
            cleaned = re.sub(r'【.*?】', '', cleaned)
            cleaned = cleaned.strip()
            if cleaned:
                cleaned_products.append(cleaned)
    
    # 为每个商品创建对应的物流信息
    result = []
    num_products = len(cleaned_products)
    
    if num_products <= 1:
        # 如果只有一个或没有商品，直接返回原始行
        return [{'product': product, 'logistics_info': logistics_row.to_dict()} for product in cleaned_products]
    
    # 如果有多个商品，我们需要为每个商品分配对应的物流信息
    # 这里我们假设物流信息中的其他字段（如单号）可能也按顺序对应商品
    for i, product in enumerate(cleaned_products):
        logistics_info = {}
        for col in logistics_row.index:
            if col == logistics_product_select:
                # 商品列直接使用当前商品
                logistics_info[col] = product
            else:
                # 其他列保持原样
                logistics_info[col] = logistics_row[col]
        result.append({'product': product, 'logistics_info': logistics_info})
    
    return result

def fuzzy_product_match(pending_product, logistics_product):
    """
    模糊匹配商品信息
    
    Args:
        pending_product: 待发货明细表中的商品名称
        logistics_product: 物流单号表中的商品名称
    
    Returns:
        bool: 是否匹配
    """
    pending_product = str(pending_product).strip()
    logistics_product = str(logistics_product).strip()
    
    # 处理空值或nan值
    if not pending_product or pending_product == 'nan' or not logistics_product or logistics_product == 'nan':
        return False
    
    # 如果两个商品名称完全相同，直接返回True
    if pending_product == logistics_product:
        return True
    
    # 如果物流商品名称不包含特殊标记，直接比较
    if '[' not in logistics_product and '(' not in logistics_product and '（' not in logistics_product:
        # 使用difflib进行相似度比较
        import difflib
        similarity = difflib.SequenceMatcher(None, pending_product, logistics_product).ratio()
        # 如果相似度超过0.8，认为匹配
        if similarity >= 0.8:
            return True
    
    # 定义商品唯一关键词映射
    product_categories = {
        '锅具': ['锅具', '炒锅', '煎锅', '汤锅'],
        '除螨仪': ['除螨仪'],
        '数据线': ['数据线', '充电线'],
        '剃须刀': ['剃须刀'],
        '抽纸': ['抽纸'],
        '湿巾': ['湿巾'],
        '封套': ['封套'],
        '礼盒': ['礼盒']
    }
    
    # 检查两个商品是否属于同一类别
    pending_category = None
    logistics_category = None
    
    # 确定待发货商品的类别
    for category, keywords in product_categories.items():
        if any(keyword in pending_product for keyword in keywords):
            pending_category = category
            break
    
    # 确定物流商品的类别
    for category, keywords in product_categories.items():
        if any(keyword in logistics_product for keyword in keywords):
            logistics_category = category
            break
    
    # 如果两个商品都不属于已知类别，尝试基于关键词的部分匹配
    if pending_category is None or logistics_category is None:
        # 使用更宽松的匹配方式
        # 检查是否有共同的关键词
        pending_words = set(pending_product)
        logistics_words = set(logistics_product)
        common_words = pending_words.intersection(logistics_words)
        
        # 如果有足够多的共同字符，认为匹配
        if len(common_words) > min(len(pending_product), len(logistics_product)) * 0.3:
            return True
        return False
    
    # 如果两个商品属于不同类别，返回False
    if pending_category != logistics_category:
        return False
    
    # 如果两个商品属于同一类别，检查品牌是否相同
    import re
    pending_brand_match = re.search(r'([\u4e00-\u9fa5a-zA-Z]+)[(（].*?[)）]', pending_product)
    logistics_brand_match = re.search(r'([\u4e00-\u9fa5a-zA-Z]+)[(（].*?[)）]', logistics_product)
    
    # 如果都有品牌信息，比较品牌
    if pending_brand_match and logistics_brand_match:
        pending_brand = pending_brand_match.group(1)
        logistics_brand = logistics_brand_match.group(1)
        return pending_brand == logistics_brand
    
    # 如果都没有品牌信息，但属于同一类别，返回True
    if not pending_brand_match and not logistics_brand_match:
        return True
    
    # 其他情况返回False
    return False

def fuzzy_product_match_multi(pending_product, logistics_product_cell):
    """
    支持物流单元格中多个商品的模糊匹配
    
    Args:
        pending_product: 待发货明细表中的单一商品名称
        logistics_product_cell: 物流单号表中的商品单元格（可能包含多个商品）
        
    Returns:
        bool: 是否匹配
    """
    # 提取物流单元格中的多个商品
    logistics_products = extract_multiple_products(logistics_product_cell)
    
    # 如果没有提取到商品，使用原来的匹配方式
    if not logistics_products:
        return fuzzy_product_match(pending_product, logistics_product_cell)
    
    # 检查待发货商品是否与物流单元格中的任何一个商品匹配
    for logistics_product in logistics_products:
        if fuzzy_product_match(pending_product, logistics_product):
            return True
    
    return False

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
                                   pending_phone_select=None, logistics_phone_select=None,
                                   pending_address_select=None, logistics_address_select=None,
                                   pending_product_select=None, logistics_product_select=None):
    """
    使用模糊电话匹配的物流信息匹配函数（优化版）
    """
    try:
        # 确保所有列都是字符串类型，避免PyArrow错误
        for col in pending_shipment_df.columns:
            pending_shipment_df[col] = pending_shipment_df[col].astype(str)
            
        for col in logistics_df.columns:
            logistics_df[col] = logistics_df[col].astype(str)
        
        # 注意：对于电话匹配，我们不预先处理重复项，因为需要保留所有记录以供电话匹配
        logistics_df_unique = logistics_df.copy()
        
        # 数据预处理：去除收货人姓名中的多余空格（包括前导、尾随和中间的多余空格）
        pending_shipment_df[pending_name_select] = pending_shipment_df[pending_name_select].astype(str).str.strip().str.replace(r'\s+', ' ', regex=True)
        logistics_df_unique[logistics_name_select] = logistics_df_unique[logistics_name_select].astype(str).str.strip().str.replace(r'\s+', ' ', regex=True)
        
        # 如果提供了电话列，进行电话号码预处理
        if pending_phone_select and logistics_phone_select:
            pending_shipment_df[pending_phone_select] = pending_shipment_df[pending_phone_select].astype(str).str.replace(r'[^0-9]', '', regex=True)
            logistics_df_unique[logistics_phone_select] = logistics_df_unique[logistics_phone_select].astype(str).str.replace(r'[^0-9*]', '', regex=True)
            
        # 如果提供了地址列，进行地址预处理，去除多余空格
        if pending_address_select and logistics_address_select:
            pending_shipment_df[pending_address_select] = pending_shipment_df[pending_address_select].astype(str).str.strip().str.replace(r'\s+', '', regex=True)
            logistics_df_unique[logistics_address_select] = logistics_df_unique[logistics_address_select].astype(str).str.strip().str.replace(r'\s+', '', regex=True)
        
        # 按索引顺序匹配，确保每个记录都能匹配到对应的物流信息
        matched_rows = []
        match_stats = {'total_matched': 0, 'name_only': 0, 'phone_refined': 0, 'address_refined': 0, 'product_refined': 0, 'no_match': 0}
        
        # 为物流表中的每个姓名创建索引映射
        logistics_name_groups = {}
        for name in logistics_df_unique[logistics_name_select].unique():
            # 修复：保留原始索引，不使用drop=True，以维持原始顺序
            logistics_name_groups[name] = logistics_df_unique[logistics_df_unique[logistics_name_select] == name]  # 移除reset_index(drop=True)
        
        # 创建物流记录使用状态跟踪（支持多次使用）
        logistics_usage = {}  # {(name, phone, address, index): usage_count}
        logistics_max_usage = {}  # {(name, phone, address, index): max_allowed_usage}
        logistics_product_mapping = {}  # {(name, phone, address, index): [{'product': product, 'logistics_info': info}]}

        # 预处理物流记录，计算每个记录可以被使用的最大次数
        for pending_name in logistics_name_groups:
            name_group_logistics = logistics_name_groups[pending_name]
            # 修复：使用原始索引，确保顺序正确
            for i, (idx, logistics_row) in enumerate(name_group_logistics.iterrows()):  # 使用iterrows()获取原始索引
                logistics_phone = logistics_row[logistics_phone_select] if logistics_phone_select else ''
                logistics_address = logistics_row[logistics_address_select] if logistics_address_select else ''
                # 修复：使用原始索引idx而不是循环索引i，确保唯一标识符正确
                logistics_record_key = (pending_name, str(logistics_phone), logistics_address, idx)
                
                # 如果提供了商品列，检查物流单元格中是否包含多个商品
                if logistics_product_select and logistics_product_select in logistics_row:
                    logistics_product_cell = logistics_row[logistics_product_select]
                    # 提取多个商品及其对应的物流信息
                    product_logistics_list = extract_multiple_products_with_logistics(
                        logistics_product_cell, logistics_row, logistics_product_select, columns_to_add)
                    logistics_product_mapping[logistics_record_key] = product_logistics_list
                    # 允许使用次数等于商品数
                    logistics_max_usage[logistics_record_key] = max(1, len(product_logistics_list))
                else:
                    logistics_max_usage[logistics_record_key] = 1
                    logistics_product_mapping[logistics_record_key] = [{'product': '', 'logistics_info': logistics_row.to_dict()}]

        # 按原始顺序处理每条待发货记录，确保顺序不变
        for idx, pending_row in pending_shipment_df.iterrows():
            pending_name = pending_row[pending_name_select]
            pending_phone = pending_row[pending_phone_select] if pending_phone_select else None
            pending_product = pending_row[pending_product_select] if pending_product_select else None
            pending_address = pending_row[pending_address_select] if pending_address_select else None

            best_match = None
            matched_by = ""
            matched_product_info = None

            # 获取该姓名对应的物流记录组
            name_group_logistics = logistics_name_groups.get(pending_name, pd.DataFrame())
            
            # 创建候选匹配列表，用于后续排序
            candidate_matches = []

            # 尝试电话匹配
            if pending_phone and pending_phone_select and logistics_phone_select and not name_group_logistics.empty:
                # 修复：使用iterrows()获取原始索引
                for i, (logistics_idx, logistics_row) in enumerate(name_group_logistics.iterrows()):
                    logistics_phone = logistics_row[logistics_phone_select]
                    
                    if fuzzy_phone_match(pending_phone, logistics_phone):
                        # 检查是否已被使用
                        logistics_address = logistics_row[logistics_address_select] if logistics_address_select else ''
                        # 修复：使用原始索引logistics_idx而不是循环索引i
                        logistics_record_key = (pending_name, str(logistics_phone), logistics_address, logistics_idx)
                        usage_count = logistics_usage.get(logistics_record_key, 0)
                        max_allowed = logistics_max_usage.get(logistics_record_key, 1)
                        product_mapping = logistics_product_mapping.get(logistics_record_key, [])

                        # 计算匹配分数 - 提高商品匹配的权重
                        match_score = 10  # 基础电话匹配分数

                        # 地址匹配加分
                        if pending_address and logistics_address_select:
                            logistics_address = logistics_row[logistics_address_select]
                            if fuzzy_address_match(pending_address, logistics_address):
                                match_score += 5

                        # 商品匹配加分 - 提高权重以确保商品匹配优先级更高
                        if pending_product and product_mapping:
                            for product_info in product_mapping:
                                if fuzzy_product_match(pending_product, product_info['product']):
                                    match_score += 10  # 提高商品匹配分数到10分
                                    break

                        candidate_matches.append({
                            'logistics_row': logistics_row.to_dict(),
                            'record_key': logistics_record_key,
                            'usage_count': usage_count,
                            'max_allowed': max_allowed,
                            'product_mapping': product_mapping,
                            'score': match_score,
                            'matched_by': 'phone'
                        })

            # 如果电话匹配失败，尝试地址匹配
            if best_match is None and pending_address and pending_address_select and logistics_address_select and not name_group_logistics.empty:
                # 修复：使用iterrows()获取原始索引
                for i, (logistics_idx, logistics_row) in enumerate(name_group_logistics.iterrows()):
                    logistics_address = logistics_row[logistics_address_select]
                    
                    if fuzzy_address_match(pending_address, logistics_address):
                        # 检查是否已在候选列表中
                        logistics_phone = logistics_row[logistics_phone_select] if logistics_phone_select else ''
                        logistics_address = logistics_row[logistics_address_select] if logistics_address_select else ''
                        # 修复：使用原始索引logistics_idx而不是循环索引i
                        logistics_record_key = (pending_name, str(logistics_phone), logistics_address, logistics_idx)
                        
                        # 检查是否已经通过电话匹配添加到候选列表
                        already_added = False
                        for candidate in candidate_matches:
                            if candidate['record_key'] == logistics_record_key:
                                already_added = True
                                # 增加地址匹配分数
                                candidate['score'] += 5
                                # 商品匹配加分 - 提高权重
                                if pending_product and candidate['product_mapping']:
                                    for product_info in candidate['product_mapping']:
                                        if fuzzy_product_match(pending_product, product_info['product']):
                                            candidate['score'] += 10  # 提高商品匹配分数到10分
                                            break
                                candidate['matched_by'] = 'address' if candidate['matched_by'] == 'phone' else candidate['matched_by']
                                break

                        if not already_added:
                            usage_count = logistics_usage.get(logistics_record_key, 0)
                            max_allowed = logistics_max_usage.get(logistics_record_key, 1)
                            product_mapping = logistics_product_mapping.get(logistics_record_key, [])

                            # 计算匹配分数
                            match_score = 7  # 基础地址匹配分数

                            # 商品匹配加分 - 提高权重
                            if pending_product and product_mapping:
                                for product_info in product_mapping:
                                    if fuzzy_product_match(pending_product, product_info['product']):
                                        match_score += 10  # 提高商品匹配分数到10分
                                        break

                            candidate_matches.append({
                                'logistics_row': logistics_row.to_dict(),
                                'record_key': logistics_record_key,
                                'usage_count': usage_count,
                                'max_allowed': max_allowed,
                                'product_mapping': product_mapping,
                                'score': match_score,
                                'matched_by': 'address'
                            })

            # 如果地址匹配也失败，尝试商品匹配（支持多商品匹配）
            if best_match is None and pending_product and pending_product_select and logistics_product_select and not name_group_logistics.empty:
                # 修复：使用iterrows()获取原始索引
                for i, (logistics_idx, logistics_row) in enumerate(name_group_logistics.iterrows()):
                    logistics_product_cell = logistics_row[logistics_product_select]
                    
                    if fuzzy_product_match_multi(pending_product, logistics_product_cell):
                        # 检查是否已在候选列表中
                        logistics_phone = logistics_row[logistics_phone_select] if logistics_phone_select else ''
                        logistics_address = logistics_row[logistics_address_select] if logistics_address_select else ''
                        # 修复：使用原始索引logistics_idx而不是循环索引i
                        logistics_record_key = (pending_name, str(logistics_phone), logistics_address, logistics_idx)

                        # 检查是否已经添加到候选列表
                        already_added = False
                        for candidate in candidate_matches:
                            if candidate['record_key'] == logistics_record_key:
                                already_added = True
                                # 增加商品匹配分数
                                candidate['score'] += 10  # 提高商品匹配分数到10分
                                break

                        if not already_added:
                            usage_count = logistics_usage.get(logistics_record_key, 0)
                            max_allowed = logistics_max_usage.get(logistics_record_key, 1)
                            product_mapping = logistics_product_mapping.get(logistics_record_key, [])

                            # 计算匹配分数 - 商品匹配作为主要匹配方式时给予更高分数
                            match_score = 15  # 基础商品匹配分数（提高到15分）

                            candidate_matches.append({
                                'logistics_row': logistics_row.to_dict(),
                                'record_key': logistics_record_key,
                                'usage_count': usage_count,
                                'max_allowed': max_allowed,
                                'product_mapping': product_mapping,
                                'score': match_score,
                                'matched_by': 'product'
                            })

            # 从候选匹配中选择最佳匹配
            best_match = None
            matched_by = ""
            matched_product_info = None

            if candidate_matches:
                # 为了解决同名同电话同地址不同商品但物流单号相反的问题，我们需要进一步优化匹配逻辑
                # 使用多级排序策略，确保匹配结果与待发货明细顺序一致
                
                # 首先按匹配分数降序排序，然后按原始索引升序排序
                # 这样可以在保证匹配质量的同时，严格按照原始顺序处理
                candidate_matches.sort(key=lambda x: (-x['score'], x['record_key'][3]))

                # 选择第一个未达到最大使用次数的候选
                for candidate in candidate_matches:
                    if candidate['usage_count'] < candidate['max_allowed']:
                        best_match = candidate['logistics_row']
                        matched_by = candidate['matched_by']
                        
                        # 查找匹配的商品信息
                        if pending_product and candidate['product_mapping']:
                            # 为了解决同名同电话同地址不同商品但物流单号相反的问题
                            # 我们需要更加精确地匹配商品，而不是依赖使用计数器
                            matched_product_info = None
                            best_product_match_score = 0
                            
                            # 遍历所有商品信息，找到最佳匹配
                            for product_info in candidate['product_mapping']:
                                if fuzzy_product_match(pending_product, product_info['product']):
                                    # 计算商品匹配分数（基于字符串相似度）
                                    import difflib
                                    similarity = difflib.SequenceMatcher(None, pending_product, product_info['product']).ratio()
                                    if similarity > best_product_match_score:
                                        best_product_match_score = similarity
                                        matched_product_info = product_info
                            
                            # 如果没有找到基于内容的匹配，再考虑使用使用次数
                            if matched_product_info is None and candidate['usage_count'] < len(candidate['product_mapping']):
                                matched_product_info = candidate['product_mapping'][candidate['usage_count']]
                        
                        # 更新使用计数器
                        logistics_usage[candidate['record_key']] = candidate['usage_count'] + 1
                        
                        match_stats['total_matched'] += 1
                        if matched_by == "phone":
                            match_stats['phone_refined'] += 1
                        elif matched_by == "address":
                            match_stats['address_refined'] += 1
                        elif matched_by == "product":
                            match_stats['product_refined'] += 1
                        else:
                            match_stats['name_only'] += 1
                        break

            # 如果所有匹配都失败，尝试使用未达到最大使用次数的物流记录
            # 为了解决同名同电话但物流单号相反的问题，这里也需要按索引顺序处理
            if best_match is None and not name_group_logistics.empty:
                # 先按索引排序，确保按原始顺序处理
                # 修复：使用原始索引进行排序
                sorted_logistics = name_group_logistics.sort_index()
                # 修复：使用iterrows()获取原始索引
                for i, (logistics_idx, logistics_row) in enumerate(sorted_logistics.iterrows()):
                    logistics_phone = logistics_row[logistics_phone_select] if logistics_phone_select else ''
                    logistics_address = logistics_row[logistics_address_select] if logistics_address_select else ''
                    # 修复：使用原始索引logistics_idx而不是循环索引i
                    logistics_record_key = (pending_name, str(logistics_phone), logistics_address, logistics_idx)
                    usage_count = logistics_usage.get(logistics_record_key, 0)
                    max_allowed = logistics_max_usage.get(logistics_record_key, 1)
                    product_mapping = logistics_product_mapping.get(logistics_record_key, [])

                    if usage_count < max_allowed:
                        # 使用对应位置的商品信息
                        # 为了解决同名同电话同地址不同商品但物流单号相反的问题
                        # 我们需要更加精确地选择商品信息
                        if product_mapping and len(product_mapping) > 1:
                            # 当有多个商品时，尝试基于待发货明细中的商品进行精确匹配
                            if pending_product:
                                matched_product_info = None
                                best_product_match_score = 0
                                
                                # 遍历所有商品信息，找到最佳匹配
                                for product_info in product_mapping:
                                    if fuzzy_product_match(pending_product, product_info['product']):
                                        # 计算商品匹配分数（基于字符串相似度）
                                        import difflib
                                        similarity = difflib.SequenceMatcher(None, pending_product, product_info['product']).ratio()
                                        if similarity > best_product_match_score:
                                            best_product_match_score = similarity
                                            matched_product_info = product_info
                                
                                # 如果没有找到基于内容的匹配，再考虑使用使用次数
                                if matched_product_info is None and usage_count < len(product_mapping):
                                    matched_product_info = product_mapping[usage_count]
                            else:
                                # 如果没有待发货商品信息，使用使用次数对应的商品信息
                                if usage_count < len(product_mapping):
                                    matched_product_info = product_mapping[usage_count]
                            
                            # 使用匹配到的商品对应的物流信息
                            if matched_product_info:
                                best_match = matched_product_info['logistics_info']
                            else:
                                # 没有找到合适的商品匹配，使用物流行的原始信息
                                best_match = logistics_row.to_dict()
                        else:
                            # 单个商品或没有商品映射的情况
                            if product_mapping and usage_count < len(product_mapping):
                                matched_product_info = product_mapping[usage_count]
                                best_match = matched_product_info['logistics_info']
                            else:
                                # 没有商品映射，直接使用物流信息
                                best_match = logistics_row.to_dict()

                        matched_by = "fallback"
                        match_stats['name_only'] += 1
                        match_stats['total_matched'] += 1

                        # 更新使用计数器
                        logistics_usage[logistics_record_key] = usage_count + 1
                        break

            # 如果仍然没有匹配，标记为无匹配
            if best_match is None:
                match_stats['no_match'] += 1

            # 合并行数据
            merged_row = pending_row.to_dict()
            if best_match is not None:
                for col in [logistics_name_select] + columns_to_add:
                    if pending_phone_select and logistics_phone_select and col == logistics_phone_select:
                        # 特别处理电话列
                        merged_row[col] = best_match[col]
                    elif pending_address_select and logistics_address_select and col == logistics_address_select:
                        # 特别处理地址列
                        merged_row[col] = best_match[col]
                    elif pending_product_select and logistics_product_select and col == logistics_product_select:
                        # 特别处理商品列
                        if matched_product_info and matched_product_info['product']:
                            merged_row[col] = matched_product_info['product']
                        else:
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
        st.write(f"- 姓名唯一匹配: {match_stats['name_only']}")
        st.write(f"- 电话细化匹配: {match_stats['phone_refined']}")
        st.write(f"- 地址细化匹配: {match_stats['address_refined']}")
        st.write(f"- 商品细化匹配: {match_stats['product_refined']}")
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
