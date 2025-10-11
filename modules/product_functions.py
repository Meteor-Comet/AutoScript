import pandas as pd
import streamlit as st
import re
import difflib

def standardize_product_names(df, product_list):
    """
    标准化商品名称
    """
    # 确保所有列都是字符串类型
    for col in df.columns:
        df[col] = df[col].astype(str)
    
    # 选择产品名称列 - 添加对"货品名称"列的支持
    if '货品名称' in df.columns:
        product_col = '货品名称'
    elif '产品名称' in df.columns:
        product_col = '产品名称'
    else:
        product_col = st.selectbox("选择包含产品名称的列", df.columns)
    
    # 创建标准化映射
    standardized_mapping = {}
    
    # 定义特定的映射规则
    special_mappings = {
        "硬盒纸抽（130抽）（250506）": "黄金葉  盒装抽纸  （计价单位：盒） 国产定制",
        "湿巾（10片/包）（250506）": "黄金葉 湿纸巾 10片/包 （计价单位：包） 国产定制"
    }
    
    # 定义关键词映射，用于提取商品的唯一关键词进行匹配
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
    
    # 获取当前产品名称的唯一值
    unique_products = df[product_col].unique()
    
    # 对每个不规范的产品名称进行匹配
    for product in unique_products:
        # 首先检查特殊映射规则
        if product in special_mappings:
            standardized_mapping[product] = special_mappings[product]
            continue
            
        # 然后检查是否能通过关键词匹配
        matched = False
        for keyword, standard_name in keyword_mappings.items():
            # 如果产品名称包含关键词，则匹配
            if keyword in product:
                standardized_mapping[product] = standard_name
                matched = True
                break
        
        # 如果通过关键词无法匹配，则使用模糊匹配
        if not matched:
            matches = difflib.get_close_matches(product, product_list, n=1, cutoff=0.6)
            if matches:
                standardized_mapping[product] = matches[0]
            else:
                # 否则保持原名称
                standardized_mapping[product] = product
    
    # 应用标准化映射
    df['标准化产品名称'] = df[product_col].map(standardized_mapping)
    
    return df, standardized_mapping

def convert_product_quantities(df, conversion_direction, selected_products, conversion_rules):
    """
    转换商品数量（箱<->个）
    """
    # 选择必要的列
    if '货品名称' in df.columns:
        product_col = '货品名称'
    elif '产品名称' in df.columns:
        product_col = '产品名称'
    else:
        product_col = st.selectbox("选择包含产品名称的列", df.columns)

    if '数量' in df.columns:
        quantity_col = '数量'
    else:
        quantity_col = st.selectbox("选择包含数量的列", df.columns)

    if '规格' in df.columns:
        unit_col = '规格'
    else:
        unit_col = st.selectbox("选择包含规格的列", df.columns)
    
    # 根据转换方向添加新列
    if conversion_direction == "箱 → 个":
        new_column_name = '数量（个）'
        new_unit_name = '个'
    else:  # 个 → 箱
        new_column_name = '数量（箱）'
        new_unit_name = '箱'
    
    # 找到数量列的位置
    quantity_col_index = list(df.columns).index(quantity_col)
    
    # 创建新列数据
    new_column_data = pd.to_numeric(df[quantity_col], errors='coerce').fillna(0)
    new_unit_data = df[unit_col].copy()
    
    # 将新列插入到数量列后面
    columns_list = list(df.columns)
    columns_list.insert(quantity_col_index + 1, new_column_name)
    columns_list.insert(quantity_col_index + 2, '规格（转换后）')
    
    # 重新排列DataFrame的列
    result_df = df.reindex(columns=columns_list)
    
    # 填充新列数据
    result_df[new_column_name] = new_column_data
    result_df['规格（转换后）'] = new_unit_data
    
    # 执行转换
    converted_count = 0
    for index, row in result_df.iterrows():
        product_name = str(row[product_col])

        # 只转换选中的商品
        if product_name in selected_products:
            # 检查是否需要转换（根据转换方向和规格判断）
            need_convert = False
            if conversion_direction == "箱 → 个":
                # 只有规格为"箱"的才需要转换
                if '箱' in str(row[unit_col]) and '个' not in str(row[unit_col]):
                    need_convert = True
            elif conversion_direction == "个 → 箱":
                # 只有规格为"个"、"台"、"条"、"包"等的才需要转换
                unit_str = str(row[unit_col])
                if ('个' in unit_str or '台' in unit_str or '条' in unit_str or '包' in unit_str) and '箱' not in unit_str:
                    need_convert = True
            
            if need_convert:
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
                    try:
                        original_quantity = float(row[quantity_col])
                        if conversion_direction == "箱 → 个":
                            converted_quantity = original_quantity * multiplier
                        else:  # 个 → 箱
                            converted_quantity = original_quantity / multiplier
                        
                        result_df.at[index, new_column_name] = converted_quantity
                        result_df.at[index, '规格（转换后）'] = new_unit_name
                        converted_count += 1
                    except (ValueError, TypeError):
                        pass
    
    return result_df, converted_count

def convert_product_quantities_manual(df, selected_products, conversion_config):
    """
    手动转换商品数量（支持多种单位之间的转换）
    """
    # 选择必要的列
    if '货品名称' in df.columns:
        product_col = '货品名称'
    elif '产品名称' in df.columns:
        product_col = '产品名称'
    else:
        product_col = "产品名称"  # 默认值

    if '数量' in df.columns:
        quantity_col = '数量'
    else:
        quantity_col = df.columns[1]  # 默认使用第二列作为数量列

    if '规格' in df.columns:
        unit_col = '规格'
    else:
        unit_col = "规格"  # 默认值
    
    source_unit = conversion_config['source_unit']
    target_unit = conversion_config['target_unit']
    conversion_rules = conversion_config['conversion_rules']
    
    # 根据目标单位添加新列
    new_column_name = f'数量（{target_unit}）'
    new_unit_name = target_unit
    
    # 找到数量列的位置
    quantity_col_index = list(df.columns).index(quantity_col)
    
    # 创建新列数据
    new_column_data = pd.to_numeric(df[quantity_col], errors='coerce').fillna(0)
    new_unit_data = df[unit_col].copy()
    
    # 将新列插入到数量列后面
    columns_list = list(df.columns)
    columns_list.insert(quantity_col_index + 1, new_column_name)
    columns_list.insert(quantity_col_index + 2, '规格（转换后）')
    
    # 重新排列DataFrame的列
    result_df = df.reindex(columns=columns_list)
    
    # 填充新列数据
    result_df[new_column_name] = new_column_data
    result_df['规格（转换后）'] = new_unit_data
    
    # 执行转换
    converted_count = 0
    for index, row in result_df.iterrows():
        product_name = str(row[product_col])

        # 只转换选中的商品
        if product_name in selected_products:
            # 检查是否具有指定的源单位
            unit_str = str(row[unit_col])
            if source_unit in unit_str and target_unit not in unit_str:
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
                    try:
                        original_quantity = float(row[quantity_col])
                        # 判断转换方向
                        if target_unit == "个":  # 箱/件等 → 个
                            converted_quantity = original_quantity * multiplier
                        else:  # 个 → 箱/件等
                            converted_quantity = original_quantity / multiplier
                        
                        result_df.at[index, new_column_name] = converted_quantity
                        result_df.at[index, '规格（转换后）'] = new_unit_name
                        converted_count += 1
                    except (ValueError, TypeError):
                        pass
    
    return result_df, converted_count
