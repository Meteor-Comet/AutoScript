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
    
    # 常见地区名称列表
    regional_names = [
        "安阳", "许昌", "北京", "上海", "天津", "重庆", "石家庄", "唐山", "秦皇岛", "邯郸", "邢台", "保定", 
        "张家口", "承德", "沧州", "廊坊", "衡水", "太原", "大同", "阳泉", "长治", "晋城", "朔州", "晋中", 
        "运城", "忻州", "临汾", "吕梁", "呼和浩特", "包头", "乌海", "赤峰", "通辽", "鄂尔多斯", "呼伦贝尔", 
        "巴彦淖尔", "乌兰察布", "沈阳", "大连", "鞍山", "抚顺", "本溪", "丹东", "锦州", "营口", "阜新", 
        "辽阳", "盘锦", "铁岭", "朝阳", "葫芦岛", "长春", "吉林", "四平", "辽源", "通化", "白山", "松原", 
        "白城", "延边", "哈尔滨", "齐齐哈尔", "鸡西", "鹤岗", "双鸭山", "大庆", "伊春", "佳木斯", "七台河", 
        "牡丹江", "黑河", "绥化", "南京", "无锡", "徐州", "常州", "苏州", "南通", "连云港", "淮安", "盐城", 
        "扬州", "镇江", "泰州", "宿迁", "杭州", "宁波", "温州", "嘉兴", "湖州", "绍兴", "金华", "衢州", 
        "舟山", "台州", "丽水", "合肥", "芜湖", "蚌埠", "淮南", "马鞍山", "淮北", "铜陵", "安庆", "黄山", 
        "滁州", "阜阳", "宿州", "六安", "亳州", "池州", "宣城", "福州", "厦门", "莆田", "三明", "泉州", 
        "漳州", "南平", "龙岩", "宁德", "南昌", "景德镇", "萍乡", "九江", "新余", "鹰潭", "赣州", "吉安", 
        "宜春", "抚州", "上饶", "济南", "青岛", "淄博", "枣庄", "东营", "烟台", "潍坊", "济宁", "泰安", 
        "威海", "日照", "临沂", "德州", "聊城", "滨州", "菏泽", "郑州", "开封", "洛阳", "平顶山", "安阳", 
        "鹤壁", "新乡", "焦作", "濮阳", "许昌", "漯河", "三门峡", "南阳", "商丘", "信阳", "周口", "驻马店", 
        "武汉", "黄石", "十堰", "宜昌", "襄阳", "鄂州", "荆门", "孝感", "荆州", "黄冈", "咸宁", "随州", 
        "恩施", "长沙", "株洲", "湘潭", "衡阳", "邵阳", "岳阳", "常德", "张家界", "益阳", "郴州", "永州", 
        "怀化", "娄底", "湘西", "广州", "韶关", "深圳", "珠海", "汕头", "佛山", "江门", "湛江", "茂名", 
        "肇庆", "惠州", "梅州", "汕尾", "河源", "阳江", "清远", "东莞", "中山", "潮州", "揭阳", "云浮", 
        "南宁", "柳州", "桂林", "梧州", "北海", "防城港", "钦州", "贵港", "玉林", "百色", "贺州", "河池", 
        "来宾", "崇左", "海口", "三亚", "三沙", "成都", "自贡", "攀枝花", "泸州", "德阳", "绵阳", "广元", 
        "遂宁", "内江", "乐山", "南充", "眉山", "宜宾", "广安", "达州", "雅安", "巴中", "资阳", "阿坝", 
        "甘孜", "凉山", "贵阳", "六盘水", "遵义", "安顺", "毕节", "铜仁", "昆明", "曲靖", "玉溪", "保山", 
        "昭通", "丽江", "普洱", "临沧", "拉萨", "西安", "铜川", "宝鸡", "咸阳", "渭南", "延安", "汉中", 
        "榆林", "安康", "商洛", "兰州", "嘉峪关", "金昌", "白银", "天水", "武威", "张掖", "平凉", "酒泉", 
        "庆阳", "定西", "陇南", "西宁", "银川", "乌鲁木齐", "克拉玛依", "文旅", "地方", "城市", "区域"
    ]
    
    # 获取当前产品名称的唯一值
    unique_products = df[product_col].unique()
    
    # 对每个不规范的产品名称进行匹配
    for product in unique_products:
        # 首先检查商品是否已经在标准商品列表中
        if product in product_list:
            # 如果已经在标准列表中，直接保留原名称
            standardized_mapping[product] = product
            continue
            
        # 然后检查是否包含地区名称
        contains_region = any(region in product for region in regional_names)
        if contains_region:
            # 保持包含地区名称的商品的原始名称
            standardized_mapping[product] = product
            continue
            
        # 然后检查特殊映射规则
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
