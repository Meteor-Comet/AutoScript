import pandas as pd
import streamlit as st

def find_header_row(df):
    """
    查找真正的表头行
    
    策略:
    1. 检查前10行，寻找包含关键字的行
    2. 如果找不到，返回None（使用默认列名）
    """
    # 关键字列表，用于识别表头行
    header_keywords = [
        '姓名', '名称', '联系人', '收货人', '客户',  # 姓名相关
        '电话', '手机', '联系方式', '联系电话', '号码',  # 电话相关
        '地址', '收货地址', '详细地址', '邮寄地址'  # 地址相关
    ]
    
    # 检查前10行（或者表格的实际行数，取较小值）
    max_row = min(10, len(df))
    
    for i in range(max_row):
        row = df.iloc[i]
        # 将行转换为字符串列表
        row_values = [str(x).lower() for x in row.values if pd.notna(x)]
        row_text = ' '.join(row_values)
        
        # 计算这一行包含多少个关键字
        keyword_count = sum(1 for keyword in header_keywords if keyword in row_text)
        
        # 如果包含至少2个关键字，认为这是表头行
        if keyword_count >= 2:
            return i
    
    return None

def identify_key_columns(df):
    """
    识别表格中的关键列
    
    返回:
    字典，键为目标字段名，值为可能包含该信息的列名列表（按优先级排序）
    """
    # 初始化结果字典
    column_mapping = {
        '客户名称': [],
        '收货地址': [],
        '收货人': [],
        '手机': []
    }
    
    # 关键字匹配规则
    patterns = {
        '客户名称': ['客户', '名称', '单位', '公司', '商户', '店铺', '药店', '药房'],
        '收货地址': ['地址', '收货地址', '详细地址', '邮寄地址', '收件地址', '位置'],
        '收货人': ['姓名', '收货人', '联系人', '收件人', '联系人姓名', '收货方'],
        '手机': ['电话', '手机', '联系方式', '联系电话', '号码', '联系号码', '电话号码', '手机号']
    }
    
    # 检查每一列
    for col in df.columns:
        col_str = str(col).lower()
        
        # 对每个目标字段
        for target_field, keywords in patterns.items():
            # 检查列名是否包含关键字
            for keyword in keywords:
                if keyword in col_str:
                    column_mapping[target_field].append(col)
                    break
    
    # 如果没有找到某些字段，尝试通过数据模式识别
    if not any(column_mapping.values()):
        column_mapping = identify_columns_by_pattern(df)
    
    return column_mapping

def identify_columns_by_pattern(df):
    """
    通过数据模式识别关键列
    """
    # 初始化结果字典
    column_mapping = {
        '客户名称': [],
        '收货地址': [],
        '收货人': [],
        '手机': []
    }
    
    # 检查每一列的数据模式
    for col in df.columns:
        # 跳过全空列
        if df[col].isna().all():
            continue
        
        # 获取非空值
        values = [str(x) for x in df[col].dropna().tolist() if str(x).strip()]
        if not values:
            continue
        
        # 样本大小
        sample_size = min(10, len(values))
        sample = values[:sample_size]
        
        # 检查是否是电话号码列
        if all(is_phone_number(x) for x in sample):
            column_mapping['手机'].append(col)
            continue
        
        # 检查是否是地址列
        if all(len(x) > 10 and ('路' in x or '街' in x or '区' in x or '市' in x or '省' in x) for x in sample):
            column_mapping['收货地址'].append(col)
            continue
        
        # 检查是否是姓名列
        if all(2 <= len(x) <= 5 and not any(c.isdigit() for c in x) for x in sample):
            column_mapping['收货人'].append(col)
            continue
        
        # 其他可能是客户名称的列
        if all(len(x) > 2 for x in sample):
            column_mapping['客户名称'].append(col)
    
    return column_mapping

def is_phone_number(text):
    """
    判断文本是否是电话号码
    """
    # 移除所有非数字字符
    digits_only = ''.join(c for c in text if c.isdigit())
    # 检查数字长度是否在合理范围内
    return 7 <= len(digits_only) <= 13

def smart_column_detection(df, filename):
    """
    智能识别表格中的姓名、联系方式、收货地址等信息，不依赖固定的列名和行号
    """
    try:
        # 打印调试信息
        st.write(f"开始智能识别文件: {filename}")
        st.write(f"原始列名: {df.columns.tolist()}")
        
        # 第一步：找到真正的表头行
        header_row_idx = find_header_row(df)
        
        # 如果找到了表头行，重新设置列名
        if header_row_idx is not None and header_row_idx > 0:
            st.write(f"找到表头行: {header_row_idx}")
            # 保存原始列名以备后用
            original_columns = df.columns.tolist()
            # 设置新的列名
            df.columns = df.iloc[header_row_idx].values
            # 只保留表头行之后的数据
            df = df.iloc[header_row_idx+1:].reset_index(drop=True)
            st.write(f"重设列名后: {df.columns.tolist()}")
        
        # 第二步：智能匹配关键列
        column_mapping = identify_key_columns(df)
        st.write(f"列映射结果: {column_mapping}")
        
        # 第三步：提取数据
        extracted_data = []
        
        # 显示前几行数据以便调试
        st.write("数据预览:")
        st.dataframe(df.head(3))
        
        for index, row in df.iterrows():
            try:
                # 跳过标题行或空行
                if index == 0:
                    # 检查是否是标题行
                    is_header = any(
                        keyword in str(value).lower() 
                        for value in row.values 
                        for keyword in ['姓名', '电话', '地址', '客户', '名称']
                    )
                    if is_header:
                        continue
                
                # 跳过空行
                if row.isnull().all():
                    continue
                
                # 创建一个新的数据项
                item = {
                    '收货人': '',
                    '手机': '',
                    '收货地址': '',
                    '客户备注': '',  # 店名放在这里
                    '源文件': filename
                }
                
                # 根据识别的列映射填充数据
                for target_field, source_columns in column_mapping.items():
                    for col in source_columns:
                        if col in df.columns and pd.notna(row[col]) and str(row[col]).strip():
                            if target_field == '客户名称':
                                # 将客户名称（店名）直接放入客户备注
                                item['客户备注'] = str(row[col]).strip()
                            elif target_field == '收货人':
                                # 确保收货人是实际人名，不是地区名
                                if '酒泉' not in str(row[col]) and '兰州' not in str(row[col]):
                                    item['收货人'] = str(row[col]).strip()
                            else:
                                item[target_field] = str(row[col]).strip()
                            break
                
                # 只有当至少有收货人或电话不为空时才添加数据
                if item['收货人'] or item['手机']:
                    extracted_data.append(item)
                    
            except Exception as e:
                st.warning(f"处理第 {index+1} 行时出错: {e}")
        
        # 显示提取结果
        if extracted_data:
            st.write(f"成功提取 {len(extracted_data)} 条数据")
            st.write("提取结果示例:")
            st.dataframe(pd.DataFrame(extracted_data[:3]))
        else:
            st.warning("未能提取到任何数据")
        
        return extracted_data
        
    except Exception as e:
        st.error(f"智能处理文件 {filename} 时出错: {str(e)}")
        import traceback
        st.error(traceback.format_exc())
        return []