import pandas as pd
import streamlit as st

def enhanced_vlookup(main_df, reference_df, match_cols_main, match_cols_ref, columns_to_add, join_type, handle_duplicates):
    """
    增强版VLOOKUP功能
    """
    try:
        # 确保所有列都是字符串类型，避免PyArrow错误
        for col in main_df.columns:
            main_df[col] = main_df[col].astype(str)
            
        for col in reference_df.columns:
            reference_df[col] = reference_df[col].astype(str)
        
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
        
        return result_df
    except Exception as e:
        st.error(f"执行过程中出现错误: {str(e)}")
        import traceback
        st.error(traceback.format_exc())
        return None