import streamlit as st
from datetime import datetime
import pandas as pd

def setup_page():
    """
    设置页面配置
    """
    st.set_page_config(
        page_title="数据处理系统",
        page_icon="🧮",
        layout="wide"
    )

    # 页面标题
    st.title("数据处理系统")
    st.markdown("---")

def setup_sidebar():
    """
    设置侧边栏
    """
    # 侧边栏
    st.sidebar.header("操作选项")
    app_mode = st.sidebar.selectbox(
        "选择功能",
        ["批量处理发放明细", "核对发放明细与供应商订单", "导入明细到导单模板",
         "增强版VLOOKUP", "物流单号匹配", "表合并", "供应商订单分析", "商品名称标准化", "商品数量转换"]
    )
    return app_mode

def show_footer():
    """
    显示页脚
    """
    st.markdown("---")
    st.markdown("© 2025 数据处理系统 v1.0")

def download_button(df, filename_prefix):
    """
    创建下载按钮
    """
    output_file = f"{filename_prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    df.to_excel(output_file, index=False)
    
    with open(output_file, "rb") as file:
        st.download_button(
            label="下载结果文件",
            data=file,
            file_name=output_file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )