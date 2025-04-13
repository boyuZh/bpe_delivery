import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import numpy as np
from datetime import datetime
import re

# 设置页面配置
st.set_page_config(page_title="运输数据看板", page_icon="🚚", layout="wide")

# 定义常量
SHEET_NAMES = ["SYD", "ADL", "BNE"]

# 定义函数：提取派送员信息
def extract_account_info(account_str):
    # 检查是否为有效字符串
    if not isinstance(account_str, str):
        return None, None, str(account_str)  # 处理非字符串值（如NaN）
    
    match = re.match(r"^([A-Z]+(?:-[A-Z]+)*)-(\w+)-(.+)$", account_str)
    if match:
        company_short = match.group(1)
        account_id = match.group(2)
        account_name = match.group(3)
        return company_short, account_id, account_name
    return None, None, account_str  # 仅有姓名时，填充 None

# 使用Streamlit的缓存功能加速数据处理
@st.cache_data(ttl=3600, show_spinner=False)
def process_data(uploaded_files):
    """处理上传的Excel文件，并缓存结果"""
    all_data = {}
    
    if not uploaded_files:
        return all_data
            
    for uploaded_file in uploaded_files:
        for sheet_name in SHEET_NAMES:
            try:
                # 尝试读取每个sheet
                df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
                
                # 跳过空的数据框
                if df.empty:
                    st.warning(f"{sheet_name}工作表为空")
                    continue
                
                # 找到关键列
                supplier_key = None
                signed_time_key = None
                receiver_key = None
                fee_key = None
                weight_key = None
                city_key = None
                tracking_key = None
                vendor_key = None  # 新增商家列
                
                for key in df.keys():
                    if "收派员所属供应商" in key:
                        supplier_key = key
                    if "签收时间" in key:
                        signed_time_key = key
                    if "收派员" in key and "driver" in key.lower() and "供应商" not in key and "编号" not in key.lower():
                        receiver_key = key
                    if "计费" in key:
                        fee_key = key
                    if "重量" in key:
                        weight_key = key
                    if "收件城市" in key:
                        city_key = key
                    if "运单号" in key:
                        tracking_key = key
                    if "商家" in key and "vendor" in key.lower() and "重量" not in key:
                        vendor_key = key
                
                # 确保找到了所需的列
                if not all([signed_time_key, receiver_key, fee_key, weight_key, city_key, tracking_key]):
                    st.warning(f"在 {sheet_name} 工作表中未找到所有必需的列")
                    continue
                
                # 提取文件名中的日期信息
                file_name = uploaded_file.name
                
                # 处理日期格式
                df['日期'] = pd.to_datetime(df[signed_time_key]).dt.date
                df['周数'] = pd.to_datetime(df[signed_time_key]).dt.isocalendar().week.astype('int32')
                
                # 处理重量数据，并添加重量区间分类
                df[weight_key] = pd.to_numeric(df[weight_key], errors='coerce')
                df['重量区间'] = pd.cut(
                    df[weight_key], 
                    bins=[0, 8, 15, float('inf')], 
                    labels=['8kg以下', '8-15kg', '15kg以上'],
                    include_lowest=True
                )
                
                # 添加到数据字典
                if sheet_name not in all_data:
                    all_data[sheet_name] = df
                else:
                    all_data[sheet_name] = pd.concat([all_data[sheet_name], df])
                    
            except Exception as e:
                st.error(f"处理 {sheet_name} 工作表时出错: {str(e)}")
    
    # 提前处理每个城市的数据，以加速后续操作
    cities = list(all_data.keys())  # 创建键的副本，避免在迭代时修改字典
    for city in cities:
        try:
            df = all_data[city]
            if not df.empty:
                # 找到关键列
                try:
                    signed_time_key = [col for col in df.columns if "签收时间" in col][0]
                    receiver_key = [col for col in df.columns if "收派员" in col and "driver" in col.lower() and "供应商" not in col and "编号" not in col.lower()][0]
                    fee_key = [col for col in df.columns if "计费" in col][0]
                    weight_key = [col for col in df.columns if "重量" in col][0]
                    city_key = [col for col in df.columns if "收件城市" in col][0]
                    tracking_key = [col for col in df.columns if "运单号" in col][0]
                    
                    # 尝试找到商家列
                    vendor_cols = [col for col in df.columns if "商家" in col and "vendor" in col.lower() and "重量" not in col]
                    vendor_key = vendor_cols[0] if vendor_cols else None
                except IndexError:
                    st.error(f"在{city}工作表中找不到所需的列，请确保数据格式正确")
                    continue
                
                # 保存关键列名，避免重复查找
                all_data[f"{city}_keys"] = {
                    "signed_time_key": signed_time_key,
                    "receiver_key": receiver_key,
                    "fee_key": fee_key,
                    "weight_key": weight_key,
                    "city_key": city_key,
                    "tracking_key": tracking_key,
                    "vendor_key": vendor_key
                }
                
                try:
                    # 处理商家信息 - 使用商家列(如果存在)，否则使用默认值"未知"
                    if vendor_key:
                        df['商家'] = df[vendor_key].fillna("未知")
                    else:
                        df['商家'] = "未知"
                    
                    # 预计算日维度数据
                    daily_data = df.groupby('日期').agg(
                        运单总量=pd.NamedAgg(column=tracking_key, aggfunc='count'),
                        总收入=pd.NamedAgg(column=fee_key, aggfunc='sum'),
                        活跃派送员数=pd.NamedAgg(column=receiver_key, aggfunc='nunique'),
                        平均包裹重量=pd.NamedAgg(column=weight_key, aggfunc='mean'),
                        小于8kg=pd.NamedAgg(column='重量区间', aggfunc=lambda x: (x == '8kg以下').sum()),
                        介于8_15kg=pd.NamedAgg(column='重量区间', aggfunc=lambda x: (x == '8-15kg').sum()),
                        大于15kg=pd.NamedAgg(column='重量区间', aggfunc=lambda x: (x == '15kg以上').sum())
                    ).reset_index()
                    all_data[f"{city}_daily"] = daily_data
                    
                    # 预计算周维度数据
                    weekly_data = df.groupby('周数').agg(
                        运单总量=pd.NamedAgg(column=tracking_key, aggfunc='count'),
                        总收入=pd.NamedAgg(column=fee_key, aggfunc='sum'),
                        活跃派送员数=pd.NamedAgg(column=receiver_key, aggfunc='nunique'),
                        平均包裹重量=pd.NamedAgg(column=weight_key, aggfunc='mean'),
                        小于8kg=pd.NamedAgg(column='重量区间', aggfunc=lambda x: (x == '8kg以下').sum()),
                        介于8_15kg=pd.NamedAgg(column='重量区间', aggfunc=lambda x: (x == '8-15kg').sum()),
                        大于15kg=pd.NamedAgg(column='重量区间', aggfunc=lambda x: (x == '15kg以上').sum())
                    ).reset_index()
                    all_data[f"{city}_weekly"] = weekly_data
                except Exception as e:
                    st.error(f"处理{city}数据时出错: {str(e)}")
                    continue
        except Exception as e:
            st.error(f"处理{city}时发生意外错误: {str(e)}")
            continue
    
    return all_data

# 使用缓存加速图表生成
@st.cache_data(ttl=3600)
def generate_line_chart(data, x, y, title):
    fig = px.line(data, x=x, y=y, title=title, markers=True)
    return fig

@st.cache_data(ttl=3600)
def generate_bar_chart(data, x, y, title):
    fig = px.bar(data, x=x, y=y, title=title, text_auto=True)
    return fig

@st.cache_data(ttl=3600)
def generate_pie_chart(data, values, names, title):
    fig = px.pie(data, values=values, names=names, title=title)
    return fig

# 构建看板函数
def create_dashboard(data):
    """创建交互式看板"""
    if not data:
        st.warning("请上传数据文件")
        return
    
    # 获取实际城市列表（过滤掉辅助键）
    cities = [city for city in data.keys() if city in SHEET_NAMES]
    
    if not cities:
        st.warning("未找到有效的城市数据")
        return
    
    # 城市选择器
    selected_city = st.sidebar.selectbox("选择城市", cities)
    
    # 创建选项卡
    tab1, tab2 = st.tabs(["按日维度", "按周维度"])
    
    # 获取所选城市的数据
    df = data[selected_city]
    
    # 获取该城市的预计算键
    keys = data[f"{selected_city}_keys"]
    signed_time_key = keys["signed_time_key"]
    receiver_key = keys["receiver_key"]
    fee_key = keys["fee_key"]
    weight_key = keys["weight_key"] 
    city_key = keys["city_key"]
    tracking_key = keys["tracking_key"]
    
    # 获取日期范围
    date_min = df['日期'].min()
    date_max = df['日期'].max()
    
    # 获取预计算的数据
    daily_data = data[f"{selected_city}_daily"]
    weekly_data = data[f"{selected_city}_weekly"]
    
    # 根据维度选择处理数据
    with tab1:  # 按日维度
        st.header(f"{selected_city} - 按日维度分析")
        
        # 日期选择器
        selected_dates = st.slider(
            "选择日期范围",
            min_value=date_min,
            max_value=date_max,
            value=(date_min, date_max)
        )
        
        # 过滤日期范围
        filtered_daily_data = daily_data[(daily_data['日期'] >= selected_dates[0]) & (daily_data['日期'] <= selected_dates[1])]
        filtered_df = df[(df['日期'] >= selected_dates[0]) & (df['日期'] <= selected_dates[1])]
        
        # 1. 运单总量和总收入
        col1, col2 = st.columns(2)
        
        with col1:
            fig1 = generate_line_chart(filtered_daily_data, '日期', '运单总量', "运单总量走势")
            st.plotly_chart(fig1, use_container_width=True)
            
        with col2:
            fig2 = generate_line_chart(filtered_daily_data, '日期', '总收入', "总收入走势")
            st.plotly_chart(fig2, use_container_width=True)
        
        # 2. 活跃派送员数量和包裹平均重量
        col3, col4 = st.columns(2)
        
        with col3:
            fig3 = generate_line_chart(filtered_daily_data, '日期', '活跃派送员数', "活跃派送员数量走势")
            st.plotly_chart(fig3, use_container_width=True)
            
        with col4:
            fig4 = generate_line_chart(filtered_daily_data, '日期', '平均包裹重量', "包裹平均重量走势")
            st.plotly_chart(fig4, use_container_width=True)
        
        # 5. 包裹重量分布
        st.subheader("包裹重量分布")
        
        # 准备包裹重量区间数据
        weight_data = pd.DataFrame({
            '重量区间': ['8kg以下', '8-15kg', '15kg以上'],
            '包裹数量': [
                filtered_daily_data['小于8kg'].sum(),
                filtered_daily_data['介于8_15kg'].sum(),
                filtered_daily_data['大于15kg'].sum()
            ]
        })
        
        col5, col6 = st.columns(2)
        
        with col5:
            # 创建饼图
            fig5 = px.pie(
                weight_data, 
                values='包裹数量', 
                names='重量区间', 
                title="包裹重量区间分布"
            )
            st.plotly_chart(fig5, use_container_width=True)
            
        with col6:
            # 创建柱状图
            fig6 = px.bar(
                weight_data, 
                x='重量区间', 
                y='包裹数量', 
                title="包裹重量区间分布", 
                text_auto=True
            )
            st.plotly_chart(fig6, use_container_width=True)
            
        # 展示详细数据
        with st.expander("查看重量区间详细数据"):
            weight_detail = pd.DataFrame({
                '日期': filtered_daily_data['日期'],
                '8kg以下': filtered_daily_data['小于8kg'],
                '8-15kg': filtered_daily_data['介于8_15kg'],
                '15kg以上': filtered_daily_data['大于15kg']
            })
            st.dataframe(weight_detail)
        
        # 3. 收件城市比例
        city_counts = filtered_df[city_key].value_counts().reset_index()
        city_counts.columns = ['城市', '数量']
        
        fig7 = generate_pie_chart(city_counts, '数量', '城市', "收件城市比例")
        st.plotly_chart(fig7, use_container_width=True)
        
        # 4. 商家比例
        supplier_counts = filtered_df['商家'].value_counts().reset_index()
        supplier_counts.columns = ['商家', '数量']
        
        fig8 = generate_pie_chart(supplier_counts, '数量', '商家', "商家比例")
        st.plotly_chart(fig8, use_container_width=True)
    
    with tab2:  # 按周维度
        st.header(f"{selected_city} - 按周维度分析")
        
        # 获取实际有数据的周列表
        available_weeks = []
        for week in sorted(df['周数'].unique()):
            week_data = df[df['周数'] == week]
            if len(week_data) > 0:  # 只添加有数据的周
                available_weeks.append(week)
        
        # 显示可用周数信息
        if not available_weeks:
            st.warning(f"未找到任何有效的周数据")
            return
            
        # 周选择器
        selected_weeks = st.multiselect("选择周数", available_weeks, default=available_weeks)
        
        if not selected_weeks:
            st.warning("请至少选择一个周")
            return
            
        # 过滤周数
        filtered_weekly_data = weekly_data[weekly_data['周数'].isin(selected_weeks)]
        filtered_df = df[df['周数'].isin(selected_weeks)]
        
        # 显示数据有效性信息
        st.info(f"数据包含的周数: {', '.join([str(w) for w in available_weeks])}")
        
        # 检查数据是否为空
        if filtered_weekly_data.empty:
            st.warning("所选周数没有数据")
            return
        
        # 1. 运单总量和总收入
        col1, col2 = st.columns(2)
        
        with col1:
            fig1 = generate_bar_chart(filtered_weekly_data, '周数', '运单总量', "运单总量走势")
            st.plotly_chart(fig1, use_container_width=True)
            
        with col2:
            fig2 = generate_bar_chart(filtered_weekly_data, '周数', '总收入', "总收入走势")
            st.plotly_chart(fig2, use_container_width=True)
        
        # 2. 活跃派送员数量和包裹平均重量
        col3, col4 = st.columns(2)
        
        with col3:
            fig3 = generate_bar_chart(filtered_weekly_data, '周数', '活跃派送员数', "活跃派送员数量走势")
            st.plotly_chart(fig3, use_container_width=True)
            
        with col4:
            fig4 = generate_bar_chart(filtered_weekly_data, '周数', '平均包裹重量', "包裹平均重量走势")
            st.plotly_chart(fig4, use_container_width=True)
        
        # 5. 包裹重量分布
        st.subheader("包裹重量分布")
        
        # 周维度下准备包裹重量区间数据
        weight_data = pd.DataFrame({
            '重量区间': ['8kg以下', '8-15kg', '15kg以上'],
            '包裹数量': [
                filtered_weekly_data['小于8kg'].sum(),
                filtered_weekly_data['介于8_15kg'].sum(),
                filtered_weekly_data['大于15kg'].sum()
            ]
        })
        
        col5, col6 = st.columns(2)
        
        with col5:
            # 创建饼图
            fig5 = px.pie(
                weight_data, 
                values='包裹数量', 
                names='重量区间', 
                title="包裹重量区间分布"
            )
            st.plotly_chart(fig5, use_container_width=True)
            
        with col6:
            # 创建柱状图 - 按周展示不同重量区间的包裹数
            # 准备按周的数据
            weeks_weight_data = pd.melt(
                filtered_weekly_data,
                id_vars=['周数'],
                value_vars=['小于8kg', '介于8_15kg', '大于15kg'],
                var_name='重量区间',
                value_name='包裹数量'
            )
            # 替换变量名为更友好的显示
            weeks_weight_data['重量区间'] = weeks_weight_data['重量区间'].replace({
                '小于8kg': '8kg以下',
                '介于8_15kg': '8-15kg',
                '大于15kg': '15kg以上'
            })
            
            fig6 = px.bar(
                weeks_weight_data,
                x='周数',
                y='包裹数量',
                color='重量区间',
                barmode='group',
                title="各周包裹重量区间分布"
            )
            st.plotly_chart(fig6, use_container_width=True)
            
        # 展示详细数据
        with st.expander("查看重量区间详细数据"):
            weight_detail = pd.DataFrame({
                '周数': filtered_weekly_data['周数'],
                '8kg以下': filtered_weekly_data['小于8kg'],
                '8-15kg': filtered_weekly_data['介于8_15kg'],
                '15kg以上': filtered_weekly_data['大于15kg'],
                '包裹总量': filtered_weekly_data['运单总量']
            })
            st.dataframe(weight_detail)
        
        # 3. 收件城市比例
        city_counts = filtered_df[city_key].value_counts().reset_index()
        city_counts.columns = ['城市', '数量']
        
        fig7 = generate_pie_chart(city_counts, '数量', '城市', "收件城市比例")
        st.plotly_chart(fig7, use_container_width=True)
        
        # 4. 商家比例
        supplier_counts = filtered_df['商家'].value_counts().reset_index()
        supplier_counts.columns = ['商家', '数量']
        
        fig8 = generate_pie_chart(supplier_counts, '数量', '商家', "商家比例")
        st.plotly_chart(fig8, use_container_width=True)

# 主应用
def main():
    st.title("🚚 运输数据分析看板")
    st.markdown("分析SYD、ADL、BNE三个城市的运输数据")
    
    # 增加进度信息展示
    with st.spinner("准备应用..."):
        # 上传文件
        uploaded_files = st.file_uploader("上传XLSX数据文件", accept_multiple_files=True, type=['xlsx'])
        
        if uploaded_files:
            # 处理数据
            with st.spinner("处理数据中，请耐心等待..."):
                data = process_data(uploaded_files)
            
            # 创建看板
            create_dashboard(data)
        else:
            st.info("请上传包含SYD、ADL、BNE工作表的Excel文件。注意：一个文件对应一周的数据，多周数据需要上传多个文件。")

if __name__ == "__main__":
    main() 