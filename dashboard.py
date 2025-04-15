import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
from datetime import datetime
import re
import os
# 导入我们的工具函数
from utils import list_supabase_files, download_supabase_file



# 设置页面配置
st.set_page_config(page_title="运输数据看板", page_icon="🚚", layout="wide")

# 定义哈希函数，用于验证密码
def check_password(password):
    """检查密码是否正确"""
    
    # 尝试从st.secrets获取密码哈希
    if "login" in st.secrets:
        return password == st.secrets["login"]["password"]
    
    return False

# 验证会话功能
def verify_password():
    """确保用户已通过密码验证"""
    # 检查会话状态
    if "password_verified" not in st.session_state:
        st.session_state.password_verified = False
    
    # 如果尚未验证，显示密码输入框
    if not st.session_state.password_verified:
        st.markdown("## 🔒 需要密码访问")
        
        password = st.text_input("请输入密码查看数据", type="password")
        
        if st.button("登录"):
            if check_password(password):
                st.session_state.password_verified = True
                st.experimental_rerun()
            else:
                st.error("❌ 密码错误")
        
        st.stop()  # 停止执行后续代码

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
def process_data(_uploaded_files):
    """处理上传的Excel文件或本地文件对象，并缓存结果"""
    all_data = {}
    
    if not _uploaded_files:
        return all_data
            
    for uploaded_file in _uploaded_files:
        # 处理从Supabase下载的文件
        file_path = getattr(uploaded_file, 'path', None)
        file_name = getattr(uploaded_file, 'name', None)
        
        for sheet_name in SHEET_NAMES:
            try:
                # 根据文件类型读取数据
                if file_path and os.path.exists(file_path):
                    # 如果是本地文件路径（从Supabase下载）
                    df = pd.read_excel(file_path, sheet_name=sheet_name)
                else:
                    # 如果是上传的文件对象
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
                # 使用之前获取的file_name变量
                if file_name is None and hasattr(uploaded_file, 'name'):
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
    
    # 创建汇总数据
    if cities:
        try:
            # 合并所有城市的数据
            combined_df = pd.concat([all_data[city] for city in cities if city in all_data])
            
            if not combined_df.empty:
                # 找到共同的关键列
                common_columns = {}
                for city in cities:
                    if f"{city}_keys" in all_data:
                        city_keys = all_data[f"{city}_keys"]
                        for key_type, column in city_keys.items():
                            if key_type not in common_columns:
                                common_columns[key_type] = []
                            common_columns[key_type].append(column)
                
                # 使用第一个城市的列名作为汇总数据的列名
                if cities and f"{cities[0]}_keys" in all_data:
                    first_city_keys = all_data[f"{cities[0]}_keys"]
                    all_data["ALL_keys"] = first_city_keys.copy()
                
                # 预计算日维度汇总数据
                if "ALL_keys" in all_data:
                    keys = all_data["ALL_keys"]
                    signed_time_key = keys["signed_time_key"]
                    receiver_key = keys["receiver_key"]
                    fee_key = keys["fee_key"]
                    weight_key = keys["weight_key"]
                    tracking_key = keys["tracking_key"]
                    
                    # 处理商家信息
                    if 'vendor_key' in keys and keys['vendor_key']:
                        combined_df['商家'] = combined_df[keys['vendor_key']].fillna("未知")
                    else:
                        combined_df['商家'] = "未知"
                    
                    # 日维度汇总数据
                    daily_data = combined_df.groupby('日期').agg(
                        运单总量=pd.NamedAgg(column=tracking_key, aggfunc='count'),
                        总收入=pd.NamedAgg(column=fee_key, aggfunc='sum'),
                        活跃派送员数=pd.NamedAgg(column=receiver_key, aggfunc='nunique'),
                        平均包裹重量=pd.NamedAgg(column=weight_key, aggfunc='mean'),
                        小于8kg=pd.NamedAgg(column='重量区间', aggfunc=lambda x: (x == '8kg以下').sum()),
                        介于8_15kg=pd.NamedAgg(column='重量区间', aggfunc=lambda x: (x == '8-15kg').sum()),
                        大于15kg=pd.NamedAgg(column='重量区间', aggfunc=lambda x: (x == '15kg以上').sum())
                    ).reset_index()
                    all_data["ALL_daily"] = daily_data
                    
                    # 周维度汇总数据
                    weekly_data = combined_df.groupby('周数').agg(
                        运单总量=pd.NamedAgg(column=tracking_key, aggfunc='count'),
                        总收入=pd.NamedAgg(column=fee_key, aggfunc='sum'),
                        活跃派送员数=pd.NamedAgg(column=receiver_key, aggfunc='nunique'),
                        平均包裹重量=pd.NamedAgg(column=weight_key, aggfunc='mean'),
                        小于8kg=pd.NamedAgg(column='重量区间', aggfunc=lambda x: (x == '8kg以下').sum()),
                        介于8_15kg=pd.NamedAgg(column='重量区间', aggfunc=lambda x: (x == '8-15kg').sum()),
                        大于15kg=pd.NamedAgg(column='重量区间', aggfunc=lambda x: (x == '15kg以上').sum())
                    ).reset_index()
                    all_data["ALL_weekly"] = weekly_data
                    
                    # 添加合并的数据帧
                    all_data["ALL"] = combined_df
        except Exception as e:
            st.error(f"汇总数据时出错: {str(e)}")
    
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
    
    # 添加汇总选项
    all_options = ["全部城市"] + cities
    
    # 城市选择器，默认选择"全部城市"
    selected_option = st.sidebar.selectbox("选择城市", all_options, index=0)
    
    # 创建选项卡
    tab1, tab2 = st.tabs(["按日维度", "按周维度"])
    
    # 根据所选选项确定数据源
    if selected_option == "全部城市":
        selected_city = "ALL"
        st.sidebar.info("当前显示所有城市的汇总数据")
    else:
        selected_city = selected_option
    
    # 如果选择了汇总但没有汇总数据
    if selected_city == "ALL" and "ALL" not in data:
        st.warning("未找到汇总数据，请确保有多个城市的数据可用")
        return
    
    # 获取所选城市/汇总的数据
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
        if selected_city == "ALL":
            st.header(f"所有城市汇总 - 按日维度分析")
            
            # 添加重要统计数据的摘要
            st.subheader("核心统计指标")
            
            # 计算总体指标
            total_deliveries = len(df)
            total_revenue = df[fee_key].sum()
            total_couriers = df[receiver_key].nunique()
            avg_weight = df[weight_key].mean()
            
            # 展示指标卡片
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("总运单量", f"{total_deliveries:,}")
            
            with col2:
                st.metric("总收入", f"${total_revenue:,.2f}")
            
            with col3:
                st.metric("活跃派送员数", f"{total_couriers}")
            
            with col4:
                st.metric("平均包裹重量", f"{avg_weight:.2f} kg")
                
            # 添加横线分隔
            st.markdown("---")
            
            # 显示各城市占比
            st.subheader("各城市数据占比")
            
            # 计算各城市的运单数和收入
            city_stats = {}
            for city in cities:
                if city in data:
                    city_df = data[city]
                    city_keys = data[f"{city}_keys"]
                    city_stats[city] = {
                        "运单数": len(city_df),
                        "总收入": city_df[city_keys["fee_key"]].sum(),
                        "百分比": len(city_df) / len(df) * 100
                    }
            
            # 创建城市占比表格
            city_df = pd.DataFrame.from_dict(city_stats, orient='index')
            
            # 使用plotly创建条形图
            cities_fig = px.bar(
                city_df, 
                y=city_df.index, 
                x="运单数", 
                orientation='h',
                title="各城市运单数量对比",
                text="运单数",
                color=city_df.index
            )
            st.plotly_chart(cities_fig, use_container_width=True)
            
            # 使用plotly创建饼图
            cities_pie = px.pie(
                city_df,
                values="运单数",
                names=city_df.index,
                title="各城市运单量占比"
            )
            st.plotly_chart(cities_pie, use_container_width=True)
        else:
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
        if selected_city == "ALL":
            st.header(f"所有城市汇总 - 按周维度分析")
            
            # 添加重要统计数据的摘要
            st.subheader("核心统计指标")
            
            # 计算总体指标
            total_deliveries = len(df)
            total_revenue = df[fee_key].sum()
            total_couriers = df[receiver_key].nunique()
            avg_per_week = total_deliveries / df['周数'].nunique()
            
            # 展示指标卡片
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("总运单量", f"{total_deliveries:,}")
            
            with col2:
                st.metric("总收入", f"${total_revenue:,.2f}")
            
            with col3:
                st.metric("活跃派送员数", f"{total_couriers}")
            
            with col4:
                st.metric("平均周运单量", f"{avg_per_week:.0f}")
                
            # 添加横线分隔
            st.markdown("---")
            
            # 显示各城市占比
            st.subheader("各城市数据占比")
            
            # 计算各城市的运单数和收入
            city_stats = {}
            for city in cities:
                if city in data:
                    city_df = data[city]
                    city_keys = data[f"{city}_keys"]
                    city_stats[city] = {
                        "运单数": len(city_df),
                        "总收入": city_df[city_keys["fee_key"]].sum(),
                        "活跃派送员数": city_df[city_keys["receiver_key"]].nunique()
                    }
            
            # 创建城市占比表格
            city_df = pd.DataFrame.from_dict(city_stats, orient='index')
            
            col1, col2 = st.columns(2)
            
            with col1:
                # 使用plotly创建条形图 - 运单数
                cities_fig1 = px.bar(
                    city_df, 
                    y=city_df.index, 
                    x="运单数", 
                    orientation='h',
                    title="各城市运单数量",
                    text="运单数",
                    color=city_df.index
                )
                st.plotly_chart(cities_fig1, use_container_width=True)
            
            with col2:
                # 使用plotly创建条形图 - 总收入
                cities_fig2 = px.bar(
                    city_df, 
                    y=city_df.index, 
                    x="总收入", 
                    orientation='h',
                    title="各城市总收入",
                    text="总收入",
                    color=city_df.index
                )
                st.plotly_chart(cities_fig2, use_container_width=True)
            
            # 使用plotly创建条形图 - 活跃派送员数
            cities_fig3 = px.bar(
                city_df, 
                y=city_df.index, 
                x="活跃派送员数", 
                orientation='h',
                title="各城市活跃派送员数",
                text="活跃派送员数",
                color=city_df.index
            )
            st.plotly_chart(cities_fig3, use_container_width=True)
        else:
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
    # 验证密码
    verify_password()
    
    # 设置页面标题
    st.title("🚚 运输数据看板")
    st.markdown("### 澳洲本地派送数据分析")
    
    # 创建侧边栏
    st.sidebar.title("数据选择")
    
    # 添加数据源选择
    data_source = st.sidebar.radio(
        "选择数据源",
        ["上传Excel文件", "使用Supabase云存储数据"]
    )
    
    # 根据选择的数据源处理数据
    all_data = {}
    if data_source == "上传Excel文件":
        # 上传文件选项
        uploaded_files = st.sidebar.file_uploader("上传BPE派送Excel文件", type=["xlsx"], accept_multiple_files=True)
        
        if uploaded_files:
            with st.spinner("正在处理数据..."):
                all_data = process_data(_uploaded_files=uploaded_files)
        else:
            st.info("请上传Excel文件...")
    else:
        # Supabase云存储选项
        try:
            st.sidebar.markdown("### 数据选择设置")
            st.sidebar.markdown("""
            **请注意:**
            - 每个Excel文件对应一周的数据
            - 系统会自动选择最新的文件
            - 已下载的文件会缓存在本地以加快加载速度
            """)
            
            # 获取可用的Excel文件列表
            with st.spinner("正在获取云端文件列表..."):
                # 设置默认分析52周(一年)的数据
                # 允许用户选择分析的周数范围
                week_num = st.sidebar.slider(
                    "选择分析过去多少周的数据", 
                    min_value=1, 
                    max_value=104,  # 最多两年
                    value=52,       # 默认一年
                    step=1
                )
                
                # 添加强制刷新选项
                force_refresh = st.sidebar.checkbox("强制刷新(忽略本地缓存)", value=False)
                
                # 根据选择的周数获取文件
                excel_files = list_supabase_files(".xlsx", week_num=week_num)
                
                if not excel_files:
                    st.warning("云端存储中未找到Excel文件")
                else:
                    # 显示找到的文件数量和总大小
                    total_size_mb = sum([
                        f.get("metadata", {}).get("size", 0) / (1024 * 1024)
                        for f in excel_files
                    ])
                    
                    st.sidebar.info(f"找到过去 {week_num} 周的 {len(excel_files)} 个文件 (总大小: {total_size_mb:.1f}MB)")
                    
                    # 自动处理所有文件而不需要用户选择
                    with st.spinner(f"正在处理过去 {week_num} 周的 {len(excel_files)} 个文件..."):
                        # 下载选定的文件
                        temp_dir = "./data"
                        os.makedirs(temp_dir, exist_ok=True)
                        
                        # 准备本地文件对象供process_data使用
                        file_objects = []
                        
                        # 处理所有文件
                        selected_files = [f.get("name") for f in excel_files]
                        
                        # 首先获取文件的详细信息，包括更新时间
                        files_details = {}
                        for file_info in excel_files:
                            file_name = file_info.get("name")
                            files_details[file_name] = {
                                "updated_at": file_info.get("updated_at", ""),
                                "size": file_info.get("metadata", {}).get("size", 0)
                            }
                        
                        # 创建处理进度条
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        # 统计信息
                        stats = {
                            "total": len(selected_files),
                            "downloaded": 0,
                            "cached": 0,
                            "failed": 0
                        }
                        
                        for i, file_name in enumerate(selected_files):
                            try:
                                # 更新进度条和状态
                                progress = int((i / len(selected_files)) * 100)
                                progress_bar.progress(progress)
                                status_text.text(f"处理文件 {i+1}/{len(selected_files)}: {file_name}")
                                
                                # 检查本地是否已有该文件
                                local_path = os.path.join(temp_dir, file_name)
                                file_details = files_details.get(file_name, {})
                                
                                # 获取Supabase文件的更新时间
                                supabase_updated_at = file_details.get("updated_at", "")
                                
                                # 检查本地文件是否存在、非空
                                local_file_exists = os.path.exists(local_path) and os.path.getsize(local_path) > 0
                                
                                # 决定是否需要下载
                                need_download = True
                                
                                # 如果强制刷新，总是下载
                                if force_refresh:
                                    need_download = True
                                # 否则，检查本地文件是否存在和最后修改时间
                                elif local_file_exists:
                                    # 获取本地文件的最后修改时间
                                    local_mtime = os.path.getmtime(local_path)
                                    local_mtime_str = datetime.fromtimestamp(local_mtime).isoformat()
                                    
                                    # 比较本地文件时间和Supabase文件时间
                                    # 如果Supabase上的文件较新或无法比较，则下载
                                    if supabase_updated_at and local_mtime_str >= supabase_updated_at:
                                        need_download = False
                                        stats["cached"] += 1
                                
                                # 如果需要下载，则从Supabase获取
                                if need_download:
                                    local_path = download_supabase_file(file_name, temp_dir)
                                    stats["downloaded"] += 1
                                
                                # 创建一个类文件对象以匹配process_data的预期
                                class FileObject:
                                    def __init__(self, path, name):
                                        self.path = path
                                        self.name = name
                                
                                file_obj = FileObject(local_path, file_name)
                                file_objects.append(file_obj)
                            except Exception as e:
                                st.error(f"处理文件 {file_name} 时出错: {str(e)}")
                                stats["failed"] += 1
                        
                        # 完成进度条
                        progress_bar.progress(100)
                        status_text.text(f"已完成 {len(file_objects)}/{len(selected_files)} 个文件的处理")
                        
                        # 显示统计信息
                        st.success(f"已处理 {stats['total']} 个文件: {stats['downloaded']} 个下载, {stats['cached']} 个使用缓存, {stats['failed']} 个失败")
                        
                        if file_objects:
                            all_data = process_data(_uploaded_files=file_objects)
                        else:
                            st.error("未能成功下载任何文件")
        except Exception as e:
            st.error(f"访问云存储时出错: {str(e)}")
    
    # 创建看板（如果有数据）
    if all_data:
        create_dashboard(all_data)
    else:
        # 展示说明信息
        st.markdown("""
        ## 使用说明
        1. 通过左侧栏选择数据源
        2. 上传Excel文件或从云端选择文件
        3. 数据加载后可查看不同维度的分析
        """)

if __name__ == "__main__":
    main() 