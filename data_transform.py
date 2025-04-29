import pandas as pd
import os
from datetime import datetime

def transform_data(input_file, output_file):
    # 定义列名映射
    column_mapping = {
        "业务时间": "签收时间",
        "业务单号": "运单号",
        "派件城市": "收件城市",
        "fee_after_adjust": "计费",
        "司机名称": "收派员driver"
    }
    
    # 定义州名和对应的sheet名映射
    state_mapping = {
        "New South Wales": "SYD",
        "South Australia": "ADL",
        "Queensland": "BNE"
    }
    
    # 读取原始Excel文件
    df = pd.read_excel(input_file)
    
    # 将业务时间转换为datetime类型
    df['业务时间'] = pd.to_datetime(df['业务时间'])
    
    # 找到最早的签收时间
    earliest_date = df['业务时间'].min()
    
    # 计算周数
    df['周数'] = ((df['业务时间'] - earliest_date).dt.days // 7) + 1
    
    # 创建Excel写入器
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # 对每个州进行处理
        for state_name, sheet_name in state_mapping.items():
            # 筛选对应州的数据
            state_df = df[df['派件省份'] == state_name].copy()
            
            # 重命名列
            state_df = state_df.rename(columns=column_mapping)
            
            # 选择需要的列
            required_columns = [
                "签收时间",
                "运单号",
                "收件城市",
                "计费",
                "收派员driver",
                "重量",
                "派件省份",
                "周数"
            ]
            
            # 确保所有需要的列都存在
            state_df = state_df[required_columns]
            
            # 写入对应的sheet
            state_df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # 打印每个sheet的周数范围
            print(f"{sheet_name} sheet的周数范围: {state_df['周数'].min()} - {state_df['周数'].max()}")

if __name__ == "__main__":
    # 设置输入输出文件路径
    input_file = "input.xlsx"  # 请替换为实际的输入文件路径
    output_file = "output.xlsx"  # 请替换为实际的输出文件路径
    
    # 执行转换
    transform_data(input_file, output_file)
    print("数据转换完成！") 