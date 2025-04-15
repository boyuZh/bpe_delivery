import os
import json
import pandas as pd
import requests
from pathlib import Path
from dotenv import load_dotenv

# 尝试加载.env文件中的环境变量
load_dotenv(override=True)

def get_supabase_config():
    """
    按照优先级获取Supabase配置:
    1. 环境变量
    2. Streamlit Secrets
    4. 默认值（仅用于开发）
    """
    # 优先尝试从环境变量获取
    url = os.getenv("SUPABASE_URL")
    key = os.getenv("SUPABASE_KEY")
    bucket = os.getenv("SUPABASE_BUCKET", "bpedeliverydata")
    
    # 如果环境变量不存在，尝试从Streamlit Secrets获取
    if (not url or not key) and 'st' in globals():
        try:
            import streamlit as st
            if "supabase" in st.secrets:
                url = url or st.secrets["supabase"]["url"]
                key = key or st.secrets["supabase"]["key"]
                bucket = bucket or st.secrets["supabase"].get("bucket", "bpedeliverydata")
        except Exception:
            pass
    
    return {
        "url": url,
        "key": key,
        "bucket": bucket
    }

def get_supabase_client():
    """获取Supabase客户端实例"""
    try:
        from supabase import create_client, Client
        
        config = get_supabase_config()
        if not config["url"] or not config["key"]:
            raise ValueError("缺少Supabase配置，请配置环境变量")
        
        supabase = create_client(config["url"], config["key"])
        return supabase
    except ImportError:
        raise ImportError("未安装supabase包，请运行: pip install supabase")

def list_supabase_files(file_extension=".xlsx", week_num=52):
    """列出Supabase存储桶中的文件"""
    config = get_supabase_config()
    bucket = config["bucket"]
    
    if not config["url"] or not config["key"]:
        raise ValueError("缺少Supabase配置，请配置环境变量")
    
    # 使用supabase-py客户端
    supabase = get_supabase_client()
    
    # 获取文件列表
    response = (
        supabase.storage
        .from_(bucket)
        .list(
            "",
            {
                "limit": week_num,
                "offset": 0,
                "sortBy": {"column": "created_at", "order": "desc"},
            }
        )
    )
    
    # 根据扩展名过滤
    if file_extension:
        response = [f for f in response if f.get("name", "").endswith(file_extension)]
    
    return response

def download_supabase_file(file_name, output_dir="./data"):
    """从Supabase下载指定文件并返回本地路径"""
    config = get_supabase_config()
    bucket = config["bucket"]
    
    if not config["url"] or not config["key"]:
        raise ValueError("缺少Supabase配置，请配置环境变量")
    
    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)
    local_path = os.path.join(output_dir, file_name)
    
    # 使用supabase-py客户端
    supabase = get_supabase_client()
    
    # 下载文件
    file_data = (
        supabase.storage
        .from_(bucket)
        .download(file_name)
    )
    
    # 保存到本地
    with open(local_path, "wb") as f:
        f.write(file_data)
    
    return local_path

def download_all_excel_files(output_dir="./data"):
    """下载所有Excel文件并返回文件路径列表"""
    excel_files = list_supabase_files(".xlsx")
    downloaded_files = []
    
    for file_data in excel_files:
        file_name = file_data.get("name")
        try:
            local_path = download_supabase_file(file_name, output_dir)
            downloaded_files.append(local_path)
        except Exception as e:
            print(f"下载文件 {file_name} 失败: {e}")
    
    return downloaded_files
