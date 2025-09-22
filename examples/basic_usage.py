# -*- coding: utf-8 -*-
"""
Created on Tue Sep 23 00:03:42 2025

@author: 17283
"""

"""
Basic usage examples for my-data-toolkit-py
"""

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

from src.file_processor import LocalFolder
from src.dataframe_processor import DataFrameOpe

def main():
    # 示例用法
    folder_path = "./sample_data"  # 使用相对路径示例
    
    # 初始化文件夹处理器
    folder_processor = LocalFolder(folder_path)
    print(f"Found {len(folder_processor.file_list)} files")
    
    # 查找特定文件
    folder_processor.find_file_name(['.xlsx', '.csv'], ('data',))
    if folder_processor.target_file:      
        print(f"Target file: {folder_processor.target_file}")
        
        # 读取文件
        df = folder_processor.read_excel('data', tab=0, header=0)
        print(f"DataFrame shape: {df.shape}")
        
        # 处理 DataFrame
        df_processor = DataFrameOpe()
        processed_df = df_processor.reset_index_by_header(df)
        print(f"Processed DataFrame shape: {processed_df.shape}")

if __name__ == "__main__":
    main()