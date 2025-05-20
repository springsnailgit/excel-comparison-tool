import os
from datetime import datetime


def save_to_new_excel(dataframe, filename):
    import pandas as pd

    # Create a new Excel file with the specified filename
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    full_filename = f"{filename}_{timestamp}.xlsx"
    
    # Save the DataFrame to the new Excel file
    dataframe.to_excel(full_filename, index=False)


def copy_rows_to_new_sheet(source_df, condition, new_sheet_name):
    # Filter the DataFrame based on the condition
    filtered_df = source_df[source_df.apply(lambda row: condition in row.values, axis=1)]
    
    return filtered_df


def delete_filtered_rows(source_df, condition):
    # Delete rows that match the condition
    return source_df[~source_df.apply(lambda row: condition in row.values, axis=1)]


def get_file_extension(file_path):
    """获取文件扩展名"""
    return os.path.splitext(file_path)[1].lower()


def is_excel_file(file_path):
    """检查文件是否为Excel文件"""
    valid_extensions = ['.xlsx', '.xls', '.xlsm']
    return get_file_extension(file_path) in valid_extensions


def generate_output_filename(original_path):
    """生成输出文件名，格式为：原文件名_比对结果_时间戳.xlsx"""
    dir_name = os.path.dirname(original_path)
    file_name = os.path.basename(original_path)
    name_without_ext = os.path.splitext(file_name)[0]
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    
    return os.path.join(dir_name, f"{name_without_ext}_比对结果_{timestamp}.xlsx")


def ensure_directory_exists(directory_path):
    """确保目录存在，如果不存在则创建"""
    if not os.path.exists(directory_path):
        os.makedirs(directory_path)
    return directory_path