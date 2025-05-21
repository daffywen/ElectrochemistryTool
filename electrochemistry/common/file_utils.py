"""
文件处理工具模块
包含文件选择、文件类型识别等功能
"""
import os
import re
import logging
from typing import List, Optional, Tuple, Dict, Any
from pathlib import Path
import tkinter as tk
from tkinter import filedialog

logger = logging.getLogger(__name__)

def select_folder() -> str:
    """选择一个文件夹并返回其路径"""
    root = tk.Tk()
    # root.withdraw()  # 隐藏主窗口
    root.withdraw()  # Hide the main window
    folder_path = filedialog.askdirectory(title="选择包含电化学数据文件的文件夹")
    root.destroy()
    return folder_path

def is_cv_file(file_path: str) -> bool:
    """
    检查文件是否为循环伏安法(CV)数据文件
    通过查找文件内容中的关键词判断
    """
    try:
        with open(file_path, 'r', errors='ignore') as f:
            header = f.read(1000)
            is_cv = ('Cyclic Voltammetry' in header or 
                     'CYCLIC VOLTAMMETRY' in header)
            is_not_cv = any(keyword in header for keyword in [
                'Linear Sweep Voltammetry', 
                'A.C. Impedance',
                'LSV',
                'Chronoamperometry',
                'Open Circuit',
                'EIS',
                'Tafel'
            ])
            return is_cv and not is_not_cv
    except Exception as e:
        logger.warning(f"检查文件{file_path}时出错: {str(e)}")
        return False

def is_lsv_file(file_path: str) -> bool:
    """
    检查文件是否为线性扫描伏安法(LSV)数据文件
    通过查找文件内容中的关键词判断
    """
    try:
        with open(file_path, 'r', errors='ignore') as f:
            header = f.read(1000)
            is_lsv = ('Linear Sweep Voltammetry' in header or 
                      'LINEAR SWEEP VOLTAMMETRY' in header or
                      'LSV' in header)
            return is_lsv
    except Exception as e:
        logger.warning(f"检查文件{file_path}时出错: {str(e)}")
        return False

def is_eis_file(file_path: str) -> bool:
    """
    检查文件是否为电化学阻抗谱(EIS)数据文件
    通过查找文件内容中的关键词判断
    """
    try:
        with open(file_path, 'r', errors='ignore') as f:
            header = f.read(1000)
            is_eis = ('A.C. Impedance' in header or 
                      'Electrochemical Impedance' in header or
                      'EIS' in header)
            return is_eis
    except Exception as e:
        logger.warning(f"检查文件{file_path}时出错: {str(e)}")
        return False

def find_files_by_type(folder_path: str, file_type: str) -> List[str]:
    """
    在指定文件夹中查找指定类型的数据文件
    
    参数:
        folder_path: 文件夹路径
        file_type: 文件类型，如 'cv', 'lsv', 'eis'
    
    返回:
        匹配文件路径列表
    """
    if not os.path.exists(folder_path):
        logger.error(f"文件夹 {folder_path} 不存在")
        return []
    
    # 获取文件夹中所有txt文件
    txt_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) 
                if f.lower().endswith('.txt')]
    
    if not txt_files:
        logger.warning(f"在{folder_path}中未找到任何.txt文件")
        return []
    
    # 根据文件类型选择合适的检测函数
    if file_type.lower() == 'cv':
        check_func = is_cv_file
    elif file_type.lower() == 'lsv':
        check_func = is_lsv_file
    elif file_type.lower() == 'eis':
        check_func = is_eis_file
    else:
        logger.error(f"不支持的文件类型: {file_type}")
        return []
    
    # 检查每个文件并收集匹配的文件
    matching_files = []
    for file_path in txt_files:
        if check_func(file_path):
            matching_files.append(file_path)
    
    return matching_files

def ensure_output_dir(folder_path: str) -> Tuple[str, str]:
    """
    确保输出目录存在，返回输出目录和基础文件名
    
    参数:
        folder_path: 原始文件夹路径
    
    返回:
        (输出目录路径, 基础文件名)
    """
    folder_basename = os.path.basename(folder_path.rstrip("\\/"))
    output_dir = os.path.join(folder_path, "processed_data")
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    return output_dir, folder_basename