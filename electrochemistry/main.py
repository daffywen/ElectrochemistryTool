"""
电化学数据处理工具主程序
提供自动处理不同电化学数据的功能
"""
import os
import sys
import logging
import importlib.util
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox # 导入 messagebox
from tkinter import ttk # 导入 ttk 模块
import time  # 为“仪式感”的延迟添加
from .common import excel_utils # 添加导入
from . import tafel # <--- 添加这一行

# 设置日志记录
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

def setup_environment():
    """设置环境，确保路径正确"""
    # 获取当前脚本所在的目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 将该目录添加到Python路径 (electrochemistry 包)
    if script_dir not in sys.path:
        sys.path.insert(0, script_dir)
    
    # 确保上级目录也在路径中 (cursor 目录, 用于 run_electrochemistry.py)
    parent_dir = os.path.dirname(script_dir) # 这是 cursor 目录
    if parent_dir not in sys.path:
        sys.path.insert(0, parent_dir)
        
    # 设置日志文件 (位于 parent_dir 中的 logs 文件夹, 即 cursor/logs)
    log_dir = os.path.join(parent_dir, "logs")
    os.makedirs(log_dir, exist_ok=True)
    
    # 添加文件处理器
    log_file = os.path.join(log_dir, f"electrochemistry_{datetime.now().strftime('%Y%m%d')}.log")
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    
    # 添加到根日志记录器
    # logger.addHandler(file_handler) # 移除直接添加，以避免在已配置日志记录器时产生重复日志
    # 相反，获取根日志记录器，并在没有处理器或需要特定配置时添加处理器
    root_logger = logging.getLogger() # 获取根日志记录器
    if not root_logger.hasHandlers(): # 仅在没有配置处理器时添加处理器
        root_logger.addHandler(file_handler)
    # 如果希望始终添加此特定文件处理器，
    # 可能需要检查是否已存在类似的处理器或清除现有处理器。
    # 为简单起见，此代码在没有处理器时添加。
    # 如果 electrochemistry.main 也配置了日志记录，这可能会导致重复日志或覆盖。
    # 如果出现问题，请考虑集中式日志记录设置。

    logger.info("程序启动")
    print("[初始化] 环境设置完成，日志系统已启动。")
    time.sleep(0.1)  # 减少延迟

def print_header():
    """打印程序头部信息"""
    try:
        from . import __version__
        version = __version__
    except ImportError:
        version = "1.0.0"  # 如果未找到，则使用默认版本
        
    header_line = "=" * 60
    title = f"电化学数据处理工具 v{version}"
    empty_line_for_title = " " * ((60 - len(title.encode('gbk')) + len(title)) // 2) # 尝试居中

    print("\\n" + header_line)
    print(empty_line_for_title + title)
    print(header_line)
    print("  作者: [您的名字或团队名称] | 联系方式: [您的联系方式]") # 可以替换为实际信息
    print("  欢迎使用！请按照提示操作。")
    print(header_line + "\\n")
    time.sleep(0.1)

def select_folder():
    """选择一个文件夹并返回其路径"""
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口

    # 设置 ttk 主题，尝试使用更现代的外观
    style = ttk.Style(root)
    available_themes = style.theme_names()
    # print(f"Available themes: {available_themes}") # 用于调试，查看可用主题
    if 'clam' in available_themes:
        style.theme_use('clam')
    elif 'alt' in available_themes:
        style.theme_use('alt')
    elif 'default' in available_themes: # ttk 的 default 通常比 tkinter 的好
        style.theme_use('default')
    # 其他常见主题：'vista', 'xpnative' (Windows)

    # 直接打开文件夹选择对话框，不显示提示信息框
    folder_path = filedialog.askdirectory(
        title="请选择包含电化学数据文件的文件夹",
        parent=root # 确保对话框是顶层
    )
    
    # 在销毁 root 之前短暂延迟，确保 messagebox 完全消失
    # root.after(100, root.destroy) # 延迟销毁，有时有助于窗口管理
    # 或者直接销毁
    root.destroy()

    if folder_path:
        print(f"\\n[INFO] 已选择文件夹: {folder_path}")
        logger.info(f"用户选择了文件夹: {folder_path}")
    else:
        print("\\n[INFO] 用户未选择任何文件夹。")
        logger.warning("用户取消了文件夹选择。")
    time.sleep(0.1)
    return folder_path

def load_module(module_name):
    """加载指定的模块 (保留此动态加载方式以减少对现有结构的更改)"""
    try:
        if module_name == 'cv':
            from . import cv
            return cv
        elif module_name == 'lsv':
            from . import lsv
            return lsv
        elif module_name == 'eis':
            from . import eis
            return eis
    except ImportError as e:
        logger.warning(f"作为包内模块导入 {module_name} 失败: {e}. 尝试其他方法...")
        pass
    
    # 尝试作为独立模块导入
    try:
        module = __import__(module_name)
        return module
    except ImportError:
        # 检查模块文件是否存在
        module_path = os.path.join(os.path.dirname(__file__), f"{module_name}.py")
        if os.path.exists(module_path):
            try:
                # 尝试动态导入
                spec = importlib.util.spec_from_file_location(module_name, module_path)
                if spec and spec.loader:
                    module = importlib.util.module_from_spec(spec)
                    sys.modules[module_name] = module
                    spec.loader.exec_module(module)
                    return module
                else:
                    logger.error(f"无法为 {module_name} 创建模块规范于 {module_path}")
            except Exception as e:
                logger.error(f"动态加载 {module_name} 模块时发生错误: {str(e)}")
        
        logger.error(f"找不到{module_name}模块，请确保{module_name}.py文件存在且可导入")
        return None

def process_all_data(folder_path):
    """处理文件夹中的所有电化学数据
    
    参数:
        folder_path: 包含数据文件的文件夹路径
    """
    if not folder_path:
        logger.warning("未选择文件夹，程序退出")
        return
        
    logger.info(f"已选择文件夹: {folder_path}")
    
    # 输出文件设置
    folder_basename = os.path.basename(folder_path.rstrip("\\/"))
    # 新建 processed_data 文件夹（如果不存在）
    output_dir = os.path.join(folder_path, "processed_data")
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    # 添加时间戳到文件名以避免覆盖
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = os.path.join(output_dir, f"{folder_basename}_processed_data_{timestamp}.xlsx")
    
    print("\n[阶段] 准备加载数据处理模块...")
    time.sleep(0.2)  # 减少延迟
    # 加载所有模块
    modules_to_process = {}
    for module_name_key in ['cv', 'lsv', 'eis']:
        module_obj = load_module(module_name_key)
        if module_obj:
            modules_to_process[module_name_key] = module_obj
            print(f"  [模块] {module_name_key.upper()} 模块已加载。")
            time.sleep(0.05)  # 减少延迟
        else:
            print(f"  [警告] 无法加载 {module_name_key.upper()} 模块，相关数据处理将跳过。")
            time.sleep(0.05)  # 减少延迟
    
    if not modules_to_process:
        print("\n[错误] 未能加载任何有效的数据处理模块。程序无法继续。")
        time.sleep(0.2)  # 减少延迟
        return
    print("[阶段] 所有可用模块加载完毕。")
    time.sleep(0.2)  # 减少延迟
        
    # 准备共享的Excel工作簿
    wb = None
    processed_count = 0
    
    # --- 存储各模块的分析数据 ---
    lsv_analysis_results = []  # 元组列表: (file_id, op_10mV, op_100mV, op_200mV)
    eis_analysis_results = []  # 字典列表: {'file_id': file_id, 'rs': rs_value}
    cv_analysis_results = {}   # 字典: {'cdl': cdl_value, 'folder_basename': folder_basename}
    processed_lsv_sheet_names_for_tafel = [] # <--- 添加这一行

    print("\n[阶段] 开始查找各类数据文件...")
    time.sleep(0.2)  # 减少延迟
    # 依次处理每种类型的数据
    try:
        # 首先查找所有相关模块的文件
        file_lists = {}
        for mod_name, mod_obj in modules_to_process.items():
            print(f"  [查找] 正在为 {mod_name.upper()} 模块查找文件...")
            time.sleep(0.1)  # 减少延迟
            find_func_name = f"find_{mod_name}_files"
            if hasattr(mod_obj, find_func_name):
                find_func = getattr(mod_obj, find_func_name)
                files = find_func(folder_path)
                if files:
                    logger.info(f"找到 {len(files)} 个{mod_name.upper()}数据文件")
                    print(f"    [发现] 找到 {len(files)} 个 {mod_name.upper()} 数据文件。")
                    file_lists[mod_name] = files
                else:
                    print(f"    [提示] 在 {folder_path} 中未找到 {mod_name.upper()} 数据文件。")
            else:
                logger.warning(f"模块 {mod_name}缺少 {find_func_name} 方法。")
                print(f"    [错误] 模块 {mod_name.upper()} 缺少文件查找功能。")
            time.sleep(0.1)  # 减少延迟
        
        print("[阶段] 文件查找完成。")
        time.sleep(0.2)  # 减少延迟

        if not any(file_lists.values()):
            print("\n[结果] 在选定的文件夹中未找到任何可处理的电化学数据文件。")
            print("请确保文件内容包含相关的电化学测试信息。")
            time.sleep(0.2)  # 减少延迟
            return
        
        print("\n[阶段] 开始处理数据...")
        time.sleep(0.2)  # 减少延迟
        # 如果找到文件则处理数据
        for mod_name in ['cv', 'lsv', 'eis']:
            if mod_name in modules_to_process and mod_name in file_lists:
                mod_obj = modules_to_process[mod_name]
                files_to_process = file_lists[mod_name]
                
                process_func_name = "process_all_files_from_paths"
                if hasattr(mod_obj, process_func_name):
                    logger.info(f"开始处理{mod_name.upper()}数据")
                    print(f"\n-> 即将处理 {mod_name.upper()} 数据 ({len(files_to_process)} 个文件)...")
                    time.sleep(0.1)  # 减少延迟
                    
                    process_func = getattr(mod_obj, process_func_name)
                    print(f"  [处理] 调用 {mod_name.upper()} 模块处理 {len(files_to_process)} 个文件...")
                    time.sleep(0.1)
                    
                    returned_data = None
                    analysis_payload = None # 初始化 analysis_payload

                    if mod_name == 'eis':
                        # EIS 需要 original_folder_path，此处即 folder_path
                        returned_data = process_func(files_to_process, output_file, folder_basename, folder_path, wb)
                    else:
                        returned_data = process_func(files_to_process, output_file, folder_basename, wb)
                    
                    if returned_data and returned_data[0]: # 检查是否返回工作簿
                        wb = returned_data[0] # 更新工作簿
                        if len(returned_data) > 1:
                            analysis_payload = returned_data[1] # 获取分析数据
                        else:
                            analysis_payload = None 
                            logger.warning(f"{mod_name.upper()} 模块未返回分析负载。")

                        if mod_name == 'lsv':
                            lsv_analysis_results = analysis_payload if analysis_payload is not None else []
                            # MODIFIED: Correctly access the third element for sheet names
                            if len(returned_data) > 2 and returned_data[2] is not None and isinstance(returned_data[2], list):
                                processed_lsv_sheet_names_for_tafel.extend(returned_data[2])
                                logger.info(f"LSV sheets processed and explicitly returned for Tafel: {returned_data[2]}")
                            else:
                                logger.warning(f"LSV module did not explicitly return a list of sheet names (expected at index 2 of return tuple). Will attempt to use default 'LSV Data' if available.")
                                # Fallback will be handled before calling Tafel

                        elif mod_name == 'eis':
                            eis_analysis_results = analysis_payload if analysis_payload is not None else []
                        elif mod_name == 'cv':
                            cv_analysis_results = analysis_payload if analysis_payload is not None else {}
                        
                        processed_count += 1
                        print(f"  [成功] {mod_name.upper()} 数据处理完成。")
                    else:
                        print(f"  [失败] {mod_name.upper()} 数据处理失败或未生成/更新结果文件。")
                    time.sleep(0.1)  # 减少延迟
                else:
                    logger.warning(f"模块 {mod_name} 缺少 {process_func_name} 方法。")
                    print(f"  [错误] 模块 {mod_name.upper()} 缺少核心处理功能。")
                    time.sleep(0.1)  # 减少延迟
            elif mod_name in modules_to_process and mod_name not in file_lists:
                 logger.info(f"未找到 {mod_name.upper()} 数据文件，跳过处理。")
        
        print("\n[阶段] 所有数据处理尝试完毕。")
        time.sleep(0.2)
        
        # --- 创建分析报告工作表 ---
        if wb and (lsv_analysis_results or eis_analysis_results or cv_analysis_results.get('cdl') is not None):
            try:
                logger.info("开始创建分析报告工作表...")
                print("  [报告] 正在创建分析报告...")
                report_ws_name = "Analysis Report" # 已改为英文
                
                # 如果工作表已存在，则移除并在开头重新创建
                if report_ws_name in wb.sheetnames:
                    idx = wb.sheetnames.index(report_ws_name)
                    wb.remove(wb.worksheets[idx])
                
                report_ws = wb.create_sheet(report_ws_name, 0) # 在第一个位置创建
                logger.info(f"已创建新的分析报告工作表: {report_ws_name} (置于最前)")

                # header_fill, thin_border, center_aligned, openpyxl_module = excel_utils.get_excel_styles()
                # Font = openpyxl_module.styles.Font # 直接获取 Font 类
                # bold_font = Font(bold=True)
                # 使用新的工具函数
                header_fill, thin_border, center_aligned, openpyxl_module = excel_utils.get_excel_styles()
                bold_font = excel_utils.get_bold_font()
                left_aligned = openpyxl_module.styles.Alignment(horizontal='left', vertical='center')
                right_aligned = openpyxl_module.styles.Alignment(horizontal='right', vertical='center')

                current_row = 1

                report_ws.cell(row=current_row, column=1).value = f"{folder_basename} - Electrochemical Analysis Summary" # 已改为英文
                title_font = openpyxl_module.styles.Font(bold=True, size=14) # 直接从 openpyxl_module 定义 Font
                report_ws.cell(row=current_row, column=1).font = title_font
                report_ws.cell(row=current_row, column=1).fill = header_fill # 为主标题应用填充颜色
                report_ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=6) # 为新的LSV列合并到第6列
                report_ws.cell(row=current_row, column=1).alignment = center_aligned
                # 为标题单元格应用边框
                for c_idx in range(1, 7): # 为合并范围内的所有单元格应用边框 (最多到6)
                    report_ws.cell(row=current_row, column=c_idx).border = thin_border
                current_row += 2

                if lsv_analysis_results: # Check if there are LSV analysis results to write
                    report_ws.cell(row=current_row, column=1).value = "LSV Analysis (Overpotential)" 
                    report_ws.cell(row=current_row, column=1).font = bold_font
                    report_ws.cell(row=current_row, column=1).fill = header_fill
                    report_ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=4) 
                    report_ws.cell(row=current_row, column=1).alignment = center_aligned
                    for c_idx in range(1, 5): 
                        report_ws.cell(row=current_row, column=c_idx).border = thin_border
                    current_row += 1
                    
                    headers_lsv = ["File ID", "Overpotential @10 mA·cm⁻² (mV)", "Overpotential @100 mA·cm⁻² (mV)", "Overpotential @200 mA·cm⁻² (mV)"] 
                    for col_idx, header_title in enumerate(headers_lsv, start=1):
                        cell = report_ws.cell(row=current_row, column=col_idx, value=header_title)
                        if bold_font: cell.font = bold_font
                        cell.fill = header_fill
                        cell.border = thin_border
                        cell.alignment = center_aligned
                    current_row += 1

                    # Ensure lsv_analysis_results is a list of tuples/lists with 4 elements
                    for item in lsv_analysis_results:
                        if isinstance(item, (list, tuple)) and len(item) == 4:
                            file_id, op10, op100, op200 = item 
                            report_ws.cell(row=current_row, column=1, value=file_id).alignment = left_aligned
                            cell_op10 = report_ws.cell(row=current_row, column=2, value=op10)
                            cell_op10.number_format = '0.0'; cell_op10.alignment = right_aligned
                            cell_op100 = report_ws.cell(row=current_row, column=3, value=op100)
                            cell_op100.number_format = '0.0'; cell_op100.alignment = right_aligned
                            cell_op200 = report_ws.cell(row=current_row, column=4, value=op200) 
                            cell_op200.number_format = '0.0'; cell_op200.alignment = right_aligned
                            for col_idx_data in range(1, 5): 
                                 report_ws.cell(row=current_row, column=col_idx_data).border = thin_border
                            current_row += 1
                        else:
                            logger.warning(f"Skipping malformed LSV analysis item: {item}")
                    current_row += 1
                
                if eis_analysis_results: # Check if there are EIS analysis results
                    report_ws.cell(row=current_row, column=1).value = "EIS Analysis (Solution Resistance)" # 已改为英文
                    report_ws.cell(row=current_row, column=1).font = bold_font
                    report_ws.cell(row=current_row, column=1).fill = header_fill
                    report_ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=2)
                    report_ws.cell(row=current_row, column=1).alignment = center_aligned
                    # 为节标题应用边框
                    for c_idx in range(1, 3):
                        report_ws.cell(row=current_row, column=c_idx).border = thin_border
                    current_row += 1

                    headers_eis = ["File ID", "Solution Resistance (Rs) / Ohm"] # 已改为英文
                    for col_idx, header_title in enumerate(headers_eis, start=1):
                        cell = report_ws.cell(row=current_row, column=col_idx, value=header_title)
                        # cell.font = bold_font; cell.fill = header_fill; cell.border = thin_border; cell.alignment = center_aligned
                        if bold_font: cell.font = bold_font
                        cell.fill = header_fill
                        cell.border = thin_border
                        cell.alignment = center_aligned
                    current_row += 1

                    for item in eis_analysis_results:
                        report_ws.cell(row=current_row, column=1, value=item.get('file_id')).alignment = left_aligned
                        cell_rs = report_ws.cell(row=current_row, column=2, value=item.get('rs'))
                        cell_rs.number_format = '0.0000'; cell_rs.alignment = right_aligned
                        for col_idx_data in range(1, 3):
                             report_ws.cell(row=current_row, column=col_idx_data).border = thin_border
                        current_row += 1
                    current_row += 1

                if cv_analysis_results and cv_analysis_results.get('cdl') is not None:
                    report_ws.cell(row=current_row, column=1).value = "CV Analysis (Double Layer Capacitance)" # 已改为英文
                    report_ws.cell(row=current_row, column=1).font = bold_font
                    report_ws.cell(row=current_row, column=1).fill = header_fill
                    report_ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=2)
                    report_ws.cell(row=current_row, column=1).alignment = center_aligned
                     # 为节标题应用边框
                    for c_idx in range(1, 3):
                        report_ws.cell(row=current_row, column=c_idx).border = thin_border
                    current_row += 1

                    headers_cv = ["Parameter", "Value"] # 已改为英文
                    for col_idx, header_title in enumerate(headers_cv, start=1):
                        cell = report_ws.cell(row=current_row, column=col_idx, value=header_title)
                        # cell.font = bold_font; cell.fill = header_fill; cell.border = thin_border; cell.alignment = center_aligned
                        if bold_font: cell.font = bold_font
                        cell.fill = header_fill
                        cell.border = thin_border
                        cell.alignment = center_aligned
                    current_row += 1
                    
                    cdl_val = cv_analysis_results.get('cdl')
                    cdl_display = f"{cdl_val:.2f} mF·cm⁻²" if isinstance(cdl_val, float) else "N/A"
                    
                    report_ws.cell(row=current_row, column=1, value=f"Cdl (from {cv_analysis_results.get('folder_basename', 'N/A')} dataset)").alignment = left_aligned # 已改为英文
                    report_ws.cell(row=current_row, column=2, value=cdl_display).alignment = right_aligned
                    for col_idx_data in range(1, 3):
                        report_ws.cell(row=current_row, column=col_idx_data).border = thin_border
                    current_row += 1
                
                # report_ws.column_dimensions[openpyxl_module.utils.get_column_letter(1)].width = 35
                # report_ws.column_dimensions[openpyxl_module.utils.get_column_letter(2)].width = 30
                # report_ws.column_dimensions[openpyxl_module.utils.get_column_letter(3)].width = 30
                # 使用工具函数设置列宽
                excel_utils.set_column_widths(report_ws, {
                    openpyxl_module.utils.get_column_letter(1): 35,
                    openpyxl_module.utils.get_column_letter(2): 30,
                    openpyxl_module.utils.get_column_letter(3): 30,
                    openpyxl_module.utils.get_column_letter(4): 30 # 新LSV列的宽度
                })
                
                logger.info("分析报告工作表创建/更新完成。")
                print("  [报告] 分析报告已生成。")

            except Exception as e_report:
                logger.error(f"创建分析报告时发生错误: {e_report}", exc_info=True)
                print(f"  [错误] 创建分析报告时出错: {e_report}")
    
    except Exception as e_main_processing: # 添加以捕获主try块中的错误
        logger.error(f"处理数据时发生主要错误: {e_main_processing}", exc_info=True)
        print(f"\n[严重错误] 数据处理过程中发生意外错误: {e_main_processing}")
    finally: # 添加以确保打印此消息
        logger.info("process_all_data 函数执行完毕或遇到错误后终止。")
        print("\n[信息] process_all_data 函数执行流程结束。")


    # 工作簿的最终保存
    try:
        if wb:
            # --- 在保存之前调用 Tafel 处理 ---
            if processed_lsv_sheet_names_for_tafel and 'eis' in modules_to_process: # 确保有LSV数据和EIS模块（用于Rs）
                logger.info("Calling Tafel data processing.")
                print("\n[阶段] 开始处理 Tafel 数据...")
                time.sleep(0.1)
                try:
                    tafel.process_tafel_data(wb, eis_analysis_results, processed_lsv_sheet_names_for_tafel, folder_basename)
                    logger.info("Tafel data processing completed.")
                    print("  [成功] Tafel 数据处理完成。")
                except Exception as e_tafel:
                    logger.error(f"Tafel data processing failed: {e_tafel}", exc_info=True)
                    print(f"  [错误] Tafel 数据处理失败: {e_tafel}")
                time.sleep(0.1)
            elif not processed_lsv_sheet_names_for_tafel:
                logger.info("No LSV sheets were processed or identified, skipping Tafel processing.")
                print("\n[提示] 未处理或识别任何LSV工作表，跳过Tafel数据处理。")
            elif 'eis' not in modules_to_process:
                logger.info("EIS module not loaded or no EIS data, Rs values might be unavailable for Tafel. Skipping Tafel processing.")
                print("\n[提示] EIS模块未加载或无EIS数据，可能无法获取Rs值，跳过Tafel数据处理。")
            # ---------------------------------

            # 确保 "Analysis Report" 是活动工作表，如果存在的话
            # 如果Tafel表创建在Analysis Report之后，这可能需要调整，或者让Tafel表成为活动表
            analysis_report_sheet_name = "Analysis Report"
            tafel_sheet_name = "Tafel Data"

            # Set "Analysis Report" as the active sheet if it exists
            if analysis_report_sheet_name in wb.sheetnames:
                try:
                    # "Analysis Report" is created at index 0
                    wb.active = wb.sheetnames.index(analysis_report_sheet_name) 
                    logger.info(f"Set '{analysis_report_sheet_name}' as the active sheet.")
                except ValueError: # Should not happen if sheet name is in sheetnames
                    logger.warning(f"Could not find sheet '{analysis_report_sheet_name}' by name to set active, though it was expected.")
            elif tafel_sheet_name in wb.sheetnames: # Fallback if "Analysis Report" is not there for some reason
                try:
                    wb.active = wb.sheetnames.index(tafel_sheet_name)
                    logger.info(f"Set '{tafel_sheet_name}' as the active sheet ('{analysis_report_sheet_name}' not found).")
                except ValueError:
                    logger.warning(f"Could not find sheet '{tafel_sheet_name}' by name to set as fallback active sheet.")
            
            wb.save(output_file)
            logger.info(f"最终处理后的Excel文件已保存到: {output_file}")
            print(f"\n[完成] 结果已保存至: {output_file}")
            # messagebox.showinfo("完成", f"处理完成!\n结果已保存至:\n{output_file}") # 移除此处的 messagebox

        else:
            logger.warning("工作簿对象 (wb) 未创建或未包含任何数据，没有文件被保存。")
    except Exception as e_save:
        logger.error(f"保存工作簿时发生错误: {e_save}")
        print(f"[错误] 保存工作簿时出错: {e_save}")

def main():
    """主程序入口"""
    setup_environment()
    # GUI 元素（如messagebox）可能需要一个 Tk 实例，尽管 select_folder 会创建自己的。
    # 为了更稳健，可以在程序开始时创建一个隐藏的根窗口，并在结束时销毁。
    # 但对于当前的简单文件对话框，select_folder 内部的处理通常足够。

    print_header()

    print("\\n[提示] 准备通过图形界面选择数据文件夹...")
    time.sleep(0.1)
    # 选择文件夹
    folder_path = select_folder()
    
    if not folder_path:
        logger.warning("未选择文件夹，程序退出")
        time.sleep(0.2)  # 减少延迟
        return

    # 处理所有数据
    process_all_data(folder_path)
    
    logger.info("程序执行完毕")
    print("\n感谢使用电化学数据处理工具！")
    time.sleep(0.2)  # 减少延迟

if __name__ == "__main__":
    main()