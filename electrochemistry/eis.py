"""
电化学阻抗谱(EIS)数据处理模块
"""
import os
import sys
import numpy as np
import logging
from typing import Tuple, List, Optional, Dict, Any
from datetime import datetime

try:
    from tqdm import tqdm
    TQDM_AVAILABLE = True
except ImportError:
    TQDM_AVAILABLE = False
    print("提示：安装tqdm包可以显示进度条。可以运行：pip install tqdm")

# 导入共享模块
try:
    from .common import file_utils, excel_utils
except ImportError:
    # 如果作为独立模块运行
    try:
        parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        if parent_dir not in sys.path:
            sys.path.insert(0, parent_dir)
        from electrochemistry.common import file_utils, excel_utils
    except ImportError:
        print("无法导入共享模块，请确保目录结构正确")
        raise

# 检测是否在 GUI 环境中运行（如 PyInstaller 打包后的应用）
def is_gui_mode():
    """判断是否在无控制台的GUI环境中运行"""
    # 检查是否通过 PyInstaller 打包
    if getattr(sys, 'frozen', False):
        return True
    # 检查是否有可用的控制台 
    try:
        # 尝试向标准输出写入，如果失败则可能在GUI模式中
        sys.stdout.write("")
        return False
    except (AttributeError, IOError):
        return True

logger = logging.getLogger(__name__)

# 在GUI模式下禁用tqdm
if is_gui_mode():
    TQDM_AVAILABLE = False
    logger.info("检测到GUI模式，已禁用进度条显示")

def generate_zview_file(original_file_path: str, output_zview_file_path: str):
    """
    生成ZView兼容的文件。
    移除所有表头、注释，仅保留从第1行开始的数字数据。
    特定的表头 "Freq/Hz, Z'/ohm, Z\"/ohm" (及其变体) 用作标记
    以查找数据块的起始位置；此表头本身将被排除。
    
    参数:
        original_file_path: 原始EIS数据文件路径
        output_zview_file_path: 输出ZView兼容文件路径，格式为"{原文件名}-for ZView.txt"
    """
    try:
        logger.info(f"生成ZView兼容文件: {os.path.basename(output_zview_file_path)}")
        with open(original_file_path, 'r', errors='ignore') as f_in:
            lines = f_in.readlines()

        data_lines_for_zview = []
        data_section_identified = False # 一旦到达表头 *之后* 的行，则为 True
        possible_headers = [
            "Freq/Hz, Z'/ohm, Z\"/ohm",
            "Frequency/Hz,Z'/ohm,Z\"/ohm",
            "Freq,Z',Z\"",
            "Frequency,Zreal,Zimag",
            "f/Hz,Z'/Ω,Z\"/Ω"
        ]

        for line in lines:
            stripped_line = line.strip()

            if not data_section_identified:
                # 查找表头行。一旦找到，数据从下一个相关行开始
                for header in possible_headers:
                    if header in stripped_line:
                        data_section_identified = True
                        logger.debug(f"找到数据表头: {stripped_line}")
                        break
                # 始终继续到下一行，有效地跳过表头前行和表头行本身
                continue 
            
            # 如果 data_section_identified 为 true，我们现在正在处理应该是数据的行
            if not stripped_line or stripped_line.startswith('//') or stripped_line.startswith('#'): 
                # 跳过数据部分中的注释和空行
                continue            # 只要找到数据部分，就直接保留所有行，不做任何处理
            # 但仍需验证数据块的开始（第一行数据至少有3列，且前3列能转换为数字）
            parts = stripped_line.split(',')
            if len(parts) >= 3: # 确保至少有频率、实部和虚部三个数据
                try:
                    # 仅验证第一行数据的前三列是否为数字，用于识别数据段的开始
                    float(parts[0].strip())
                    float(parts[1].strip())
                    float(parts[2].strip())
                    
                    # 完全保留原始行，不做任何处理
                    data_lines_for_zview.append(stripped_line)
                except ValueError:
                    # 如果已识别数据块中的某一行不是数字，则假定它是数据的末尾
                    logger.debug(f"停止ZView数据采集，在非数据行: {stripped_line}")
                    break
            else: 
                # 如果行中的列数少于3，跳过该行
                continue
                
        if data_lines_for_zview:
            # 确保输出目录存在
            os.makedirs(os.path.dirname(output_zview_file_path), exist_ok=True)
            # 写入ZView兼容文件
            with open(output_zview_file_path, 'w', encoding='utf-8') as f_out:
                # 写入数据行
                for data_line in data_lines_for_zview:
                    f_out.write(data_line + "\n")
            logger.info(f"成功生成ZView兼容文件: {os.path.basename(output_zview_file_path)}")
            logger.info(f"提取了 {len(data_lines_for_zview)} 行数据")
        else:
            logger.warning(f"未找到数据或表头，无法生成ZView文件: {os.path.basename(original_file_path)}")

    except Exception as e:
        logger.error(f"生成ZView文件时出错 {os.path.basename(original_file_path)}: {e}")


def extract_eis_data(filename: str) -> Tuple[List[float], List[float], List[float]]:
    """
    从EIS数据文件中提取频率、实部(Z')和虚部(Z'')数据。
    数据列应为：Freq/Hz, Z'/ohm, Z\"/ohm
    
    参数:
        filename: 文件路径
        
    返回:
        (频率列表, 实部列表, 虚部列表)
    """
    frequencies = []
    z_real = []
    z_imaginary = []
    
    try:
        with open(filename, 'r', errors='ignore') as f:
            lines = f.readlines()
        
        data_started = False
        for line in lines:
            line = line.strip()
            # if not line or line.startswith('//'): # 跳过空行和注释行
            if not line or line.startswith('//'): # Skip empty lines and comment lines
                continue

            # if "Freq/Hz, Z'/ohm, Z\"/ohm" in line: # 数据列标题
            if "Freq/Hz, Z'/ohm, Z\"/ohm" in line: # Data column header
                data_started = True
                logger.info(f"在文件 {os.path.basename(filename)} 中找到数据列标题: {line}")
                continue
            
            if data_started:
                parts = line.split(',')
                # if len(parts) >= 3: # 确保至少有三列数据
                if len(parts) >= 3: # Ensure at least three columns of data
                    try:
                        freq = float(parts[0].strip())
                        z_r = float(parts[1].strip())
                        z_i = float(parts[2].strip())
                        
                        frequencies.append(freq)
                        z_real.append(z_r)
                        z_imaginary.append(z_i)
                    except ValueError:
                        logger.warning(f"无法解析数据行: {line} in {os.path.basename(filename)}。跳过此行。")
                        continue

        if not data_started:
            logger.warning(f"在文件 {os.path.basename(filename)} 中未找到期望的数据列标题 'Freq/Hz, Z'/ohm, Z\"/ohm'。")
        elif not frequencies:
            logger.warning(f"虽然找到了数据列标题，但在文件 {os.path.basename(filename)} 中未能提取任何有效数据点。")
        else:
            logger.info(f"从文件 {os.path.basename(filename)} 中成功提取 {len(frequencies)} 个数据点。")
            
    except Exception as e:
        logger.error(f"提取EIS数据时发生错误 {filename}: {e}")
        # return [], [], [] # 发生错误时返回空列表
        return [], [], [] # Return empty lists on error
        
    return frequencies, z_real, z_imaginary

def process_eis_files(file_paths: List[str], output_file: str, folder_basename: str, original_folder_path: str, wb: Optional[excel_utils.openpyxl.Workbook] = None) -> Optional[excel_utils.openpyxl.Workbook]:
    """
    处理EIS文件，提取Z'和Z''，将Z''*-1后与Z'一同保存到Excel报告中。
    同时生成ZView兼容的txt文件。

    参数:
        file_paths: EIS文件路径列表
        output_file: 输出Excel文件路径
        folder_basename: 文件夹基础名称 (用于表头或日志)
        original_folder_path: 用户最初选择的文件夹路径，用于保存ZView文件
        wb: 可选，现有的工作簿对象
        
    返回:
        处理成功则返回工作簿对象，否则返回None
    """
    if not file_paths:
        logger.info("没有EIS文件需要处理。")
        return wb if wb else None

    all_eis_data_processed = []
    max_length = 0

    file_iterator = tqdm(file_paths, desc="处理EIS文件") if TQDM_AVAILABLE else file_paths
    logger.info(f"开始处理 {len(file_paths)} 个EIS文件...")

    for file_path in file_iterator:
        try:            # 首先生成与 ZView 兼容的文件
            original_filename_no_ext = os.path.splitext(os.path.basename(file_path))[0]
            # 按照要求修改文件名格式为 {原文件名}-for ZView.txt
            zview_filename = f"{original_filename_no_ext}-for ZView.txt"

            # ZView 文件将保存在用户选择的 original_folder_path 中
            zview_output_path = os.path.join(original_folder_path, zview_filename)
            generate_zview_file(file_path, zview_output_path)

            # _, z_real, z_imaginary = extract_eis_data(file_path) # Continues to use original file for Excel
            _, z_real, z_imaginary = extract_eis_data(file_path) # 继续使用原始文件进行 Excel 操作

            if not z_real or not z_imaginary:
                logger.warning(f"从文件 {os.path.basename(file_path)} 中未提取到有效的Z'或Z''数据，跳过。")
                continue

            processed_z_imaginary = [-val for val in z_imaginary]
            file_id = os.path.splitext(os.path.basename(file_path))[0]
            
            all_eis_data_processed.append({
                "file_id": file_id,
                "z_prime": z_real,
                "minus_z_double_prime": processed_z_imaginary
            })
            max_length = max(max_length, len(z_real))
            logger.info(f"已处理文件: {file_id}, 数据点: {len(z_real)}")

        except Exception as e:
            logger.error(f"处理EIS文件 {os.path.basename(file_path)} 时发生错误: {e}")
            continue
    
    if not all_eis_data_processed:
        logger.warning("未能从任何EIS文件中成功提取和处理数据。")
        return wb

    # 使用提供的工作簿或创建新的
    # # new_workbook_created_internally = False # No longer needed
    # new_workbook_created_internally = False # 不再需要
    if wb is None:
        wb, ws, header_fill, thin_border, center_aligned, openpyxl_module = excel_utils.setup_excel_workbook("EIS Data")
        logger.info("为EIS数据创建了新的工作簿和工作表 'EIS Data'")
        # # new_workbook_created_internally = True
        # new_workbook_created_internally = True
    else:
        header_fill, thin_border, center_aligned, openpyxl_module = excel_utils.get_excel_styles()
        worksheet_name = "EIS Data"
        if worksheet_name in wb.sheetnames:
            ws = wb[worksheet_name]
            logger.info(f"使用现有的工作表 '{worksheet_name}' 处理EIS数据")
        else:
            ws = wb.create_sheet(worksheet_name)
            logger.info(f"已创建新的工作表 '{worksheet_name}' 用于EIS数据")

    # # PatternFill = openpyxl_module.styles.PatternFill # Not needed if yellow_fill is defined locally or via utility
    # # Font = openpyxl_module.styles.Font # Will use excel_utils.get_bold_font
    # # Alignment = openpyxl_module.styles.Alignment # Will use openpyxl_module.styles.Alignment directly
    # PatternFill = openpyxl_module.styles.PatternFill # 如果 yellow_fill 是本地定义或通过实用程序定义，则不需要
    # Font = openpyxl_module.styles.Font # 将使用 excel_utils.get_bold_font
    # Alignment = openpyxl_module.styles.Alignment # 将直接使用 openpyxl_module.styles.Alignment
    
    # # Define yellow_fill locally if not part of a broader utility yet
    # 如果 yellow_fill 尚不属于更广泛的实用程序的一部分，则在本地定义它
    yellow_fill = openpyxl_module.styles.PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    # bold_font = excel_utils.get_bold_font() # Use utility
    bold_font = excel_utils.get_bold_font() # 使用实用程序

    rs_data_to_write = []

    current_col = 1

    for data_item in all_eis_data_processed:
        file_id = data_item["file_id"]
        z_prime_list = data_item["z_prime"]
        minus_z_double_prime_list = data_item["minus_z_double_prime"]

        ws.cell(row=1, column=current_col).value = "Z′"
        ws.cell(row=1, column=current_col + 1).value = "-Z″"
        ws.cell(row=2, column=current_col).value = "Ohm"
        ws.cell(row=2, column=current_col + 1).value = "Ohm"
        ws.cell(row=3, column=current_col).value = ""
        ws.cell(row=3, column=current_col + 1).value = file_id

        for r in range(1, 4):
            for c_offset in range(2):
                cell = ws.cell(row=r, column=current_col + c_offset)
                cell.fill = header_fill
                cell.border = thin_border
                # cell.alignment = center_aligned # center_aligned is already an Alignment object from excel_utils
                cell.alignment = center_aligned # center_aligned 已经是来自 excel_utils 的 Alignment 对象
                # if bold_font: # Apply bold font to data headers
                if bold_font: # 将粗体字体应用于数据表头
                    cell.font = bold_font

        for c_offset in range(2):
            cell_r4 = ws.cell(row=4, column=current_col + c_offset)
            cell_r4.value = None
            cell_r4.fill = header_fill
            cell_r4.border = thin_border
            cell_r4.alignment = center_aligned

        crossover_indices = []
        solution_resistance = None

        for i in range(len(minus_z_double_prime_list) - 1):
            current_mzd = minus_z_double_prime_list[i]
            next_mzd = minus_z_double_prime_list[i + 1]

            if current_mzd > 0 and next_mzd <= 0:
                crossover_indices = [i, i + 1]
                break
            elif current_mzd < 0 and next_mzd >= 0 and not crossover_indices:
                crossover_indices = [i, i + 1]
                break
            elif current_mzd > 0 and next_mzd == 0:
                crossover_indices = [i, i + 1]
                break
            elif current_mzd < 0 and next_mzd == 0 and not crossover_indices:
                crossover_indices = [i, i + 1]
                break

        if not crossover_indices and minus_z_double_prime_list:
            min_pos_mzdp_val = float('inf')
            min_pos_idx = -1
            for i, val in enumerate(minus_z_double_prime_list):
                if val > 0 and val < min_pos_mzdp_val:
                    min_pos_mzdp_val = val
                    min_pos_idx = i

            if min_pos_idx != -1:
                if min_pos_idx + 1 < len(z_prime_list):
                    crossover_indices = [min_pos_idx, min_pos_idx + 1]
                elif min_pos_idx > 0:
                    crossover_indices = [min_pos_idx - 1, min_pos_idx]
                else:
                    crossover_indices = [min_pos_idx]
            else:
                max_neg_mzdp_val = -float('inf')
                max_neg_idx = -1
                for i, val in enumerate(minus_z_double_prime_list):
                    if val <= 0 and val > max_neg_mzdp_val:
                        max_neg_mzdp_val = val
                        max_neg_idx = i
                if max_neg_idx != -1:
                    if max_neg_idx + 1 < len(z_prime_list):
                        crossover_indices = [max_neg_idx, max_neg_idx + 1]
                    elif max_neg_idx > 0:
                        crossover_indices = [max_neg_idx - 1, max_neg_idx]
                    else:
                        crossover_indices = [max_neg_idx]

        if crossover_indices:
            if len(crossover_indices) == 2:
                idx1, idx2 = crossover_indices
                z_prime_at_crossover1 = z_prime_list[idx1]
                z_prime_at_crossover2 = z_prime_list[idx2]
                solution_resistance = (z_prime_at_crossover1 + z_prime_at_crossover2) / 2.0
                logger.info(f"File {file_id}: Crossover for -Z'' found at indices {idx1}, {idx2}. "
                            f"Z' values: {z_prime_at_crossover1}, {z_prime_at_crossover2}. Rs = {solution_resistance:.4f}")
            elif len(crossover_indices) == 1:
                idx1 = crossover_indices[0]
                solution_resistance = z_prime_list[idx1]
                logger.info(f"File {file_id}: Crossover for -Z'' determined by single point at index {idx1}. "
                            f"Z' value: {solution_resistance}. Rs = {solution_resistance:.4f}")

        for row_idx_offset, (zp_val, m_zd_val) in enumerate(zip(z_prime_list, minus_z_double_prime_list)):
            actual_row_idx = row_idx_offset + 5
            cell_zp = ws.cell(row=actual_row_idx, column=current_col, value=zp_val)
            cell_m_zd = ws.cell(row=actual_row_idx, column=current_col + 1, value=m_zd_val)
            cell_zp.border = thin_border
            cell_m_zd.border = thin_border

            # # Removed bolding and fill for crossover points
            # # if crossover_indices and row_idx_offset in crossover_indices:
            # #     cell_zp.fill = yellow_fill
            # #     cell_zp.font = bold_font
            # #     cell_m_zd.fill = yellow_fill
            # #     cell_m_zd.font = bold_font
            # 移除了交叉点的加粗和填充
            # if crossover_indices and row_idx_offset in crossover_indices:
            #     cell_zp.fill = yellow_fill
            #     cell_zp.font = bold_font
            #     cell_m_zd.fill = yellow_fill
            #     cell_m_zd.font = bold_font

        if solution_resistance is not None:
            rs_data_to_write.append({
                "file_id": file_id,
                "rs_value": solution_resistance,
                "insert_col": current_col + 2
            })

        current_col += 2

    rs_section_start_col = current_col + 1
    if rs_data_to_write:
        # # Style for "分析结果" main header
        # “分析结果”主表头的样式
        analysis_header_cell = ws.cell(row=1, column=rs_section_start_col)
        # analysis_header_cell.value = "Analysis Results" # Changed to English
        analysis_header_cell.value = "Analysis Results" # 已更改为英文
        if bold_font:
            analysis_header_cell.font = bold_font
        # analysis_header_cell.alignment = center_aligned # Use center_aligned from excel_utils
        analysis_header_cell.alignment = center_aligned # 使用来自 excel_utils 的 center_aligned
        analysis_header_cell.fill = header_fill 
        analysis_header_cell.border = thin_border
        
        ws.merge_cells(start_row=1, start_column=rs_section_start_col, end_row=1, end_column=rs_section_start_col + 1)
        
        if rs_section_start_col + 1 <= ws.max_column:
            merged_cell_part2 = ws.cell(row=1, column=rs_section_start_col + 1)
            merged_cell_part2.fill = header_fill 
            merged_cell_part2.border = thin_border
            # if bold_font: # Ensure merged part is also bold
            if bold_font: # 确保合并部分也为粗体
                merged_cell_part2.font = bold_font


        header_row_for_rs = 2
        # # Revert to single cell for "Solution Resistance (Rs) / Ohm"
        # 恢复为单个单元格以显示“溶液电阻 (Rs) / Ohm”
        # ws.cell(row=header_row_for_rs, column=rs_section_start_col).value = "File ID" # Changed to English
        ws.cell(row=header_row_for_rs, column=rs_section_start_col).value = "File ID" # 已更改为英文
        # ws.cell(row=header_row_for_rs, column=rs_section_start_col + 1).value = "Solution Resistance (Rs) / Ohm" # Changed to English
        ws.cell(row=header_row_for_rs, column=rs_section_start_col + 1).value = "Solution Resistance (Rs) / Ohm" # 已更改为英文

        # # Adjust header styling loop for the single cell structure
        # 调整单个单元格结构的表头样式循环
        # for r_offset, col_title in enumerate(["File ID", "Solution Resistance (Rs) / Ohm"]): # Changed to English
        for r_offset, col_title in enumerate(["File ID", "Solution Resistance (Rs) / Ohm"]): # 已更改为英文
            cell = ws.cell(row=header_row_for_rs, column=rs_section_start_col + r_offset)
            cell.fill = header_fill
            cell.border = thin_border
            if bold_font:
                cell.font = bold_font
            # cell.alignment = center_aligned # Use center_aligned from excel_utils
            cell.alignment = center_aligned # 使用来自 excel_utils 的 center_aligned
            # # ws.column_dimensions[openpyxl_module.utils.get_column_letter(rs_section_start_col + r_offset)].width = 30 # Increased width
            # ws.column_dimensions[openpyxl_module.utils.get_column_letter(rs_section_start_col + r_offset)].width = 30 # 增加宽度
        # # Use utility for column widths
        # 使用实用程序设置列宽
        excel_utils.set_column_widths(ws, {
            openpyxl_module.utils.get_column_letter(rs_section_start_col): 30,
            openpyxl_module.utils.get_column_letter(rs_section_start_col + 1): 30
        })

        # # Remove styling for the now non-existent "Ohm" cell
        # # cell_ohm = ws.cell(row=header_row_for_rs + 1, column=rs_section_start_col + 1)
        # # cell_ohm.fill = header_fill
        # # cell_ohm.border = thin_border
        # # cell_ohm.font = bold_font
        # # cell_ohm.alignment = center_aligned
        # 移除现在不存在的“Ohm”单元格的样式
        # cell_ohm = ws.cell(row=header_row_for_rs + 1, column=rs_section_start_col + 1)
        # cell_ohm.fill = header_fill
        # cell_ohm.border = thin_border
        # cell_ohm.font = bold_font
        # cell_ohm.alignment = center_aligned

        # # Adjust data writing loop for the single cell header structure
        # 调整单个单元格表头结构的数据写入循环
        for idx, rs_item in enumerate(rs_data_to_write):
            # data_row = header_row_for_rs + 1 + idx # Adjusted start row for data
            data_row = header_row_for_rs + 1 + idx # 调整了数据的起始行

            cell_file_id = ws.cell(row=data_row, column=rs_section_start_col, value=rs_item["file_id"])
            cell_rs_val = ws.cell(row=data_row, column=rs_section_start_col + 1, value=rs_item['rs_value'])
            cell_rs_val.number_format = '0.0000'

            for c_offset in range(2):
                cell = ws.cell(row=data_row, column=rs_section_start_col + c_offset)
                cell.border = thin_border
                # # Create Alignment object directly if not using a predefined one like center_aligned
                # 如果不使用像 center_aligned 这样的预定义对象，则直接创建 Alignment 对象
                cell.alignment = openpyxl_module.styles.Alignment(horizontal='left') 

    # # for col_idx in range(1, current_col):
    # #     col_letter = openpyxl_module.utils.get_column_letter(col_idx)
    # #     ws.column_dimensions[col_letter].width = 18
    # # Use utility for column widths
    # 使用实用程序设置列宽
    col_width_map_data = {}
    for col_idx in range(1, current_col):
        col_width_map_data[openpyxl_module.utils.get_column_letter(col_idx)] = 18
    excel_utils.set_column_widths(ws, col_width_map_data)

    # # 保存工作簿 - Removed, main module will save.
    # # if output_file and new_workbook_created_internally: 
    # #     try:
    # #         wb.save(output_file)
    # #         logger.info(f"EIS数据已成功保存到 {output_file}")
    # #     except PermissionError:
    # #         logger.error(f"保存文件 {output_file} 失败：权限不足或文件可能已被其他程序打开。")
    # #         print(f"错误: 无法保存文件 {output_file}。请确保文件未被打开并具有写入权限。")
    # #         # return None # Don't return None, wb might be valid. Main module saves.
    # #     except Exception as e:
    # #         logger.error(f"保存EIS数据到 {output_file} 时发生未知错误: {e}")
    #         # return None # Don't return None, wb might be valid. Main module saves
    # 保存工作簿 - 已移除，主模块将保存。
    # if output_file and new_workbook_created_internally: 
    #     try:
    #         wb.save(output_file)
    #         logger.info(f"EIS数据已成功保存到 {output_file}")
    #     except PermissionError:
    #         logger.error(f"保存文件 {output_file} 失败：权限不足或文件可能已被其他程序打开。")
    #         print(f"错误: 无法保存文件 {output_file}。请确保文件未被打开并具有写入权限。")
    #         # return None # 不要返回 None，wb 可能有效。主模块保存。
    #     except Exception as e:
    #         logger.error(f"保存EIS数据到 {output_file} 时发生未知错误: {e}")
            # return None # 不要返回 None，wb 可能有效。主模块保存
            
    # # Prepare data for returning
    # 准备要返回的数据
    returned_rs_data = [{'file_id': item['file_id'], 'rs': item['rs_value']} for item in rs_data_to_write]
    return wb, returned_rs_data

# # process_all_files_from_paths is an alias for process_eis_files
# # The signature of process_eis_files is (file_paths, output_file, folder_basename, original_folder_path, wb)
# # It now returns Tuple[Workbook, List[Dict[str, Any]]]
# process_all_files_from_paths 是 process_eis_files 的别名
# process_eis_files 的签名为 (file_paths, output_file, folder_basename, original_folder_path, wb)
# 它现在返回 Tuple[Workbook, List[Dict[str, Any]]]
process_all_files_from_paths = process_eis_files

def find_eis_files(folder_path: str) -> List[str]:
    """
    在指定文件夹中查找所有可能的EIS数据文件
    通过检查文件内容中是否包含 "A.C. Impedance" 来判断是否为EIS文件。
    """
    eis_files_list = []
    if not os.path.isdir(folder_path):
        logger.warning(f"提供的路径 {folder_path} 不是一个有效的文件夹。")
        return eis_files_list

    txt_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path)
                 if f.lower().endswith('.txt') and os.path.isfile(os.path.join(folder_path, f))]

    if not txt_files:
        logger.info(f"在 {folder_path} 中未找到任何 .txt 文件。")
        return eis_files_list

    logger.info(f"在 {folder_path} 找到 {len(txt_files)} 个 .txt 文件，正在检查EIS标识...")

    file_iterator = tqdm(txt_files, desc="识别EIS数据文件") if TQDM_AVAILABLE else txt_files
    for file_path in file_iterator:
        try:
            with open(file_path, 'r', errors='ignore') as f:
                header_content = f.read(1024)
                if "A.C. Impedance" in header_content:
                    eis_files_list.append(file_path)
                    logger.info(f"文件 {os.path.basename(file_path)} 被识别为EIS数据文件。")
        except Exception as e:
            logger.warning(f"检查文件 {file_path} 时发生错误: {e}")
            
    if not eis_files_list:
        logger.info(f"在 {folder_path} 中未识别到任何EIS数据文件。")
    else:
        logger.info(f"在 {folder_path} 中共识别到 {len(eis_files_list)} 个EIS数据文件。")
        
    return eis_files_list

def main():
    """EIS数据处理主函数"""
    logger.info("EIS数据处理程序启动")
    
    folder_path = file_utils.select_folder()
    
    if not folder_path:
        logger.warning("未选择文件夹，程序退出。")
        return
    
    logger.info(f"已选择文件夹: {folder_path}")
    
    eis_files = find_eis_files(folder_path)
    
    if not eis_files:
        logger.error(f"在选定的文件夹中未找到任何EIS数据文件。")
        print("\n在选定的文件夹中未找到任何EIS数据文件。请确保文件内容包含电化学阻抗谱相关信息。")
        return
    
    logger.info(f"找到 {len(eis_files)} 个EIS数据文件:")
    for file_path in eis_files:
        logger.info(f"  {os.path.basename(file_path)}")
    
    output_dir, folder_basename = file_utils.ensure_output_dir(folder_path)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = os.path.join(output_dir, f"{folder_basename}_processed_eis_data_{timestamp}.xlsx")
    
    try:
        wb = process_eis_files(eis_files, output_file, folder_basename, folder_path)
        if wb:
            print(f"\nEIS数据已成功处理并保存到 {output_file}")
        else:
            print("\nEIS数据处理失败或未找到有效数据。")
        
    except Exception as e:
        logger.error(f"发生意外错误: {str(e)}")
        print(f"\n发生意外错误: {str(e)}")
        print("\n建议解决问题的方法:")
        print("1. 确保安装了所有必需的包")
        print("2. 检查输入文件是否存在并具有预期格式")
    
    logger.info("程序执行完毕")
    input("\n按Enter键返回主菜单...")

if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    main()