from flask import Flask, request, jsonify, send_from_directory
import threading
import dropbox
import os
import re
from copy import copy
from openpyxl.styles import Alignment, Font, Border, Side, Color
from openpyxl.styles.numbers import FORMAT_NUMBER_COMMA_SEPARATED1
import gc # Import garbage collector
import csv # Import CSV module
from excel_processing_utils import _create_excel_from_csv_files # Import the new helper function

print("=== Flask API 啟動 ===")

app = Flask(__name__)

status = {"running": False, "result": None, "progress": "尚未開始"}

# 將原本的主流程包成一個函數
def main_job():
    global status, config # Add config here if it's read globally and needed
    status["running"] = True
    status["result"] = "處理中...請稍候..."
    status["progress"] = "初始化中..."
    print("=== main_job 啟動 ===")
    # --- Ensure all necessary imports are here, including dropbox if not already top-level ---
    import os
    import json
    import ddddocr
    from PIL import Image
    from io import BytesIO
    import time
    from bs4 import BeautifulSoup
    import requests
    import pandas as pd
    import openpyxl
    from openpyxl import Workbook
    from openpyxl.utils.dataframe import dataframe_to_rows
    from collections import defaultdict
    from datetime import datetime
    import threading
    import glob
    from concurrent.futures import ThreadPoolExecutor, as_completed
    import random
    from copy import copy
    from openpyxl.styles import Alignment, Font, Border, Side, Color
    from openpyxl.styles.numbers import FORMAT_NUMBER_COMMA_SEPARATED1
    import dropbox # Ensure dropbox is imported

    # --- Re-read or ensure config is loaded with new parameters ---
    # This might be redundant if config is already a global dictionary loaded at startup.
    # For safety, let's ensure the required keys are accessible.
    # The global `config` dictionary should be populated by `load_config()` or similar at app start.
    
    # Essential Dropbox configuration from global config
    current_dropbox_token = config.get('dropbox_token', '')
    dropbox_account_file_path_config = config.get('dropbox_account_file_path', '/Apps/ExcelAPI-app/account/account.txt') # Default if not in config, but should be

    result_log = [] # Initialize result_log

    # --- NEW: Define output_dir earlier in main_job scope ---
    folder_name = datetime.now().strftime('%Y%m%d_%H%M')
    output_dir = os.path.join('資料夾路徑', folder_name) 
    try:
        os.makedirs(output_dir, exist_ok=True)
        result_log.append(f"輸出目錄已準備: {output_dir}")
    except Exception as e_mkdir:
        print(f"CRITICAL: 無法創建輸出目錄 {output_dir}: {e_mkdir}")
        status["result"] = f"錯誤: 無法創建輸出目錄 {output_dir}: {e_mkdir}"
        status["progress"] = "目錄創建失敗"
        status["running"] = False
        return
    # --- End NEW ---
    
    headers_for_csv_and_excel = [
        "獎金周期", "獎金周期", "消費對等", "經營分紅", "安置獎金", "推薦獎金",
        "消費分紅", "經營對等", "收件中心", "新增加權", "小計", "其他加項",
        "其他減項", "稅額", "補充費", "總計", "紅利積分", "電子錢包",
        "獎金暫存", "註冊分", "商品券", "星級", "左區人數", "右區人數"
    ]

    try:
        ocr_instance = ddddocr.DdddOcr(show_ad=False)
        print("ddddocr.DdddOcr 實例已在 main_job 中成功初始化。")
    except Exception as e_ocr_init:
        print(f"錯誤：在 main_job 中初始化 ddddocr.DdddOcr 失敗: {e_ocr_init}")
        status["result"] = f"錯誤：OCR組件初始化失敗: {e_ocr_init}"
        status["progress"] = "OCR 初始化失敗"
        status["running"] = False
        return

    # --- Start of functions copied and adapted from Auto.py (HELPER FUNCTIONS) ---
    # ... (is_number_value, apply_formatting_to_cell, copy_cell_format_for_api, sort_sheets_by_gold_level_in_api) ...
    # ... (_internal_generate_bonus2_report, _internal_split_bonus2_sheets) ...
    # ... (get_random_ua, DEBUG definition) ...
    # These helper functions remain the same.
    # Make sure global_color_map_for_reports and global_thin_border_for_reports are defined if used by them.
    global_color_map_for_reports = {
        "紅利積分": "FF0000", "電子錢包": "00008B", "獎金暫存": "8B4513",
        "註冊分": "FF8C00", "商品券": "2F4F4F", "星級": "708090"
    }
    global_thin_border_for_reports = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    def is_number_value(value):
        if value is None: return False
        try: float(str(value).replace(',', '')); return True
        except (ValueError, TypeError): return False

    def apply_formatting_to_cell(cell, bold=False, font_color_hex=None, border=None, alignment_horizontal='center', alignment_vertical='center'):
        if border: cell.border = border
        cell.alignment = Alignment(horizontal=alignment_horizontal, vertical=alignment_vertical)
        current_font = cell.font if cell.has_style and cell.font else Font()
        new_font_attributes = {'name': current_font.name, 'sz': current_font.sz if current_font.sz else 11,'b': bold if bold is not None else current_font.b,'i': current_font.i,'vertAlign': current_font.vertAlign,'underline': current_font.underline,'strike': current_font.strike,}
        if font_color_hex: new_font_attributes['color'] = Color(rgb=font_color_hex)
        elif current_font.color: new_font_attributes['color'] = current_font.color
        cell.font = Font(**new_font_attributes)

    def copy_cell_format_for_api(source_cell, target_cell):
        if source_cell.has_style:
            target_cell.font = copy(source_cell.font)
            target_cell.border = copy(source_cell.border)
            target_cell.fill = copy(source_cell.fill)
            target_cell.number_format = source_cell.number_format
            target_cell.protection = copy(source_cell.protection)
            target_cell.alignment = copy(source_cell.alignment)

    def sort_sheets_by_gold_level_in_api(sheet_names_list, workbook_source):
        def get_sheet_order_key(sheet_name_str):
            ws = workbook_source[sheet_name_str]
            star_level_val = ws['V2'].value
            return 1 if star_level_val and "金級" in str(star_level_val) else 0
        return sorted(sheet_names_list, key=get_sheet_order_key)

    def _internal_generate_bonus2_report(source_bonus_xlsx_path, output_bonus2_xlsx_path):
        # ... (This function's internal logic remains the same) ...
        # Make sure it uses `result_log.append` for logging
        result_log.append(f"內部函數：開始生成 Bonus2.xlsx 從 {source_bonus_xlsx_path}")
        try:
            if not os.path.exists(source_bonus_xlsx_path):
                result_log.append(f"❌ 錯誤: 來源 bonus.xlsx '{source_bonus_xlsx_path}' 不存在。")
                return False
            wb_source = openpyxl.load_workbook(source_bonus_xlsx_path, data_only=True)
            # ... (rest of the function) ...
            # Ensure it has its own try-except and memory optimization
            wb_target = openpyxl.Workbook()
            if 'Sheet' in wb_target.sheetnames: wb_target.remove(wb_target.active)
            person_sheets_map = defaultdict(list)
            for sheet_name_from_bonus in wb_source.sheetnames:
                name_raw_part = sheet_name_from_bonus.split("_")[0]
                person_identifier = re.sub(r'\\d+', '', name_raw_part)
                person_sheets_map[person_identifier].append(sheet_name_from_bonus)
            all_dates_globally = set()
            for _, source_sheet_names in person_sheets_map.items():
                for s_name in source_sheet_names:
                    ws_s = wb_source[s_name]
                    dates_in_sheet = [row[0] for row in ws_s.iter_rows(min_row=2, min_col=1, max_col=1, values_only=True) if row[0] is not None]
                    all_dates_globally.update(dates_in_sheet)
            sorted_all_dates_desc = sorted(list(all_dates_globally), reverse=True)
            for person_id, source_sheet_names_for_person in person_sheets_map.items():
                if not source_sheet_names_for_person: continue
                sorted_source_sheets_for_person = sort_sheets_by_gold_level_in_api(source_sheet_names_for_person, wb_source)
                ws_target = wb_target.create_sheet(title=person_id[:31])
                ws_target['A1'] = "名稱"; apply_formatting_to_cell(ws_target['A1'], bold=True, border=global_thin_border_for_reports)
                for col_idx_acc, sheet_name_src in enumerate(sorted_source_sheets_for_person):
                    name_part = sheet_name_src.split('_')[0]
                    target_cell_name = ws_target.cell(row=1, column=2 + col_idx_acc, value=name_part)
                    apply_formatting_to_cell(target_cell_name, bold=True, border=global_thin_border_for_reports)
                titles_for_a2_a7 = ["紅利積分", "電子錢包", "獎金暫存", "註冊分", "商品券", "星級"]
                source_col_letters_for_a2_a7_data = ['Q', 'R', 'S', 'T', 'U', 'V']
                for row_offset, title_a_col in enumerate(titles_for_a2_a7):
                    current_row_bonus2 = 2 + row_offset
                    cell_a_title = ws_target.cell(row=current_row_bonus2, column=1, value=title_a_col)
                    font_clr = global_color_map_for_reports.get(title_a_col)
                    apply_formatting_to_cell(cell_a_title, bold=True, font_color_hex=font_clr, border=global_thin_border_for_reports)
                    for acc_col_idx, sheet_name_src in enumerate(sorted_source_sheets_for_person):
                        ws_src_current_acc = wb_source[sheet_name_src]
                        source_cell_value = ws_src_current_acc[f'{source_col_letters_for_a2_a7_data[row_offset]}2'].value
                        target_data_cell = ws_target.cell(row=current_row_bonus2, column=2 + acc_col_idx)
                        if is_number_value(source_cell_value): target_data_cell.value = float(str(source_cell_value).replace(',', '')); target_data_cell.number_format = FORMAT_NUMBER_COMMA_SEPARATED1
                        else: target_data_cell.value = source_cell_value
                        apply_formatting_to_cell(target_data_cell, bold=True, font_color_hex=font_clr, border=global_thin_border_for_reports)
                ws_target['A8'] = "帳號"; apply_formatting_to_cell(ws_target['A8'], bold=True, border=global_thin_border_for_reports)
                for col_idx_acc, sheet_name_src in enumerate(sorted_source_sheets_for_person):
                    acc_num_part = sheet_name_src.split('_', 1)[-1] if '_' in sheet_name_src else sheet_name_src
                    target_cell_acc_num = ws_target.cell(row=8, column=2 + col_idx_acc, value=acc_num_part)
                    apply_formatting_to_cell(target_cell_acc_num, border=global_thin_border_for_reports)
                ws_target['A9'] = "左右人數"; apply_formatting_to_cell(ws_target['A9'], bold=True, font_color_hex="006400", border=global_thin_border_for_reports)
                for col_idx_acc, sheet_name_src in enumerate(sorted_source_sheets_for_person):
                    ws_src_current_acc = wb_source[sheet_name_src]
                    left_count_val = ws_src_current_acc['W2'].value; right_count_val = ws_src_current_acc['X2'].value
                    lr_text = f"{left_count_val or 0} <> {right_count_val or 0}"
                    target_cell_lr = ws_target.cell(row=9, column=2 + col_idx_acc, value=lr_text)
                    apply_formatting_to_cell(target_cell_lr, font_color_hex="006400", border=global_thin_border_for_reports)
                ws_target['A10'] = "總計"; apply_formatting_to_cell(ws_target['A10'], bold=True, font_color_hex="8B008B", border=global_thin_border_for_reports)
                for date_row_idx, date_val in enumerate(sorted_all_dates_desc):
                    cell_date = ws_target.cell(row=11 + date_row_idx, column=1, value=date_val)
                    if isinstance(date_val, datetime): cell_date.number_format = 'YYYY/MM/DD'
                    apply_formatting_to_cell(cell_date, border=global_thin_border_for_reports)
                for acc_col_idx, sheet_name_src in enumerate(sorted_source_sheets_for_person):
                    ws_src_current_acc = wb_source[sheet_name_src]; date_to_m_column_value_map = {}
                    for src_row_data in ws_src_current_acc.iter_rows(min_row=2, max_col=13, values_only=True):
                        date_in_src_row = src_row_data[0]; m_column_value_in_src_row = src_row_data[12] if len(src_row_data) > 12 else None
                        if date_in_src_row is not None: date_to_m_column_value_map[date_in_src_row] = m_column_value_in_src_row
                    sum_for_this_account_col_10 = 0
                    for date_row_idx, date_val_target in enumerate(sorted_all_dates_desc):
                        value_for_date = date_to_m_column_value_map.get(date_val_target)
                        data_cell = ws_target.cell(row=11 + date_row_idx, column=2 + acc_col_idx)
                        if is_number_value(value_for_date):
                            numeric_value = float(str(value_for_date).replace(',', '')); data_cell.value = numeric_value
                            data_cell.number_format = FORMAT_NUMBER_COMMA_SEPARATED1; sum_for_this_account_col_10 += numeric_value
                        else: data_cell.value = value_for_date
                        apply_formatting_to_cell(data_cell, border=global_thin_border_for_reports)
                    target_cell_total_r10_calculated = ws_target.cell(row=10, column=2 + acc_col_idx)
                    target_cell_total_r10_calculated.value = sum_for_this_account_col_10; target_cell_total_r10_calculated.number_format = FORMAT_NUMBER_COMMA_SEPARATED1
                    apply_formatting_to_cell(target_cell_total_r10_calculated, bold=True, font_color_hex="8B008B", border=global_thin_border_for_reports)
                num_data_cols_for_person = len(sorted_source_sheets_for_person)
                usd_total_col_bonus2 = 2 + num_data_cols_for_person; twd_total_col_bonus2 = usd_total_col_bonus2 + 1
                ws_target.cell(row=9, column=usd_total_col_bonus2, value="美元收入").font = Font(color="8B008B", bold=True)
                apply_formatting_to_cell(ws_target.cell(row=9, column=usd_total_col_bonus2), border=global_thin_border_for_reports)
                ws_target.cell(row=9, column=twd_total_col_bonus2, value="台幣收入").font = Font(color="0000FF", bold=True)
                apply_formatting_to_cell(ws_target.cell(row=9, column=twd_total_col_bonus2), border=global_thin_border_for_reports)
                sum_usd_row10 = sum(ws_target.cell(row=10, column=2 + i).value or 0 for i in range(num_data_cols_for_person) if is_number_value(ws_target.cell(row=10, column=2+i).value))
                cell_usd_total_r10 = ws_target.cell(row=10, column=usd_total_col_bonus2, value=sum_usd_row10)
                cell_usd_total_r10.number_format = FORMAT_NUMBER_COMMA_SEPARATED1
                apply_formatting_to_cell(cell_usd_total_r10, bold=True, font_color_hex="8B008B", border=global_thin_border_for_reports)
                twd_val_r10 = sum_usd_row10 * 33
                cell_twd_total_r10 = ws_target.cell(row=10, column=twd_total_col_bonus2, value=twd_val_r10)
                cell_twd_total_r10.number_format = FORMAT_NUMBER_COMMA_SEPARATED1
                apply_formatting_to_cell(cell_twd_total_r10, bold=True, font_color_hex="0000FF", border=global_thin_border_for_reports)
                for date_row_idx_calc in range(len(sorted_all_dates_desc)):
                    current_data_row_bonus2 = 11 + date_row_idx_calc
                    sum_usd_for_date_row = sum(ws_target.cell(row=current_data_row_bonus2, column=2 + i).value or 0 for i in range(num_data_cols_for_person) if is_number_value(ws_target.cell(row=current_data_row_bonus2, column=2 + i).value))
                    cell_usd_date_row = ws_target.cell(row=current_data_row_bonus2, column=usd_total_col_bonus2, value=sum_usd_for_date_row)
                    cell_usd_date_row.number_format = FORMAT_NUMBER_COMMA_SEPARATED1
                    apply_formatting_to_cell(cell_usd_date_row, font_color_hex="8B008B", border=global_thin_border_for_reports)
                    twd_val_for_date_row = sum_usd_for_date_row * 33
                    cell_twd_date_row = ws_target.cell(row=current_data_row_bonus2, column=twd_total_col_bonus2, value=twd_val_for_date_row)
                    cell_twd_date_row.number_format = FORMAT_NUMBER_COMMA_SEPARATED1
                    apply_formatting_to_cell(cell_twd_date_row, font_color_hex="0000FF", border=global_thin_border_for_reports)
                sum_electronic_wallet = sum(ws_target.cell(row=3, column=2 + i).value or 0 for i in range(num_data_cols_for_person) if is_number_value(ws_target.cell(row=3, column=2 + i).value))
                cell_sum_ew = ws_target.cell(row=3, column=usd_total_col_bonus2, value=sum_electronic_wallet); cell_sum_ew.number_format = FORMAT_NUMBER_COMMA_SEPARATED1
                apply_formatting_to_cell(cell_sum_ew, bold=True, font_color_hex=global_color_map_for_reports.get("電子錢包"), border=global_thin_border_for_reports)
                ws_target.cell(row=3, column=twd_total_col_bonus2, value="←電子錢包總和").font = Font(color=global_color_map_for_reports.get("電子錢包"), bold=True)
                apply_formatting_to_cell(ws_target.cell(row=3, column=twd_total_col_bonus2), border=global_thin_border_for_reports, alignment_horizontal='left')
                sum_bonus_storage = sum(ws_target.cell(row=4, column=2 + i).value or 0 for i in range(num_data_cols_for_person) if is_number_value(ws_target.cell(row=4, column=2 + i).value))
                cell_sum_bs = ws_target.cell(row=4, column=usd_total_col_bonus2, value=sum_bonus_storage); cell_sum_bs.number_format = FORMAT_NUMBER_COMMA_SEPARATED1
                apply_formatting_to_cell(cell_sum_bs, bold=True, font_color_hex=global_color_map_for_reports.get("獎金暫存"), border=global_thin_border_for_reports)
                ws_target.cell(row=4, column=twd_total_col_bonus2, value="←獎金暫存總和").font = Font(color=global_color_map_for_reports.get("獎金暫存"), bold=True)
                apply_formatting_to_cell(ws_target.cell(row=4, column=twd_total_col_bonus2), border=global_thin_border_for_reports, alignment_horizontal='left')
                ws_target.column_dimensions['A'].width = 12
                for i in range(num_data_cols_for_person): ws_target.column_dimensions[openpyxl.utils.get_column_letter(2 + i)].width = 15
                ws_target.column_dimensions[openpyxl.utils.get_column_letter(usd_total_col_bonus2)].width = 15
                ws_target.column_dimensions[openpyxl.utils.get_column_letter(twd_total_col_bonus2)].width = 15
            wb_target.save(output_bonus2_xlsx_path)
            result_log.append(f"✅ Bonus2.xlsx 已成功生成並儲存於 {output_bonus2_xlsx_path}")
            if 'wb_source' in locals(): del wb_source; gc.collect()
            if 'wb_target' in locals(): del wb_target; gc.collect()
            return True
        except Exception as e_gen_b2:
            result_log.append(f"❌ 生成 Bonus2.xlsx 時發生錯誤: {str(e_gen_b2)}"); print(f"PYTHON_ERROR in _internal_generate_bonus2_report: {e_gen_b2}"); import traceback; traceback.print_exc(); return False

    def _internal_split_bonus2_sheets(bonus2_xlsx_path, output_directory_for_split_files):
        # ... (This function's internal logic remains the same) ...
        result_log.append(f"內部函數：開始分割 Bonus2.xlsx 從 {bonus2_xlsx_path} 到目錄 {output_directory_for_split_files}")
        split_files_generated_paths = []
        try:
            if not os.path.exists(bonus2_xlsx_path):
                result_log.append(f"❌ 錯誤: Bonus2.xlsx '{bonus2_xlsx_path}' 不存在，無法分割。"); return []
            workbook_to_split = openpyxl.load_workbook(bonus2_xlsx_path); date_str_prefix = datetime.now().strftime("%Y%m%d")
            if not os.path.exists(output_directory_for_split_files): os.makedirs(output_directory_for_split_files, exist_ok=True)
            for sheet_name_to_split in workbook_to_split.sheetnames:
                new_wb_for_sheet = openpyxl.Workbook()
                if new_wb_for_sheet.sheetnames[0] == 'Sheet': new_wb_for_sheet.remove(new_wb_for_sheet.active)
                source_sheet_obj = workbook_to_split[sheet_name_to_split]
                target_sheet_in_new_wb = new_wb_for_sheet.create_sheet(title=sheet_name_to_split)
                for col_letter, dim in source_sheet_obj.column_dimensions.items(): target_sheet_in_new_wb.column_dimensions[col_letter].width = dim.width # etc.
                for row_idx, dim in source_sheet_obj.row_dimensions.items(): target_sheet_in_new_wb.row_dimensions[row_idx].height = dim.height # etc.
                for row in source_sheet_obj.iter_rows():
                    for cell in row:
                        new_cell = target_sheet_in_new_wb[cell.coordinate]; new_cell.value = cell.value
                        if cell.has_style: copy_cell_format_for_api(cell, new_cell)
                split_filename = f"{date_str_prefix}{sheet_name_to_split}.xlsx"
                full_split_filepath = os.path.join(output_directory_for_split_files, split_filename)
                new_wb_for_sheet.save(full_split_filepath); split_files_generated_paths.append(full_split_filepath)
            result_log.append(f"✅ Bonus2.xlsx 已成功按工作表分割。共生成 {len(split_files_generated_paths)} 個檔案。")
            if 'workbook_to_split' in locals(): del workbook_to_split; gc.collect()
            return split_files_generated_paths
        except Exception as e_split_b2:
            result_log.append(f"❌ 分割 Bonus2.xlsx 時發生錯誤: {str(e_split_b2)}"); print(f"PYTHON_ERROR in _internal_split_bonus2_sheets: {e_split_b2}"); import traceback; traceback.print_exc(); return []

    def get_random_ua(): # This helper function is fine
        uas = ['Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',]
        return random.choice(uas)
    DEBUG = False # This is fine

    # --- CONFIGURATION LOADING (main_job specific, ensure it reflects global config) ---
    # `config` should be the global dictionary already loaded.
    # If not, it needs to be loaded here or passed to main_job.
    # For now, assume `config` is the global dictionary.
    max_concurrent_accounts = int(config.get('max_concurrent_accounts', 30))
    start_date = config.get('start_date', '2025/01/01')
    end_date = config.get('end_date', '2025/12/31')
    thread_start_delay = float(config.get('thread_start_delay', 0.5))
    max_login_attempts = int(config.get('max_login_attempts', 3))
    request_delay = float(config.get('request_delay', 2.0))
    max_request_retries = int(config.get('max_request_retries', 3))
    retry_delay = float(config.get('retry_delay', 3.0))
    dropbox_folder_for_output = config.get('dropbox_folder', '/output') # Renamed for clarity
    # dropbox_token is already current_dropbox_token
    # dropbox_account_file_path is already dropbox_account_file_path_config
    
    # --- NEW: Load accounts from Dropbox ---
    accounts = []
    result_log.append(f"準備從 Dropbox 讀取帳號列表: {dropbox_account_file_path_config}")
    if not current_dropbox_token:
        result_log.append("❌ 錯誤: Dropbox Token 未設定，無法讀取帳號檔案。")
        status["result"] = '\\n'.join(result_log)
        status["progress"] = "錯誤: Dropbox Token 未設定"
        status["running"] = False
        return '\\n'.join(result_log)

    try:
        dbx = dropbox.Dropbox(current_dropbox_token)
        _, res = dbx.files_download(path=dropbox_account_file_path_config)
        account_content = res.content.decode('utf-8')
        
        lines = [line.strip() for line in account_content.splitlines() if line.strip()]
        if not lines:
            result_log.append(f"❌ 警告: 從 Dropbox 路徑 {dropbox_account_file_path_config} 下載的帳號檔案內容為空。")
        else:
            for i in range(0, len(lines), 3):
                if i + 2 < len(lines):
                    name = lines[i]
                    acc = lines[i+1] # Renamed to avoid conflict with ACCOUNT parameter in fetch_account_data
                    password = lines[i+2]
                    accounts.append((name, acc, password))
            result_log.append(f"✅ 成功從 Dropbox ({dropbox_account_file_path_config}) 讀取並解析了 {len(accounts)} 個帳號。")

    except dropbox.exceptions.ApiError as dbx_err:
        err_msg = f"❌ Dropbox API 錯誤 (讀取帳號檔): {dbx_err}. 請檢查路徑 '{dropbox_account_file_path_config}' 是否正確且 Token 有讀取權限。"
        if isinstance(dbx_err.error, dropbox.files.DownloadError) and dbx_err.error.is_path():
            err_msg += " 錯誤細節: 路徑問題 (例如檔案不存在或路徑錯誤)."
        result_log.append(err_msg)
        status["result"] = '\\n'.join(result_log)
        status["progress"] = "錯誤: Dropbox帳號檔讀取失敗"
        status["running"] = False
        return '\\n'.join(result_log)
    except Exception as e_acc_read:
        result_log.append(f"❌ 從 Dropbox 讀取或解析帳號檔案時發生未知錯誤: {e_acc_read}")
        status["result"] = '\\n'.join(result_log)
        status["progress"] = "錯誤: 讀取帳號檔失敗"
        status["running"] = False
        return '\\n'.join(result_log)

    if not accounts:
        result_log.append('帳號列表為空。請確保 Dropbox 上的帳號檔案有內容且格式正確。')
        status["result"] = '\\n'.join(result_log)
        status["progress"] = "錯誤: 帳號列表為空"
        # Do not set running to False here if we want to allow an empty account list to "complete" without error
        # For now, let's consider it an error if no accounts to process, and stop.
        status["running"] = False
        return '\\n'.join(result_log)
    
    # --- End of new account loading logic ---

    # --- The rest of main_job (ThreadPoolExecutor, report generation, Dropbox upload of results) ---
    # This part uses the `accounts` list populated from Dropbox.
    # Ensure variable names for account username/password in the loop match what's in `accounts` tuple.
    # Example: for idx, (name, ACCOUNT_USER, PASSWORD_USER) in enumerate(accounts, 1):
    #            futures.append(executor.submit(fetch_account_data_and_save_to_csv, name, ACCOUNT_USER, PASSWORD_USER, ...))

    all_data_lock = threading.Lock() # Should be fine
    log_folder = os.path.join('logs', datetime.now().strftime('%Y%m%d_%H%M')) # For retry/fail logs
    os.makedirs(log_folder, exist_ok=True)
    retry_log_path = os.path.join(log_folder, 'retry.txt') # These might become less relevant if accounts always from Dropbox
    fail_log_path = os.path.join(log_folder, 'fail_log.txt')

    # --- make_request function (should be fine) ---
    def make_request(session, url, method='get', headers=None, data=None, retry_count=0):
        # ... (definition of make_request) ...
        if headers is None: headers = {}
        headers['User-Agent'] = get_random_ua()
        time.sleep(request_delay) # request_delay should be from config
        try:
            if method.lower() == 'get': resp = session.get(url, headers=headers, timeout=20) # Added timeout
            else: resp = session.post(url, headers=headers, data=data, timeout=20) # Added timeout
            if resp.status_code == 200: return resp
            if retry_count < max_request_retries: # max_request_retries from config
                time.sleep(retry_delay) # retry_delay from config
                return make_request(session, url, method, headers, data, retry_count+1)
            resp.raise_for_status() # Raise HTTPError for bad responses (4xx or 5xx) after retries
        except requests.exceptions.RequestException as e_req_make: # Catch specific request exceptions
            if retry_count < max_request_retries:
                time.sleep(retry_delay)
                return make_request(session, url, method, headers, data, retry_count+1)
            raise Exception(f"請求 {url} 最終失敗 ({type(e_req_make).__name__}): {e_req_make}") from e_req_make


    # --- fetch_account_data_and_save_to_csv function ---
    # Ensure the parameters are (name, ACCOUNT, PASSWORD, ocr, current_output_dir, headers_for_file)
    # ACCOUNT and PASSWORD here are from the tuple in the `accounts` list.
    def fetch_account_data_and_save_to_csv(name, user_account_id, user_password, ocr, current_output_dir, headers_for_file):
        # ... (The entire body of fetch_account_data_and_save_to_csv) ...
        # Make sure it uses user_account_id and user_password for login, not global ACCOUNT/PASSWORD.
        # It should return: name, user_account_id, csv_saved_boolean, num_rows
        # This function's internal logic for scraping, CSV saving, and the outer try-except wrapper remains as previously refined.
        # Example of parameter usage inside: data["MemberLogin1$txtAccound"] = user_account_id
        try:
            thread_id = threading.get_ident()
            def log_detail(message):
                timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
                print(f"[{timestamp}] [Thread-{thread_id}] [Acc: {name}] {message}")
            log_detail("處理開始")
            # ... (rest of the function as it was, ensuring ACCOUNT becomes user_account_id, PASSWORD becomes user_password)
            # Ensure the final return is: return name, user_account_id, save_successful, num_data_rows
            process_start_time = time.time()
            login_successful = False
            actual_login_attempts = 0
            ocr_total_classification_time_for_account = 0.0
            for attempt in range(1, max_login_attempts + 1): # max_login_attempts from config
                actual_login_attempts = attempt
                session = requests.Session()
                log_detail(f"登入嘗試 {attempt}/{max_login_attempts} - 開始")
                # ... login logic using user_account_id, user_password ...
                # ... error handling for login, including retry_log_path, fail_log_path ...
                # ... if login fails after max_attempts, raise Exception ...
                current_login_url = "https://member.star-rich.net/login"
                current_headers = {"Referer": current_login_url}
                resp_page = make_request(session, current_login_url, headers=current_headers)
                soup_login = BeautifulSoup(resp_page.text, "html.parser")
                # ... (captcha logic using ocr_instance) ...
                img_tag = soup_login.find("img", {"id": "MemberLogin1_Image1"})
                if not img_tag or not img_tag.get("src"): # Simplified error path
                    log_detail(f"登入嘗試 {attempt}: 找不到驗證碼圖片。"); time.sleep(retry_delay); continue
                img_url = "https://member.star-rich.net/" + img_tag["src"]
                img_resp = make_request(session, img_url, headers=current_headers)
                img_bytes = img_resp.content
                login_code = ocr.classification(img_bytes)
                if not login_code: log_detail(f"登入嘗試 {attempt}: OCR 未返回驗證碼。"); time.sleep(retry_delay); continue
                if login_code[-1] == '4': log_detail(f"登入嘗試 {attempt}: 驗證碼 {login_code} 以4結尾。"); time.sleep(retry_delay); continue
                
                login_data = { inp.get("name"): inp.get("value", "") for inp in soup_login.find_all("input") if inp.get("name") }
                login_data.update({
                    "MemberLogin1$txtAccound": user_account_id,
                    "MemberLogin1$txtPassword": user_password,
                    "MemberLogin1$txtCode": login_code,
                    "__EVENTTARGET": "MemberLogin1$lkbSignIn", "__EVENTARGUMENT": ""
                })
                login_resp = make_request(session, current_login_url, method='post', headers=current_headers, data=login_data)
                if "登出" in login_resp.text or "歡迎" in login_resp.text: login_successful = True; break
                log_detail(f"登入嘗試 {attempt} 失敗。") # Add more specific error detection if possible
                time.sleep(retry_delay) # retry_delay from config

            if not login_successful:
                log_detail(f"帳號 {name} ({user_account_id}) 連續 {actual_login_attempts} 次登入失敗。")
                # Append to retry/fail logs if needed, though primary account source is now Dropbox
                with open(retry_log_path, 'a', encoding='utf-8') as retry_f: retry_f.write(f"{name}\\n{user_account_id}\\n{user_password}\\n")
                with open(fail_log_path, 'a', encoding='utf-8') as fail_f: fail_f.write(f"{name}_{user_account_id} 連續{actual_login_attempts}次登入失敗\\n")
                raise Exception(f"帳號 {name} ({user_account_id}) 登入失敗")

            # ... (data scraping logic: home_resp, member_resp, bonus_history_resp etc.) ...
            # ... (This part should remain largely the same, ensure it uses the established `session`)
            home_url = "https://member.star-rich.net/default"; home_resp = make_request(session, home_url, headers=current_headers)
            home_soup = BeautifulSoup(home_resp.text, "html.parser"); h4s = home_soup.select(".h4")
            extra_data_home = [h.text.strip() for h in h4s[:5]] # bonus_point to item4
            while len(extra_data_home) < 5: extra_data_home.append("") # Ensure 5 elements
            star_level_tag = home_soup.select_one("#ctl00_cphPageInner_Label_Pin")
            extra_data_home.append(star_level_tag.text.strip() if star_level_tag else "")

            member_url = "https://member.star-rich.net/mem_memlist"; member_resp = make_request(session, member_url, headers=current_headers)
            member_soup = BeautifulSoup(member_resp.text, "html.parser")
            left_count_tag = member_soup.select_one("#ctl00_cphPageInner_cphContent_Label_LeftCount")
            right_count_tag = member_soup.select_one("#ctl00_cphPageInner_cphContent_Label_RightCount")
            extra_data_counts = [left_count_tag.text.strip() if left_count_tag else "", right_count_tag.text.strip() if right_count_tag else ""]
            
            all_extra_data_for_row = extra_data_home + extra_data_counts # Should be 5 + 1 + 2 = 8 elements

            bonus_history_url = "https://member.star-rich.net/bonushistory"
            resp_bonus_init = make_request(session, bonus_history_url, headers=current_headers)
            soup_bonus_init = BeautifulSoup(resp_bonus_init.text, "html.parser")
            # ... (Extract viewstate, eventvalidation, viewstategen for bonus history form) ...
            viewstate = soup_bonus_init.find("input", {"name": "__VIEWSTATE"})["value"]
            eventvalidation = soup_bonus_init.find("input", {"name": "__EVENTVALIDATION"})["value"]
            viewstategen = soup_bonus_init.find("input", {"name": "__VIEWSTATEGENERATOR"})["value"]
            form_data_bonus_hist = {
                "__EVENTTARGET": "ctl00$cphPageInner$cphContent$Button_Enter", "__EVENTARGUMENT": "",
                "__VIEWSTATE": viewstate, "__VIEWSTATEGENERATOR": viewstategen, "__EVENTVALIDATION": eventvalidation,
                "ctl00$cphPageInner$cphContent$txtStartDate": start_date, # start_date from config
                "ctl00$cphPageInner$cphContent$txtEndDate": end_date, # end_date from config
            }
            response_bonus_page = make_request(session, bonus_history_url, method='post', headers=current_headers, data=form_data_bonus_hist)
            current_bonus_soup = BeautifulSoup(response_bonus_page.text, "html.parser")
            account_all_data_rows = []
            # ... (Loop for bonus history pages, extracting table data) ...
            first_data_row_processed = False
            while True: # Loop through pages of bonus history
                target_table = None # Find the correct table
                for tbl_iter in current_bonus_soup.find_all("table"):
                    if any("獎金" in th.get_text(strip=True) for th in tbl_iter.find_all("th")): target_table = tbl_iter; break
                if not target_table: break # No more tables or data
                for tr_idx, tr_iter in enumerate(target_table.find_all("tr")):
                    if tr_idx == 0: continue # Skip header row
                    cols = [td.get_text(strip=True) for td in tr_iter.find_all("td")]
                    if cols and "總計" not in cols[0]:
                        # Original cols[:-1] are 16 bonus items. Headers are 24.
                        # So, bonus data + extra_data should match headers_for_file.
                        # headers_for_file[:16] are bonus items. headers_for_file[16:] are the 8 extra_data items.
                        bonus_data_part = cols[:-1] # These are the first 16 columns
                        if not first_data_row_processed:
                             full_row = bonus_data_part + all_extra_data_for_row
                             first_data_row_processed = True
                        else:
                             full_row = bonus_data_part + [""] * len(all_extra_data_for_row)
                        account_all_data_rows.append(full_row)
                
                next_btn = current_bonus_soup.find(id="ctl00_cphPageInner$cphContent$hpl_Forward")
                if not next_btn or 'disabled' in next_btn.attrs.get('class', []): break # No next page
                # Prepare form data for next page...
                viewstate_next = current_bonus_soup.find("input", {"name": "__VIEWSTATE"})["value"]
                eventvalidation_next = current_bonus_soup.find("input", {"name": "__EVENTVALIDATION"})["value"]
                viewstategen_next = current_bonus_soup.find("input", {"name": "__VIEWSTATEGENERATOR"})["value"]
                form_data_bonus_next = {
                    "__EVENTTARGET": "ctl00$cphPageInner$cphContent$hpl_Forward", "__EVENTARGUMENT": "",
                    "__VIEWSTATE": viewstate_next, "__VIEWSTATEGENERATOR": viewstategen_next, "__EVENTVALIDATION": eventvalidation_next,
                    "ctl00$cphPageInner$cphContent$txtStartDate": start_date, 
                    "ctl00$cphPageInner$cphContent$txtEndDate": end_date,
                }
                response_bonus_page = make_request(session, bonus_history_url, method='post', headers=current_headers, data=form_data_bonus_next)
                current_bonus_soup = BeautifulSoup(response_bonus_page.text, "html.parser")

            if not account_all_data_rows and first_data_row_processed == False: # No transactions, but logged in
                 blank_bonus_part = [""] * (len(headers_for_file) - len(all_extra_data_for_row))
                 account_all_data_rows.append(blank_bonus_part + all_extra_data_for_row)
                 log_detail("無交易紀錄，但已記錄主頁統計數據。")
            
            # --- CSV Saving ---
            csv_file_name = f"{name}_{user_account_id}.csv" # Use user_account_id
            csv_file_path = os.path.join(current_output_dir, csv_file_name)
            save_successful_csv = False
            try:
                with open(csv_file_path, 'w', newline='', encoding='utf-8-sig') as cf:
                    writer = csv.writer(cf)
                    writer.writerow(headers_for_file) # headers_for_file from main_job scope
                    if account_all_data_rows: writer.writerows(account_all_data_rows)
                log_detail(f"CSV 儲存成功: {csv_file_path} ({len(account_all_data_rows)} 行)")
                save_successful_csv = True
            except Exception as e_csv_write:
                log_detail(f"❌ CSV 儲存失敗 {csv_file_path}: {e_csv_write}")

            log_detail(f"帳號 {name} ({user_account_id}) 處理完成。")
            return name, user_account_id, save_successful_csv, len(account_all_data_rows) if account_all_data_rows else 0

        except Exception as e_fetch_outer: # Outer try-except for the whole fetch function
            log_detail(f"[CRITICAL WRAPPER] fetch_account_data_and_save_to_csv 發生未處理錯誤: {e_fetch_outer}")
            # import traceback; traceback.print_exc() # Uncomment for detailed stack trace in logs
            raise # Re-raise to be caught by as_completed loop

    # ThreadPoolExecutor part
    total_accounts = len(accounts)
    success_count = 0
    failed_accounts_info = [] # Store more info like (name, id, error_message)
    
    csv_generation_success_details = [] # Store (name, id, num_rows)
    csv_generation_failed_details = []  # Store (name, id, reason)

    if total_accounts > 0:
        with ThreadPoolExecutor(max_workers=max_concurrent_accounts) as executor:
            futures = []
            for idx, (name_acc, user_id_acc, pass_acc) in enumerate(accounts, 1): # Use new variable names
                futures.append(executor.submit(fetch_account_data_and_save_to_csv, name_acc, user_id_acc, pass_acc, ocr_instance, output_dir, headers_for_csv_and_excel))
                time.sleep(thread_start_delay) # thread_start_delay from config
            
            completed_count_threads = 0
            for future in as_completed(futures):
                try:
                    acc_name_res, acc_id_res, csv_saved_res, num_rows_res = future.result()
                    if csv_saved_res:
                        success_count += 1
                        csv_generation_success_details.append({'name': acc_name_res, 'id': acc_id_res, 'rows': num_rows_res})
                        result_log.append(f"帳號 {acc_name_res}_{acc_id_res} CSV 資料儲存成功 ({num_rows_res} 行)。")
                    else: # CSV saving reported as False by the function itself
                        csv_generation_failed_details.append({'name': acc_name_res, 'id': acc_id_res, 'reason': 'CSV儲存標記為失敗'})
                        result_log.append(f"帳號 {acc_name_res}_{acc_id_res} CSV 資料儲存失敗 (由函數回報)。")
                        failed_accounts_info.append(f"{acc_name_res}_{acc_id_res} (CSV儲存失敗)")


                except Exception as e_future: # Exception from fetch_account_data_and_save_to_csv
                    # Try to get account name/id if possible, e.g. if the exception message contains it.
                    # This part is tricky as the future might not easily give back its original arguments.
                    # The exception 'e_future' itself might contain the account name if we format it well in fetch_account_data
                    msg_future = str(e_future)
                    failed_accounts_info.append(msg_future) # Store the error message
                    result_log.append(f"[錯誤] 執行緒處理時發生錯誤: {msg_future}")
                finally:
                    completed_count_threads += 1
                    status["progress"] = f"CSV處理中: {completed_count_threads}/{total_accounts} (成功儲存CSV: {success_count})"
    else: # No accounts to process
        result_log.append("無帳號可處理 (從 Dropbox 讀取的列表為空)。")
        status["progress"] = "完成 (0/0)"


    # ... (rest of main_job: _create_excel_from_csv_files, Bonus2 generation, splitting, Dropbox upload of reports)
    # This part should mostly remain the same.
    # Ensure dropbox_folder_for_output is used for uploading reports, not the account file path.

    excel_file_path_local = None
    if success_count > 0:
        target_bonus_xlsx_path = os.path.join(output_dir, "bonus.xlsx")
        try:
            excel_file_path_local = _create_excel_from_csv_files(output_dir, target_bonus_xlsx_path, headers_for_csv_and_excel, result_log.append)
            if excel_file_path_local: result_log.append(f"主要的 bonus.xlsx 已成功從 CSV 生成於: {excel_file_path_local}")
            else: result_log.append(f"錯誤或警告: 未能從 CSV 檔案生成主要的 bonus.xlsx。")
        except Exception as e_csv_to_excel: result_log.append(f"❌ _create_excel_from_csv_files 生成 bonus.xlsx 錯誤: {e_csv_to_excel}")
    else: result_log.append("沒有成功生成的 CSV 檔案，跳過 bonus.xlsx 的創建。")
    
    bonus2_file_path_local = None; split_excel_files_paths = []
    if excel_file_path_local and os.path.exists(excel_file_path_local):
        bonus2_filename = 'Bonus2.xlsx'
        bonus2_file_path_local = os.path.join(output_dir, bonus2_filename)
        if _internal_generate_bonus2_report(excel_file_path_local, bonus2_file_path_local):
            result_log.append(f"Bonus2.xlsx 已成功生成於: {bonus2_file_path_local}")
            split_excel_files_paths = _internal_split_bonus2_sheets(bonus2_file_path_local, output_dir)
            if split_excel_files_paths: result_log.append(f"Bonus2.xlsx 已成功分割成 {len(split_excel_files_paths)} 個檔案。")
            else: result_log.append("警告: Bonus2.xlsx 分割未產生任何檔案或發生錯誤。")
        else:
            result_log.append(f"錯誤或警告: Bonus2.xlsx 未能成功生成。跳過分割。"); bonus2_file_path_local = None
    else: result_log.append("錯誤: 主要 bonus.xlsx 不存在，無法生成 Bonus2.xlsx。")

    # --- Dropbox upload of generated reports ---
    # ... (This section for uploading bonus.xlsx, Bonus2.xlsx, and split files remains, using `current_dropbox_token` and `dropbox_folder_for_output`)
    # Ensure folder_name (YYYYMMDD_HHMM) is used for the subfolder on Dropbox for reports.
    dropbox_status_msg_for_summary = ""
    uploaded_count_for_summary = 0
    upload_errors_for_summary = 0
    if current_dropbox_token:
        try:
            dbx_reports = dropbox.Dropbox(current_dropbox_token) # Use a clear variable if needed
            all_files_to_upload_reports = [f for f in os.listdir(output_dir) if os.path.isfile(os.path.join(output_dir, f)) and f.endswith('.xlsx')]
            
            base_dropbox_folder_reports = dropbox_folder_for_output.rstrip('/') # Use the config for reports output
            dropbox_target_run_folder_reports = f"{base_dropbox_folder_reports}/{folder_name}" # folder_name is YYYYMMDD_HHMM

            if all_files_to_upload_reports:
                for f_report in all_files_to_upload_reports:
                    local_path_report = os.path.join(output_dir, f_report)
                    try:
                        with open(local_path_report, 'rb') as content_report:
                            dbx_path_report = f"{dropbox_target_run_folder_reports}/{f_report}"
                            dbx_reports.files_upload(content_report.read(), dbx_path_report, mode=dropbox.files.WriteMode.overwrite)
                            result_log.append(f"  ✅ 已上傳報表 {f_report} 到 Dropbox ({dbx_path_report})")
                            uploaded_count_for_summary +=1
                    except Exception as e_dbx_report_upload:
                        result_log.append(f"  ❌ 上傳報表 {f_report} 到 Dropbox 失敗: {e_dbx_report_upload}")
                        upload_errors_for_summary +=1
                # ... (Update dropbox_status_msg_for_summary based on these counts) ...
                if uploaded_count_for_summary > 0 and upload_errors_for_summary == 0: dropbox_status_msg_for_summary = f"Dropbox報表: ✅ 所有 {uploaded_count_for_summary} 個已上傳到 {dropbox_target_run_folder_reports}"
                # ... other conditions for dropbox_status_msg_for_summary ...
            else: dropbox_status_msg_for_summary = f"Dropbox報表: ⚠️ {output_dir} 中無 .xlsx 報表可上傳 (目標: {dropbox_target_run_folder_reports})"
        except Exception as e_dbx_reports_generic: dropbox_status_msg_for_summary = f"Dropbox報表: ❌ 連接或操作時發生錯誤 - {str(e_dbx_reports_generic)}"
    else: dropbox_status_msg_for_summary = "Dropbox報表: ⚠️ Token未設定，跳過上傳"


    # --- Final summary ---
    final_summary_lines = []
    final_summary_lines.append(f"帳號處理成功: {success_count} / {total_accounts if total_accounts > 0 else 'N/A'}")
    if failed_accounts_info:
        final_summary_lines.append(f"帳號處理失敗: {len(failed_accounts_info)}")
        final_summary_lines.append("失敗詳情 (部分):")
        for fail_msg in failed_accounts_info[:5]: # Show first 5 errors
            final_summary_lines.append(f"  - {fail_msg}")
    else:
        final_summary_lines.append("帳號處理失敗: 0")
    
    if dropbox_status_msg_for_summary: final_summary_lines.append(dropbox_status_msg_for_summary)
    # ... (Add total time) ...
    end_time = time.time()
    # ... (Calculate hours, minutes, seconds for total_time_job = end_time - main_start_time) ...
    # Need main_start_time defined at the beginning of main_job's try block.
    # For now, just log `result_log`.
    
    status["result"] = '\\n'.join(final_summary_lines + result_log[-10:]) # Show summary and last 10 detailed logs
    status["progress"] = f"完成: {completed_count_threads if total_accounts > 0 else 0}/{total_accounts if total_accounts > 0 else 0} (成功: {success_count})"
    
    print("main_job 執行完畢.")

    # except Exception as e_main_job: (ensure main_job has its own outer try-except)
    # ... error logging ...
    # finally: status["running"] = False
    # return '\\n'.join(result_log)
    try:
        # ... (all of main_job's primary logic as above) ...
        pass # Placeholder for the main logic block
    except Exception as e_main_job_outer:
        error_message_main = f"main_job 執行時發生頂層錯誤: {str(e_main_job_outer)}"
        print(error_message_main)
        result_log.append(error_message_main)
        # import traceback; result_log.append(traceback.format_exc()) # For detailed stack trace
        status["result"] = '\\n'.join(result_log)
        status["progress"] = "發生嚴重錯誤，請查看日誌"
    finally:
        status["running"] = False
        print("main_job 執行緒結束 (無論成功或失敗)")
    return '\\n'.join(result_log) # Return all logs

# --- Global Config Loading (should happen once at app start) ---
config = {}
def load_config():
    global config
    default_config = {
        'mode': 0, 'max_concurrent_accounts': 5, 'start_date': '2024/01/01', 'end_date': '2024/12/31',
        'thread_start_delay': 0.5, 'max_login_attempts': 2, 'request_delay': 1.0,
        'max_request_retries': 2, 'retry_delay': 3.0,
        'dropbox_token': '', 'dropbox_folder': '/output', # This is for report outputs
        'dropbox_app_key': '', 'dropbox_app_secret': '', 'dropbox_refresh_token': '',
        'dropbox_account_file_path': '/Apps/ExcelAPI-app/account/account.txt', # Default path for account file
        'API_ACTION_PASSWORD': 'CHANGEME' # Default password, user should change
    }
    if os.path.exists('config.txt'):
        with open('config.txt', 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith('#'): continue
                if '=' in line:
                    k, v = line.split('=', 1)
                    k, v = k.strip(), v.split('#')[0].strip()
                    if k in default_config: # Check against default_config keys
                        if isinstance(default_config[k], int): config[k] = int(v)
                        elif isinstance(default_config[k], float): config[k] = float(v)
                        else: config[k] = v
                    else: # Allow new keys not in default_config (e.g. if user adds custom ones)
                        config[k] = v
    # Ensure all default keys are present in config, using default value if not in file
    for k_default, v_default in default_config.items():
        if k_default not in config:
            config[k_default] = v_default
    
    # Handle Dropbox token refresh if necessary (this logic was already present)
    if not config.get('dropbox_token') and config.get('dropbox_app_key') and config.get('dropbox_app_secret') and config.get('dropbox_refresh_token'):
        # ... (get_access_token_from_refresh logic as before) ...
        def get_access_token_from_refresh_global(): # Make it distinct or pass config
            try:
                resp_token = requests.post(
                    'https://api.dropbox.com/oauth2/token',
                    data={
                        'grant_type': 'refresh_token', 'refresh_token': config['dropbox_refresh_token'],
                        'client_id': config['dropbox_app_key'], 'client_secret': config['dropbox_app_secret'],
                    }
                )
                resp_token.raise_for_status()
                return resp_token.json().get('access_token', '')
            except Exception as e_token_refresh:
                print(f"❌ Dropbox refresh token 換取 access token 失敗 (global load): {e_token_refresh}")
                return ''
        
        refreshed_token = get_access_token_from_refresh_global()
        if refreshed_token:
            config['dropbox_token'] = refreshed_token
            print("✅ 已自動用 refresh token 取得 Dropbox access token (global load)")
        else:
            print("❌ 無法自動取得 Dropbox access token (global load)，請檢查 refresh token 設定")

    print(f"配置已載入: {config}") # Log loaded config
load_config() # Load config at app startup

@app.route('/run_main', methods=['POST'])
# ... (run_main remains the same, it uses the global `config` implicitly via `main_job`)
# ... unchanged ...
def run_main():
    print("收到 /run_main 請求")
    if status["running"]:
        print("狀態 busy - 先前任務仍在執行")
        return jsonify({"status": "busy", "message": "先前的任務仍在執行中，請稍後再試。"})
    status["running"] = True
    status["result"] = "初始化中，準備開始執行主要腳本..."
    status["progress"] = "初始化中..."
    print("準備啟動新 thread 執行 main_job")
    thread = threading.Thread(target=main_job) # main_job will use global config
    thread.start()
    print("已啟動新 thread 執行 main_job")
    return jsonify({"status": "started", "message": "主要腳本已啟動執行。請稍後透過 /status 檢查進度。"})

@app.route('/status', methods=['GET'])
# ... (get_status remains the same) ...
# ... unchanged ...
def get_status():
    print(f"[STATUS_ENDPOINT] 目前的 status 字典是: {status}")
    return jsonify(status)

@app.route('/')
# ... (serve_index remains the same) ...
# ... unchanged ...
def serve_index():
    return send_from_directory('.', 'index.html')

# --- NEW API Endpoints for Account File Management ---
@app.route('/api/account_file', methods=['GET', 'POST'])
def manage_account_file():
    global config # Access the global config
    # Ensure necessary Dropbox config is available
    api_password = config.get('API_ACTION_PASSWORD', 'CHANGEME_IF_NOT_SET_IN_CONFIG') # Fallback if key missing
    token = config.get('dropbox_token')
    account_file_dbx_path = config.get('dropbox_account_file_path')

    if not token or not account_file_dbx_path:
        return jsonify({"error": "Dropbox token 或帳號檔案路徑未在伺服器正確設定。"}), 500

    if request.method == 'POST':
        try:
            data = request.get_json()
            if not data:
                return jsonify({"error": "請求中未找到 JSON 資料"}), 400
            
            user_password = data.get('password')
            account_data_content = data.get('account_data')

            if not user_password or account_data_content is None: # account_data can be empty string
                return jsonify({"error": "請求中缺少 'password' 或 'account_data' 欄位"}), 400

            if user_password != api_password:
                return jsonify({"error": "密碼錯誤"}), 403 # Forbidden

            dbx = dropbox.Dropbox(token)
            try:
                # Encode account_data to bytes for upload
                file_bytes = account_data_content.encode('utf-8')
                dbx.files_upload(file_bytes, account_file_dbx_path, mode=dropbox.files.WriteMode.overwrite)
                print(f"帳號檔案已成功上傳到 Dropbox: {account_file_dbx_path}")
                return jsonify({"message": "帳號檔案已成功更新到 Dropbox。"}), 200
            except dropbox.exceptions.ApiError as e:
                print(f"Dropbox API 錯誤 (上傳帳號檔): {e}")
                return jsonify({"error": f"Dropbox API 錯誤: {str(e)}"}), 500
            except Exception as e_upload:
                print(f"上傳帳號檔案到 Dropbox 時發生未知錯誤: {e_upload}")
                return jsonify({"error": f"上傳時發生內部錯誤: {str(e_upload)}"}), 500
        except Exception as e_json:
            return jsonify({"error": f"處理請求時發生錯誤: {str(e_json)}"}), 400


    elif request.method == 'GET':
        user_password_query = request.args.get('password')
        if not user_password_query:
            return jsonify({"error": "請求中缺少 'password' 查詢參數"}), 400
        
        if user_password_query != api_password:
            return jsonify({"error": "密碼錯誤"}), 403

        dbx = dropbox.Dropbox(token)
        try:
            _, res = dbx.files_download(path=account_file_dbx_path)
            account_file_content_bytes = res.content
            # Send as a file download
            return send_from_directory(
                directory='.', # Dummy, not used as we provide content directly
                path='account.txt', # Suggested filename for download
                mimetype='text/plain',
                as_attachment=True,
                attachment_filename='account.txt', # For Flask 2.0+ use download_name
                # For older Flask, you might need a BytesIO wrapper for send_file
                # from io import BytesIO
                # return send_file(BytesIO(account_file_content_bytes), mimetype='text/plain', as_attachment=True, download_name='account.txt')
                response_class=lambda r, s, h: app.response_class(r, status=s, headers=h), # Hack to make it work with string content
                environ=request.environ # For Flask to construct response
            )
            # For modern Flask (send_file with BytesIO is more robust):
            # from io import BytesIO
            # return send_file(BytesIO(account_file_content_bytes), mimetype='text/plain', download_name='account.txt', as_attachment=True)
            # The above send_from_directory is a bit of a hack for direct content.
            # A more robust way for direct content without an actual file:
            from flask import Response
            return Response(
                account_file_content_bytes,
                mimetype="text/plain",
                headers={"Content-disposition": "attachment; filename=account.txt"}
            )

        except dropbox.exceptions.ApiError as e:
            if isinstance(e.error, dropbox.files.DownloadError) and e.error.is_path():
                 print(f"Dropbox API 錯誤 (下載帳號檔): 找不到檔案或路徑錯誤 {account_file_dbx_path} - {e}")
                 return jsonify({"error": f"Dropbox 找不到帳號檔案或路徑錯誤: {account_file_dbx_path}"}), 404
            print(f"Dropbox API 錯誤 (下載帳號檔): {e}")
            return jsonify({"error": f"Dropbox API 錯誤: {str(e)}"}), 500
        except Exception as e_download:
            print(f"下載帳號檔案時發生未知錯誤: {e_download}")
            return jsonify({"error": f"下載時發生內部錯誤: {str(e_download)}"}), 500

if __name__ == '__main__':
# ... (main app run remains the same) ...
# ... unchanged ...
    print("=== 進入 __main__ 啟動 Flask ===")
    # load_config() # Config is now loaded globally when script is imported/run
    port = int(os.environ.get("PORT", int(config.get("FLASK_PORT", 5000)))) # Use config for port too
    app.run(host=config.get("FLASK_HOST", '0.0.0.0'), port=port)