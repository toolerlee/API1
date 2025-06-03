from flask import Flask, request, jsonify, send_from_directory
import threading
import dropbox
import os

print("=== Flask API 啟動 ===")

app = Flask(__name__)

status = {"running": False, "result": None, "progress": "尚未開始"}

# 將原本的主流程包成一個函數
def main_job():
    global status
    status["running"] = True
    status["result"] = "處理中...請稍候..."
    status["progress"] = "初始化中..."
    print("=== main_job 啟動 ===")
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

    # 初始化 OCR 實例 (在所有 import 之後，config 讀取之前或之後均可，但在 ThreadPoolExecutor 之前)
    # 使用 show_ad=False 可以避免一些不必要的日誌或行為
    try:
        ocr_instance = ddddocr.DdddOcr(show_ad=False)
        print("ddddocr.DdddOcr 實例已在 main_job 中成功初始化。")
    except Exception as e_ocr_init:
        print(f"錯誤：在 main_job 中初始化 ddddocr.DdddOcr 失敗: {e_ocr_init}")
        # 根據您的錯誤處理策略，這裡可能需要 return 或引發更上層的錯誤
        status["result"] = f"錯誤：OCR組件初始化失敗: {e_ocr_init}"
        status["progress"] = "OCR 初始化失敗"
        status["running"] = False
        return # 如果 OCR 初始化失敗，則無法繼續

    result_log = []
    try:
        def get_random_ua():
            uas = [
                'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
                'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15',
                'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:124.0) Gecko/20100101 Firefox/124.0',
                'Mozilla/5.0 (Linux; Android 10; SM-G975F) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Mobile Safari/537.36',
            ]
            return random.choice(uas)
        DEBUG = False

        # 讀取 config.txt 設定
        print("讀取 config.txt ...")
        config = {
            'mode': 0,
            'max_concurrent_accounts': 30,
            'start_date': '2025/01/01',
            'end_date': '2025/12/31',
            'thread_start_delay': 0.5,
            'max_login_attempts': 3,
            'request_delay': 2.0,
            'max_request_retries': 3,
            'retry_delay': 3.0,
            'dropbox_token': '',
            'dropbox_folder': '/output',
            'dropbox_app_key': '',
            'dropbox_app_secret': '',
            'dropbox_refresh_token': '',
        }
        if os.path.exists('config.txt'):
            with open('config.txt', 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if not line or line.startswith('#'):
                        continue
                    if '=' in line:
                        k, v = line.split('=', 1)
                        k, v = k.strip(), v.split('#')[0].strip()
                        if k in config:
                            if k in ['mode', 'max_concurrent_accounts', 'max_login_attempts', 'max_request_retries']:
                                config[k] = int(v)
                            elif k in ['thread_start_delay', 'request_delay', 'retry_delay']:
                                config[k] = float(v)
                            else:
                                config[k] = v
        mode = config['mode']
        max_concurrent_accounts = config['max_concurrent_accounts']
        start_date = config['start_date']
        end_date = config['end_date']
        thread_start_delay = config['thread_start_delay']
        max_login_attempts = config['max_login_attempts']
        request_delay = config['request_delay']
        max_request_retries = config['max_request_retries']
        retry_delay = config['retry_delay']
        dropbox_token = config.get('dropbox_token', '')
        dropbox_folder = config.get('dropbox_folder', '/output')
        dropbox_app_key = config.get('dropbox_app_key', '')
        dropbox_app_secret = config.get('dropbox_app_secret', '')
        dropbox_refresh_token = config.get('dropbox_refresh_token', '')
        # 自動 refresh token
        def get_access_token_from_refresh():
            if not (dropbox_app_key and dropbox_app_secret and dropbox_refresh_token):
                return ''
            try:
                resp = requests.post(
                    'https://api.dropbox.com/oauth2/token',
                    data={
                        'grant_type': 'refresh_token',
                        'refresh_token': dropbox_refresh_token,
                        'client_id': dropbox_app_key,
                        'client_secret': dropbox_app_secret,
                    }
                )
                resp.raise_for_status()
                return resp.json().get('access_token', '')
            except Exception as e:
                result_log.append(f"❌ Dropbox refresh token 換取 access token 失敗: {e}")
                return ''
        # 若 dropbox_token 為空，則自動 refresh
        if not dropbox_token and dropbox_app_key and dropbox_app_secret and dropbox_refresh_token:
            dropbox_token = get_access_token_from_refresh()
            if dropbox_token:
                result_log.append("✅ 已自動用 refresh token 取得 Dropbox access token")
            else:
                result_log.append("❌ 無法自動取得 Dropbox access token，請檢查 refresh token 設定")
        print(f"API.py 讀到的 dropbox_token: {dropbox_token}")
        # 依 mode 決定帳號來源
        if mode == 1:
            log_dirs = [d for d in glob.glob(os.path.join('logs', '*')) if os.path.isdir(d)]
            if not log_dirs:
                result_log.append('❌ [重試模式] 找不到 logs 目錄下的任何執行資料夾，無法進行重試。')
                status["result"] = '\n'.join(result_log)
                status["progress"] = "錯誤: 找不到logs資料夾"
                return '\n'.join(result_log)
            latest_log_dir = max(log_dirs, key=os.path.getmtime)
            retry_file = os.path.join(latest_log_dir, 'retry.txt')
            if not os.path.exists(retry_file):
                result_log.append(f'❌ [重試模式] 找不到 {retry_file}，無法進行重試。')
                status["result"] = '\n'.join(result_log)
                status["progress"] = f"錯誤: 找不到 {os.path.basename(retry_file)}"
                return '\n'.join(result_log)
            account_file = retry_file
            result_log.append(f'[重試模式] 來源: {retry_file}')
        else:
            account_file = 'account.txt'
            result_log.append(f'[一般模式] 來源: account.txt')
        
        accounts = []

        if os.path.exists(account_file):
            with open(account_file, 'r', encoding='utf-8') as f:
                lines = [line.strip() for line in f if line.strip()]
                for i in range(0, len(lines), 3):
                    if i + 2 < len(lines):
                        name = lines[i]
                        account = lines[i+1]
                        password = lines[i+2]
                        accounts.append((name, account, password))
        else:
            result_log.append(f'找不到 {account_file}，請建立並填入帳號資料。')
            status["result"] = '\n'.join(result_log)
            status["progress"] = f"錯誤: 找不到 {os.path.basename(account_file)}"
            return '\n'.join(result_log)
        
        # --- 在 accounts 列表確定後 (無論是成功載入還是為空)，更新進度 --- 
        if not accounts and not os.path.exists(account_file): 
            # 這種情況是 account_file 不存在，上面的 else 已經處理並 return，理論上不會到這裡
            # 但為了防禦性程式設計，保留一個判斷
            pass # status["progress"] 已在上面的 else 中設定
        elif not accounts:
             status["progress"] = f"準備中 (0/0) - {os.path.basename(account_file)} 為空或格式不正確"
        else: # accounts 有內容
            status["progress"] = f"準備中 (0/{len(accounts)})"

        all_data = []
        all_data_lock = threading.Lock()
        log_folder = os.path.join('logs', datetime.now().strftime('%Y%m%d_%H%M'))
        os.makedirs(log_folder, exist_ok=True)
        retry_log_path = os.path.join(log_folder, 'retry.txt')
        fail_log_path = os.path.join(log_folder, 'fail_log.txt')
        def make_request(session, url, method='get', headers=None, data=None, retry_count=0):
            if headers is None:
                headers = {}
            headers['User-Agent'] = get_random_ua()
            time.sleep(request_delay)
            try:
                if method.lower() == 'get':
                    resp = session.get(url, headers=headers)
                else:
                    resp = session.post(url, headers=headers, data=data)
                if resp.status_code == 200:
                    return resp
                if retry_count < max_request_retries:
                    time.sleep(retry_delay)
                    return make_request(session, url, method, headers, data, retry_count+1)
                raise Exception(f"請求失敗，HTTP狀態碼: {resp.status_code}")
            except Exception as e:
                if retry_count < max_request_retries:
                    time.sleep(retry_delay)
                    return make_request(session, url, method, headers, data, retry_count+1)
                raise
        def fetch_account_data(name, ACCOUNT, PASSWORD, ocr):
            thread_id = threading.get_ident()
            
            def log_detail(message):
                timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
                full_message = f"[{timestamp}] [Thread-{thread_id}] [Acc: {name}] {message}"
                print(full_message)

            log_detail("處理開始")
            process_start_time = time.time()
            
            session = requests.Session()
            login_url = "https://member.star-rich.net/login"
            headers = {"Referer": login_url}
            
            login_successful = False
            actual_login_attempts = 0

            for attempt in range(1, max_login_attempts + 1):
                actual_login_attempts = attempt
                log_detail(f"登入嘗試 {attempt}/{max_login_attempts} - 開始")
                login_attempt_start_time = time.time()

                log_detail("  B1. 請求登入頁面 - 開始")
                page_req_start_time = time.time()
                resp = make_request(session, login_url, headers=headers)
                log_detail(f"  B1. 請求登入頁面 - 完成 (耗時: {time.time() - page_req_start_time:.2f}s)")
                
                soup = BeautifulSoup(resp.text, "html.parser")
                inputs = soup.find_all("input")
                data = {}
                for inp in inputs:
                    name_attr = inp.get("name")
                    value_attr = inp.get("value", "")
                    if name_attr:
                        data[name_attr] = value_attr
                
                img_tag = soup.find("img", {"id": "MemberLogin1_Image1"})
                if not img_tag or not img_tag.get("src"):
                    log_detail("  B2. 錯誤：找不到驗證碼圖片標籤或 src。跳過此登入嘗試。")
                    time.sleep(retry_delay)
                    continue 

                img_url = "https://member.star-rich.net/" + img_tag["src"]
                
                ocr_total_classification_time = 0.0
                ocr_loop_attempts = 0

                current_code = ""
                while True:
                    ocr_loop_attempts += 1
                    log_detail(f"    B2a. OCR嘗試 {ocr_loop_attempts} - 請求驗證碼圖片 - 開始")
                    ocr_img_req_start_time = time.time()
                    img_resp = make_request(session, img_url, headers=headers)
                    log_detail(f"    B2a. OCR嘗試 {ocr_loop_attempts} - 請求驗證碼圖片 - 完成 (耗時: {time.time() - ocr_img_req_start_time:.2f}s)")
                    
                    img_bytes = img_resp.content
                    
                    log_detail(f"    B2b. OCR嘗試 {ocr_loop_attempts} - 驗證碼識別(ddddocr) - 開始")
                    ocr_classify_start_time = time.time()
                    current_code = ocr.classification(img_bytes)
                    single_ocr_duration = time.time() - ocr_classify_start_time
                    ocr_total_classification_time += single_ocr_duration
                    log_detail(f"    B2b. OCR嘗試 {ocr_loop_attempts} - 驗證碼識別(ddddocr) - 完成 (耗時: {single_ocr_duration:.2f}s, 識別結果: {current_code})")
                    
                    if not (len(current_code) > 0 and current_code[-1] == '4'):
                        log_detail(f"    B2c. OCR嘗試 {ocr_loop_attempts} - 驗證碼 '{current_code}' 符合要求，跳出OCR迴圈")
                        break
                    else:
                        log_detail(f"    B2c. OCR嘗試 {ocr_loop_attempts} - 驗證碼 '{current_code}' 以'4'結尾，重新獲取")
                    
                    if ocr_loop_attempts >= 5:
                        log_detail(f"    B2d. OCR嘗試超過 {ocr_loop_attempts} 次，強制使用最後結果 '{current_code}' 並跳出OCR迴圈")
                        break
                
                data["MemberLogin1$txtAccound"] = ACCOUNT
                data["MemberLogin1$txtPassword"] = PASSWORD
                data["MemberLogin1$txtCode"] = current_code
                data["__EVENTTARGET"] = "MemberLogin1$lkbSignIn"
                data["__EVENTARGUMENT"] = ""

                log_detail("  B3. 提交登入表單 - 開始")
                submit_login_start_time = time.time()
                login_resp = make_request(session, login_url, method='post', headers=headers, data=data)
                log_detail(f"  B3. 提交登入表單 - 完成 (耗時: {time.time() - submit_login_start_time:.2f}s)")

                if "登出" in login_resp.text or "歡迎" in login_resp.text:
                    log_detail(f"登入嘗試 {attempt} - 成功 (耗時: {time.time() - login_attempt_start_time:.2f}s)")
                    login_successful = True
                    break
                
                error_msg_detected = ""
                if "驗證碼" in login_resp.text or "驗證碼錯誤" in login_resp.text or "請輸入驗證碼" in login_resp.text:
                    error_msg_detected = "驗證碼相關錯誤"
                else:
                    error_msg_detected = "其他登入失敗"
                
                log_detail(f"登入嘗試 {attempt} - 失敗 ({error_msg_detected}, 本次嘗試耗時: {time.time() - login_attempt_start_time:.2f}s)")
                if attempt < max_login_attempts:
                    log_detail("    準備進行下一次登入嘗試...")
                    time.sleep(retry_delay)
            
            if not login_successful:
                log_detail(f"連續 {actual_login_attempts} 次登入失敗")
                with open(retry_log_path, 'a', encoding='utf-8') as retry_file:
                    retry_file.write(f"{name}\n{ACCOUNT}\n{PASSWORD}\n")
                with open(fail_log_path, 'a', encoding='utf-8') as fail_file:
                    fail_file.write(f"{name}_{ACCOUNT} 連續{actual_login_attempts}次登入失敗\n")
                raise Exception(f"{name}_{ACCOUNT} 連續{actual_login_attempts}次登入失敗 (OCR總耗時: {ocr_total_classification_time:.2f}s)")

            log_detail("C1. 請求主頁 - 開始")
            home_page_start_time = time.time()
            home_url = "https://member.star-rich.net/default"
            home_resp = make_request(session, home_url, headers=headers)
            log_detail(f"C1. 請求主頁 - 完成 (耗時: {time.time() - home_page_start_time:.2f}s)")

            log_detail("C2. 解析主頁內容 - 開始")
            parse_home_start_time = time.time()
            home_soup = BeautifulSoup(home_resp.text, "html.parser")
            h4s = home_soup.select(".h4")
            bonus_point = h4s[0].text.strip() if len(h4s) > 0 else ""
            item1 = h4s[1].text.strip() if len(h4s) > 1 else ""
            item2 = h4s[2].text.strip() if len(h4s) > 2 else ""
            item3 = h4s[3].text.strip() if len(h4s) > 3 else ""
            item4 = h4s[4].text.strip() if len(h4s) > 4 else ""
            star_level_tag = home_soup.select_one("#ctl00_cphPageInner_Label_Pin")
            star_level = star_level_tag.text.strip() if star_level_tag else ""
            log_detail(f"C2. 解析主頁內容 - 完成 (耗時: {time.time() - parse_home_start_time:.2f}s)")

            log_detail("D1. 請求會員列表頁 - 開始")
            member_list_start_time = time.time()
            member_url = "https://member.star-rich.net/mem_memlist"
            member_resp = make_request(session, member_url, headers=headers)
            log_detail(f"D1. 請求會員列表頁 - 完成 (耗時: {time.time() - member_list_start_time:.2f}s)")

            log_detail("D2. 解析會員列表頁 - 開始")
            parse_member_start_time = time.time()
            member_soup = BeautifulSoup(member_resp.text, "html.parser")
            left_count_tag = member_soup.select_one("#ctl00_cphPageInner_cphContent_Label_LeftCount")
            right_count_tag = member_soup.select_one("#ctl00_cphPageInner_cphContent_Label_RightCount")
            left_count = left_count_tag.text.strip() if left_count_tag else ""
            right_count = right_count_tag.text.strip() if right_count_tag else ""
            log_detail(f"D2. 解析會員列表頁 - 完成 (耗時: {time.time() - parse_member_start_time:.2f}s)")

            extra_data = [bonus_point, item1, item2, item3, item4, star_level, left_count, right_count]

            log_detail("E1. 請求獎金歷史初始頁 - 開始")
            bonus_init_start_time = time.time()
            bonus_history_url = "https://member.star-rich.net/bonushistory"
            resp_bonus_init = make_request(session, bonus_history_url, headers=headers)
            log_detail(f"E1. 請求獎金歷史初始頁 - 完成 (耗時: {time.time() - bonus_init_start_time:.2f}s)")

            log_detail("E2. 解析獎金歷史初始頁 - 開始")
            parse_bonus_init_start_time = time.time()
            soup_bonus = BeautifulSoup(resp_bonus_init.text, "html.parser")
            viewstate = soup_bonus.find("input", {"name": "__VIEWSTATE"})["value"]
            eventvalidation = soup_bonus.find("input", {"name": "__EVENTVALIDATION"})["value"]
            viewstategen = soup_bonus.find("input", {"name": "__VIEWSTATEGENERATOR"})["value"]
            log_detail(f"E2. 解析獎金歷史初始頁 - 完成 (耗時: {time.time() - parse_bonus_init_start_time:.2f}s)")
            
            form_data_bonus = {
                "__EVENTTARGET": "ctl00$cphPageInner$cphContent$Button_Enter",
                "__EVENTARGUMENT": "",
                "__VIEWSTATE": viewstate,
                "__VIEWSTATEGENERATOR": viewstategen,
                "__EVENTVALIDATION": eventvalidation,
                "ctl00$cphPageInner$cphContent$txtStartDate": start_date,
                "ctl00$cphPageInner$cphContent$txtEndDate": end_date,
            }

            log_detail("F1. 提交獎金歷史查詢表單 (第一頁) - 開始")
            bonus_submit_start_time = time.time()
            response_bonus_page = make_request(session, bonus_history_url, method='post', headers=headers, data=form_data_bonus)
            log_detail(f"F1. 提交獎金歷史查詢表單 (第一頁) - 完成 (耗時: {time.time() - bonus_submit_start_time:.2f}s)")
            current_bonus_soup = BeautifulSoup(response_bonus_page.text, "html.parser")

            account_all_rows = []
            bonus_page_count = 0
            first_bonus_page_processed = False
            while True:
                bonus_page_count += 1
                log_detail(f"  G{bonus_page_count}. 處理獎金歷史第 {bonus_page_count} 頁 - 開始解析")
                page_parse_start_time = time.time()
                
                tables = current_bonus_soup.find_all("table")
                target_table = None
                for t_table in tables:
                    ths = [th.get_text(strip=True) for th in t_table.find_all("th")]
                    if any("獎金" in th_text for th_text in ths):
                        target_table = t_table
                        break
                
                if target_table is None:
                    log_detail(f"  G{bonus_page_count}. 在第 {bonus_page_count} 頁未找到目標表格，結束獎金歷史處理。")
                    if bonus_page_count == 1:
                         log_detail(f"    注意：帳號 {name} 未抓取到任何獎金歷史資料。")
                    break

                rows_on_page = 0
                for row_idx, row_element in enumerate(target_table.find_all("tr")):
                    if row_idx == 0:
                        continue
                    cols = [td.get_text(strip=True) for td in row_element.find_all("td")]
                    if cols:
                        if "總計" in cols[0]:
                            continue
                        rows_on_page +=1
                        if not first_bonus_page_processed:
                            account_all_rows.append(cols[:-1] + extra_data)
                            first_bonus_page_processed = True
                        else:
                            account_all_rows.append(cols[:-1] + [""] * len(extra_data))
                
                log_detail(f"  G{bonus_page_count}. 處理獎金歷史第 {bonus_page_count} 頁 - 完成解析 (找到 {rows_on_page} 行資料, 耗時: {time.time() - page_parse_start_time:.2f}s)")
                
                next_btn = current_bonus_soup.find(id="ctl00_cphPageInner$cphContent$hpl_Forward")
                if not next_btn or 'disabled' in next_btn.attrs.get('class', []):
                    log_detail(f"  獎金歷史第 {bonus_page_count} 頁 - 無下一頁按鈕或已禁用，結束分頁。")
                    break
                
                log_detail(f"  請求獎金歷史下一頁 (第 {bonus_page_count + 1} 頁) - 開始")
                next_page_req_start_time = time.time()
                
                viewstate = current_bonus_soup.find("input", {"name": "__VIEWSTATE"})["value"]
                eventvalidation = current_bonus_soup.find("input", {"name": "__EVENTVALIDATION"})["value"]
                viewstategen = current_bonus_soup.find("input", {"name": "__VIEWSTATEGENERATOR"})["value"]
                
                form_data_bonus_next_page = {
                    "__EVENTTARGET": "ctl00$cphPageInner$cphContent$hpl_Forward",
                    "__EVENTARGUMENT": "",
                    "__VIEWSTATE": viewstate,
                    "__VIEWSTATEGENERATOR": viewstategen,
                    "__EVENTVALIDATION": eventvalidation,
                    "ctl00$cphPageInner$cphContent$txtStartDate": start_date,
                    "ctl00$cphPageInner$cphContent$txtEndDate": end_date,
                }
                response_bonus_page = make_request(session, bonus_history_url, method='post', headers=headers, data=form_data_bonus_next_page)
                log_detail(f"  請求獎金歷史下一頁 (第 {bonus_page_count + 1} 頁) - 完成 (耗時: {time.time() - next_page_req_start_time:.2f}s)")
                current_bonus_soup = BeautifulSoup(response_bonus_page.text, "html.parser")
            
            if not account_all_rows:
                 log_detail(f"    最終：帳號 {name} 未收集到任何獎金歷史資料列。")

            with all_data_lock:
                for single_data_row in account_all_rows:
                    row_to_add_globally = [name, ACCOUNT] + single_data_row
                    all_data.append(row_to_add_globally)
                log_detail(f"H1. 已將 {len(account_all_rows)} 行資料添加完成。")

            log_detail(f"處理完成 (總耗時: {time.time() - process_start_time:.2f}s, 其中OCR總耗時: {ocr_total_classification_time:.2f}s)")

        total_accounts = len(accounts)
        success_count = 0
        failed_accounts = []
        start_time = time.time()
        result_log.append(f"\n開始處理，總帳號數量: {total_accounts}")
        with ThreadPoolExecutor(max_workers=max_concurrent_accounts) as executor:
            futures = []
            started_count = 0
            completed_count = 0
            for idx, (name, ACCOUNT, PASSWORD) in enumerate(accounts, 1):
                futures.append(executor.submit(fetch_account_data, name, ACCOUNT, PASSWORD, ocr_instance))
                started_count += 1
                status["progress"] = f"已提交任務: {started_count}/{total_accounts} (處理中: {completed_count}/{total_accounts})"
                result_log.append(f"已啟動處理帳號: {started_count}/{total_accounts}")
                time.sleep(thread_start_delay)
            
            for future in as_completed(futures):
                try:
                    future.result()
                    success_count += 1
                    result_log.append(f"目前登入成功進度: {success_count}/{total_accounts}")
                except Exception as e:
                    msg = str(e)
                    if '連續' in msg and '登入失敗' in msg:
                        failed_accounts.append(msg)
                    else:
                        result_log.append(f"[警告] 非帳號登入失敗異常：{msg}")
                finally:
                    completed_count += 1
                    status["progress"] = f"處理中: {completed_count}/{total_accounts} (成功: {success_count})"

        end_time = time.time()
        total_time = end_time - start_time
        hours = int(total_time // 3600)
        minutes = int((total_time % 3600) // 60)
        seconds = int(total_time % 60)

        excel_file_path_local = None
        if all_data:
            # 定義 Excel 工作表中的欄位標頭 (不包含 "帳號名稱" 和 "登入帳號")
            # 您需要根據實際從網頁解析的獎金表格欄位來準確定義 bonus_table_actual_headers
            # 以下為範例，假設獎金表有這些欄位 (通常是 td 的內容)
            bonus_table_actual_headers = [
                "日期", "內容", "項目", "變動類型", 
                "變動前", "變動數", "變動後", "狀態" # 假設這些是您 cols[:-1] 對應的標頭
            ]
            extra_data_headers = [
                "紅利積分", "電子錢包", "獎金暫存", "註冊分", 
                "商品券", "星級", "左區人數", "右區人數"
            ]
            headers_for_excel_sheet = bonus_table_actual_headers + extra_data_headers

            folder_name = datetime.now().strftime('%Y%m%d_%H%M')
            output_dir = os.path.join('output', folder_name)
            os.makedirs(output_dir, exist_ok=True)
            wb = Workbook()
            if 'Sheet' in wb.sheetnames:
                del wb['Sheet']
            
            acc_dict = defaultdict(list)
            # 現在 all_data 中的每個元素就是一個包含帳號資訊和數據的列表
            # row_with_name_acc 的格式是 [name, ACCOUNT, data_col1, data_col2, ...]
            for row_with_name_acc in all_data:
                acc_key = f"{row_with_name_acc[0]}_{row_with_name_acc[1]}" # name_ACCOUNT
                data_part = row_with_name_acc[2:] # 實際要寫入 Excel 的數據行 (去除了 name 和 ACCOUNT)
                acc_dict[acc_key].append(data_part)
            
            for acc_key, data_rows_for_acc in acc_dict.items():
                ws = wb.create_sheet(acc_key[:31])
                ws.append(headers_for_excel_sheet) # 寫入我們定義的標頭
                for single_data_row_part in data_rows_for_acc:
                    ws.append(single_data_row_part)
            
            excel_file_path_local = os.path.join(output_dir, 'bonus.xlsx')
            wb.save(excel_file_path_local)

        final_summary_for_status = []
        final_summary_for_status.append("=== 處理結果摘要 ===")

        if total_accounts > 0:
            final_summary_for_status.append(f"帳號處理進度: {completed_count}/{total_accounts} 個帳號已嘗試")
            final_summary_for_status.append(f"成功擷取資料: {success_count} 個帳號")
            final_summary_for_status.append(f"登入/處理失敗: {len(failed_accounts)} 個帳號")
            if failed_accounts:
                final_summary_for_status.append("失敗帳號列表:")
                for acc_failure_msg_item in failed_accounts:
                    final_summary_for_status.append(f"  - {acc_failure_msg_item}")
        else:
            final_summary_for_status.append("資訊: 未載入任何帳號進行處理。")

        if all_data and excel_file_path_local:
            dropbox_status_msg_for_summary = ""
            if dropbox_token:
                try:
                    dbx = dropbox.Dropbox(dropbox_token)
                    files_to_upload_in_dir = [f_name for f_name in os.listdir(output_dir) if os.path.isfile(os.path.join(output_dir, f_name))]
                    if files_to_upload_in_dir:
                        for f_to_upload_name in files_to_upload_in_dir:
                            path_of_file_to_upload = os.path.join(output_dir, f_to_upload_name)
                            with open(path_of_file_to_upload, 'rb') as content_f_upload:
                                dropbox_upload_target_path = dropbox_folder.rstrip('/') + '/' + f_to_upload_name
                                dbx.files_upload(content_f_upload.read(), dropbox_upload_target_path, mode=dropbox.files.WriteMode.overwrite)
                        dropbox_status_msg_for_summary = f"Dropbox狀態: ✅ 檔案已從 {output_dir} 上傳到 {dropbox_folder}"
                    else:
                        dropbox_status_msg_for_summary = f"Dropbox狀態: ⚠️ {output_dir} 中無檔案 (Excel可能未儲存成功)，未執行上傳。"
                except Exception as e_dbx_upload:
                    dropbox_status_msg_for_summary = f"Dropbox狀態: ❌ 上傳失敗 - {str(e_dbx_upload)}"
            else:
                dropbox_status_msg_for_summary = "Dropbox狀態: ⚠️ 未設定Dropbox Token，跳過上傳"
            
            final_summary_for_status.append(dropbox_status_msg_for_summary)
            final_summary_for_status.append(f"輸出檔案: {excel_file_path_local}")
        
        elif not all_data and total_accounts > 0 :
            final_summary_for_status.append("最終結果: 未產生任何資料檔案。")

        final_summary_for_status.append(f"總耗時: {hours}小時 {minutes}分鐘 {seconds}秒")
        
        status["result"] = '\\n'.join(final_summary_for_status)
        status["progress"] = f"完成: {completed_count}/{total_accounts} (成功: {success_count})"

        console_message_final = "main_job 執行完畢. "
        if all_data:
            console_message_final += "資料已產生並嘗試上傳."
        elif total_accounts > 0:
            console_message_final += "但未產出任何資料."
        else:
            console_message_final += "無帳號可處理."
        print(console_message_final)

    except Exception as e:
        error_message = f"main_job 執行時發生嚴重錯誤: {str(e)}"
        print(error_message)
        result_log.append(error_message)
        status["result"] = '\n'.join(result_log)
        status["progress"] = "發生錯誤，請查看日誌"
    finally:
        status["running"] = False
        print("main_job 執行緒結束 (無論成功或失敗)")

    return '\n'.join(result_log)

@app.route('/run_main', methods=['POST'])
def run_main():
    print("收到 /run_main 請求")
    if status["running"]:
        print("狀態 busy - 先前任務仍在執行")
        return jsonify({"status": "busy", "message": "先前的任務仍在執行中，請稍後再試。"})
    
    status["running"] = True
    status["result"] = "初始化中，準備開始執行主要腳本..."
    status["progress"] = "初始化中..."
    
    print("準備啟動新 thread 執行 main_job")
    thread = threading.Thread(target=main_job)
    thread.start()
    print("已啟動新 thread 執行 main_job")
    return jsonify({"status": "started", "message": "主要腳本已啟動執行。請稍後透過 /status 檢查進度。"})

@app.route('/status', methods=['GET'])
def get_status():
    print(f"[STATUS_ENDPOINT] 目前的 status 字典是: {status}")
    return jsonify(status)

@app.route('/')
def serve_index():
    return send_from_directory('.', 'index.html')

if __name__ == '__main__':
    print("=== 進入 __main__ 啟動 Flask ===")
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)