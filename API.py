from flask import Flask, request, jsonify
import threading
import dropbox
import os

app = Flask(__name__)

status = {"running": False, "result": None}

# 將原本的主流程包成一個函數
def main_job():
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
    result_log = []
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
            result_log.append(f"\n❌ Dropbox refresh token 換取 access token 失敗: {e}")
            return ''
    # 若 dropbox_token 為空，則自動 refresh
    if not dropbox_token and dropbox_app_key and dropbox_app_secret and dropbox_refresh_token:
        dropbox_token = get_access_token_from_refresh()
        if dropbox_token:
            result_log.append("\n✅ 已自動用 refresh token 取得 Dropbox access token")
        else:
            result_log.append("\n❌ 無法自動取得 Dropbox access token，請檢查 refresh token 設定")
    print(f"API.py 讀到的 dropbox_token: {dropbox_token}")
    # 依 mode 決定帳號來源
    if mode == 1:
        log_dirs = [d for d in glob.glob(os.path.join('logs', '*')) if os.path.isdir(d)]
        if not log_dirs:
            result_log.append('❌ [重試模式] 找不到 logs 目錄下的任何執行資料夾，無法進行重試。')
            return '\n'.join(result_log)
        latest_log_dir = max(log_dirs, key=os.path.getmtime)
        retry_file = os.path.join(latest_log_dir, 'retry.txt')
        if not os.path.exists(retry_file):
            result_log.append(f'❌ [重試模式] 找不到 {retry_file}，無法進行重試。')
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
        return '\n'.join(result_log)
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
    def fetch_account_data(name, ACCOUNT, PASSWORD):
        session = requests.Session()
        login_url = "https://member.star-rich.net/login"
        headers = {"Referer": login_url}
        for attempt in range(1, max_login_attempts + 1):
            resp = make_request(session, login_url, headers=headers)
            soup = BeautifulSoup(resp.text, "html.parser")
            inputs = soup.find_all("input")
            data = {}
            for inp in inputs:
                name_attr = inp.get("name")
                value = inp.get("value", "")
                if name_attr:
                    data[name_attr] = value
            img_tag = soup.find("img", {"id": "MemberLogin1_Image1"})
            img_url = "https://member.star-rich.net/" + img_tag["src"]
            ocr = ddddocr.DdddOcr()
            while True:
                img_resp = make_request(session, img_url, headers=headers)
                img_bytes = img_resp.content
                code = ocr.classification(img_bytes)
                if not (len(code) > 0 and code[-1] == '4'):
                    break
            data["MemberLogin1$txtAccound"] = ACCOUNT
            data["MemberLogin1$txtPassword"] = PASSWORD
            data["MemberLogin1$txtCode"] = code
            data["__EVENTTARGET"] = "MemberLogin1$lkbSignIn"
            data["__EVENTARGUMENT"] = ""
            login_resp = make_request(session, login_url, method='post', headers=headers, data=data)
            if "登出" in login_resp.text or "歡迎" in login_resp.text:
                break
            if "驗證碼" in login_resp.text or "驗證碼錯誤" in login_resp.text or "請輸入驗證碼" in login_resp.text:
                continue
            with open(retry_log_path, 'a', encoding='utf-8') as retry_log:
                retry_log.write(f"{name}\n{ACCOUNT}\n{PASSWORD}\n")
            break
        else:
            with open(retry_log_path, 'a', encoding='utf-8') as retry_log:
                retry_log.write(f"{name}\n{ACCOUNT}\n{PASSWORD}\n")
            with open(fail_log_path, 'a', encoding='utf-8') as fail_log:
                fail_log.write(f"{name}_{ACCOUNT} 連續{max_login_attempts}次登入失敗\n")
            raise Exception(f"{name}_{ACCOUNT} 連續{max_login_attempts}次登入失敗")
        soup = BeautifulSoup(login_resp.text, "html.parser")
        home_url = "https://member.star-rich.net/default"
        home_resp = make_request(session, home_url, headers=headers)
        home_soup = BeautifulSoup(home_resp.text, "html.parser")
        h4s = home_soup.select(".h4")
        bonus_point = h4s[0].text.strip() if len(h4s) > 0 else ""
        item1 = h4s[1].text.strip() if len(h4s) > 1 else ""
        item2 = h4s[2].text.strip() if len(h4s) > 2 else ""
        item3 = h4s[3].text.strip() if len(h4s) > 3 else ""
        item4 = h4s[4].text.strip() if len(h4s) > 4 else ""
        star_level = home_soup.select_one("#ctl00_cphPageInner_Label_Pin")
        star_level = star_level.text.strip() if star_level else ""
        member_url = "https://member.star-rich.net/mem_memlist"
        member_resp = make_request(session, member_url, headers=headers)
        member_soup = BeautifulSoup(member_resp.text, "html.parser")
        left_count = member_soup.select_one("#ctl00_cphPageInner_cphContent_Label_LeftCount")
        right_count = member_soup.select_one("#ctl00_cphPageInner_cphContent_Label_RightCount")
        left_count = left_count.text.strip() if left_count else ""
        right_count = right_count.text.strip() if right_count else ""
        extra_data = [bonus_point, item1, item2, item3, item4, star_level, left_count, right_count]
        url = "https://member.star-rich.net/bonushistory"
        resp = make_request(session, url, headers=headers)
        soup = BeautifulSoup(resp.text, "html.parser")
        viewstate = soup.find("input", {"name": "__VIEWSTATE"})["value"]
        eventvalidation = soup.find("input", {"name": "__EVENTVALIDATION"})["value"]
        viewstategen = soup.find("input", {"name": "__VIEWSTATEGENERATOR"})["value"]
        data = {
            "__EVENTTARGET": "ctl00$cphPageInner$cphContent$Button_Enter",
            "__EVENTARGUMENT": "",
            "__VIEWSTATE": viewstate,
            "__VIEWSTATEGENERATOR": viewstategen,
            "__EVENTVALIDATION": eventvalidation,
            "ctl00$cphPageInner$cphContent$txtStartDate": "2025/01/01",
            "ctl00$cphPageInner$cphContent$txtEndDate": "2025/12/31",
        }
        response = make_request(session, url, method='post', headers=headers, data=data)
        soup = BeautifulSoup(response.text, "html.parser")
        tables = soup.find_all("table")
        target_table = None
        for t in tables:
            ths = [th.get_text(strip=True) for th in t.find_all("th")]
            if any("獎金" in th for th in ths):
                target_table = t
                break
        if target_table is None:
            return
        headers_row = [th.get_text(strip=True) for th in target_table.find_all("th")][:-1]
        headers_row += ["紅利積分", "電子錢包", "獎金暫存", "註冊分", "商品券", "星級", "左區人數", "右區人數"]
        all_rows = []
        first_page = True
        while True:
            tables = soup.find_all("table")
            target_table = None
            for t in tables:
                ths = [th.get_text(strip=True) for th in t.find_all("th")]
                if any("獎金" in th for th in ths):
                    target_table = t
                    break
            if target_table is None:
                break
            for row in target_table.find_all("tr")[1:]:
                cols = [td.get_text(strip=True) for td in row.find_all("td")]
                if cols:
                    if "總計" in cols[0]:
                        continue
                    if first_page and len(all_rows) == 0:
                        all_rows.append(cols[:-1] + extra_data)
                    else:
                        all_rows.append(cols[:-1] + [""] * len(extra_data))
            first_page = False
            viewstate = soup.find("input", {"name": "__VIEWSTATE"})["value"]
            eventvalidation = soup.find("input", {"name": "__EVENTVALIDATION"})["value"]
            viewstategen = soup.find("input", {"name": "__VIEWSTATEGENERATOR"})["value"]
            next_btn = soup.find(id="ctl00_cphPageInner_cphContent_hpl_Forward")
            if not next_btn or 'disabled' in next_btn.attrs.get('class', []):
                break
            data = {
                "__EVENTTARGET": "ctl00$cphPageInner$cphContent$hpl_Forward",
                "__EVENTARGUMENT": "",
                "__VIEWSTATE": viewstate,
                "__VIEWSTATEGENERATOR": viewstategen,
                "__EVENTVALIDATION": eventvalidation,
                "ctl00$cphPageInner$cphContent$txtStartDate": "2025/01/01",
                "ctl00$cphPageInner$cphContent$txtEndDate": "2025/12/31",
            }
            response = make_request(session, url, method='post', headers=headers, data=data)
            soup = BeautifulSoup(response.text, "html.parser")
        for row in all_rows:
            row_with_acc = [name, ACCOUNT] + row
            with all_data_lock:
                all_data.append((headers_row, row_with_acc))
    total_accounts = len(accounts)
    success_count = 0
    failed_accounts = []
    start_time = time.time()
    result_log.append(f"\n開始處理，總帳號數量: {total_accounts}")
    with ThreadPoolExecutor(max_workers=max_concurrent_accounts) as executor:
        futures = []
        started_count = 0
        for idx, (name, ACCOUNT, PASSWORD) in enumerate(accounts, 1):
            futures.append(executor.submit(fetch_account_data, name, ACCOUNT, PASSWORD))
            started_count += 1
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
    end_time = time.time()
    total_time = end_time - start_time
    hours = int(total_time // 3600)
    minutes = int((total_time % 3600) // 60)
    seconds = int(total_time % 60)
    if all_data:
        headers_row = all_data[0][0]
        all_rows_data = [row_with_acc for _, row_with_acc in all_data]
        folder_name = datetime.now().strftime('%Y%m%d_%H%M')
        output_dir = os.path.join('output', folder_name)
        os.makedirs(output_dir, exist_ok=True)
        wb = Workbook()
        if 'Sheet' in wb.sheetnames:
            del wb['Sheet']
        acc_dict = defaultdict(list)
        for row in all_rows_data:
            acc_key = f"{row[0]}_{row[1]}"
            acc_dict[acc_key].append(row[2:])
        for acc_key, rows in acc_dict.items():
            ws = wb.create_sheet(acc_key[:31])
            ws.append(headers_row)
            for row in rows:
                ws.append(row)
        wb.save(os.path.join(output_dir, 'bonus.xlsx'))
        # === Dropbox 自動上傳 ===
        if dropbox_token:
            try:
                dbx = dropbox.Dropbox(dropbox_token)
                # 找到 output 目錄下最新的資料夾
                output_root = 'output'
                subfolders = [os.path.join(output_root, d) for d in os.listdir(output_root) if os.path.isdir(os.path.join(output_root, d))]
                if subfolders:
                    latest_folder = max(subfolders, key=os.path.getmtime)
                    for fname in os.listdir(latest_folder):
                        fpath = os.path.join(latest_folder, fname)
                        if os.path.isfile(fpath):
                            with open(fpath, 'rb') as f:
                                dropbox_path = dropbox_folder.rstrip('/') + '/' + fname
                                dbx.files_upload(f.read(), dropbox_path, mode=dropbox.files.WriteMode.overwrite)
                    result_log.append(f"\n✅ 已自動上傳 {latest_folder} 內所有檔案到 Dropbox {dropbox_folder}")
                else:
                    result_log.append("\n⚠️ 找不到 output 目錄下的任何資料夾，無法上傳 Dropbox")
            except Exception as e:
                result_log.append(f"\n❌ Dropbox 上傳失敗: {e}")
        else:
            result_log.append("\n⚠️ 未設定 Dropbox Token，未執行自動上傳")
        result_log.append("\n\n=== 處理完成總結報告 ===")
        result_log.append(f"總帳號數量: {total_accounts}")
        result_log.append(f"成功寫入數量: {success_count}")
        result_log.append(f"登入失敗數量: {len(failed_accounts)}")
        if failed_accounts:
            result_log.append("\n登入失敗帳號:")
            for acc in failed_accounts:
                result_log.append(f"- {acc}")
        result_log.append(f"\n總耗時: {hours}小時 {minutes}分鐘 {seconds}秒")
        result_log.append(f"資料已寫入: {output_dir}/bonus.xlsx")
    else:
        result_log.append("\n沒有任何帳號抓到資料")
    return '\n'.join(result_log)

@app.route('/run_main', methods=['POST'])
def run_main():
    if status["running"]:
        return jsonify({"status": "busy"})
    def job():
        status["running"] = True
        status["result"] = main_job()
        status["running"] = False
    threading.Thread(target=job).start()
    return jsonify({"status": "started"})

@app.route('/status', methods=['GET'])
def get_status():
    return jsonify(status)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)