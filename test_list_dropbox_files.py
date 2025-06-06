import os
import json
import requests
import dropbox

# 讀取環境變數
DROPBOX_REFRESH_TOKEN = os.getenv('dropbox_refresh_token')
DROPBOX_APP_KEY = os.getenv('dropbox_app_key')
DROPBOX_APP_SECRET = os.getenv('dropbox_app_secret')
DROPBOX_ACCOUNT_FILE_PATH = os.getenv('dropbox_account_file_path')

print('DEBUG: DROPBOX_REFRESH_TOKEN =', repr(DROPBOX_REFRESH_TOKEN))
print('DEBUG: DROPBOX_APP_KEY =', repr(DROPBOX_APP_KEY))
print('DEBUG: DROPBOX_APP_SECRET =', repr(DROPBOX_APP_SECRET))

print('RENDER DEBUG: dropbox_app_key =', repr(os.getenv('dropbox_app_key')))
print('RENDER DEBUG: dropbox_app_secret =', repr(os.getenv('dropbox_app_secret')))
print('RENDER DEBUG: dropbox_refresh_token =', repr(os.getenv('dropbox_refresh_token')))
print('RENDER DEBUG: dropbox_account_file_path =', repr(os.getenv('dropbox_account_file_path')))

def get_access_token():
    url = "https://api.dropbox.com/oauth2/token"
    data = {
        "grant_type": "refresh_token",
        "refresh_token": DROPBOX_REFRESH_TOKEN,
    }
    auth = (DROPBOX_APP_KEY, DROPBOX_APP_SECRET)
    response = requests.post(url, data=data, auth=auth)
    response.raise_for_status()
    return response.json().get("access_token")

def list_dropbox_files(path=''):
    access_token = get_access_token()
    dbx = dropbox.Dropbox(access_token)
    try:
        result = dbx.files_list_folder(path=path)
        files = []
        for entry in result.entries:
            files.append({
                'name': entry.name,
                'path_display': getattr(entry, 'path_display', ''),
                'type': type(entry).__name__
            })
        print(json.dumps({'status': 'success', 'files': files}, ensure_ascii=False, indent=2))
    except Exception as e:
        print(json.dumps({'status': 'error', 'message': str(e)}, ensure_ascii=False, indent=2))

if __name__ == '__main__':
    # 預設列出 App Folder 根目錄
    list_dropbox_files()
    # 如需列出 /account 目錄，請改成 list_dropbox_files('/account') 