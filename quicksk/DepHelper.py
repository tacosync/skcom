import os.path
import re
import subprocess
import zipfile

import comtypes.client
import requests

# 用 Power Shell 執行一個指令
def ps_exec(cmd, admin=False):
    tokens = cmd.split(' ')
    prog = tokens[0]
    args = '"%s"' % ' '.join(tokens[1:])
    spcmd = [
        'powershell.exe', 'Start-Process',
        '-FilePath', prog,
        '-ArgumentList', args
    ]

    # 用系統管理員身分執行
    if admin:
        # TODO: 需要想一下系統管理員模式怎麼取得 stdout
        spcmd.append('-Verb')
        spcmd.append('RunAs')
    else:
        # 接收 stdout
        spcmd.append('-NoNewWindow')
        spcmd.append('-Wait')

    completed = subprocess.run(spcmd, capture_output=True)
    if completed.returncode == 0:
        return completed.stdout.decode('cp950')
    else:
        return None

## 檢查 Visual Studio 2010 可轉發套件是否已安裝
# TODO: 目前檢查版本有缺陷, 等之後明朗化再實作
def check_vsredist():
    target = r'(Microsoft Visual C\+\+ 2010\s{1,2}x64 Redistributable) - (.+)'
    cmd = 'wmic product get name'.split(' ')
    completed = subprocess.run(cmd, shell=True, capture_output=True)
    packages = completed.stdout.decode('cp950').split('\r\r\n')

    pkg_name = None
    pkg_ver = None
    for p in packages:
        m = re.match(target, p)
        if m is not None:
            pkg_name = m.group(1)
            pkg_ver  = m.group(2) # 先留著以後檢查版本用
            break

    return (pkg_name is not None)

## 安裝 Visual Studio 2010 可轉發套件
def install_vsredist():
    return False

## 檢查群益 API 元件是否已註冊
def check_skcom():
    cmd = r'reg query HKLM\SOFTWARE\Classes\TypeLib /s /f SKCOM.dll'
    result = ps_exec(cmd)

    version = '0.0.0.0'
    if result is not None:
        lines = result.split('\r\n')
        for line in lines:
            m = re.match(r'.+REG_SZ\s+(.+SKCOM.dll)', line)
            if m is not None:
                dll_path = m.group(1)
                fso = comtypes.client.CreateObject('Scripting.FileSystemObject')
                version = fso.GetFileVersion(dll_path)

    return version

# 安裝群益 API 元件
def install_skcom():
    url = 'https://www.capital.com.tw/Service2/download/api_zip/CapitalAPI_2.13.16.zip'

    # 建立元件目錄
    com_path = os.path.expanduser(r'~\.skcom\lib')
    if not os.path.isdir(com_path):
        os.makedirs(com_path)

    # 下載
    file_path = download_file(url, com_path)

    # 解壓縮
    # 只解壓縮 64-bits DLL 檔案, 其他非必要檔案不處理
    # 讀檔時需要用 CP437 編碼處理檔名, 寫檔時需要用 CP950 處理檔名
    with zipfile.ZipFile(file_path, 'r') as archive:
        for name437 in archive.namelist():
            name950 = name437.encode('cp437').decode('cp950')
            if re.match(r'元件/x64/.+\.dll', name950):
                dest_path = r'%s\%s' % (com_path, name950.split('/')[-1])
                with archive.open(name437, 'r') as cf, \
                     open(dest_path, 'wb') as xf:
                    xf.write(cf.read())

    # 註冊元件
    cmd = r'regsvr32 %s\SKCOM.dll' % com_path
    ret = ps_exec(cmd, admin=True)

    return True

# 使用 8K 緩衝下載檔案
def download_file(url, dest_path):
    file_name = url.split('/')[-1]
    file_path = dest_path + '\\' + file_name

    with requests.get(url, stream=True) as r:
        r.raise_for_status()
        with open(file_path, 'wb') as f:
            for chunk in r.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
            f.flush()

    return file_path
