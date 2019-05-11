import os.path
import re
import subprocess
import winreg
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
    try:
        keyname = r'SOFTWARE\WOW6432Node\Microsoft\VisualStudio\10.0\VC\VCRedist\x64'
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, keyname)
        pkg_ver = winreg.QueryValueEx(key, 'Version')[0].strip('v')
    except FileNotFoundError:
        pkg_ver = '0.0.0.0'

    return pkg_ver

## 安裝 Visual C++ 2010 x64 Redistributable 10.0.40219.325
def install_vsredist():
    url = 'https://download.microsoft.com/download/1/6/5/165255E7-1014-4D0A-B094-B6A430A6BFFC/vcredist_x64.exe'
    vcdist = download_file(url, check_dir('~/.skcom'))
    ps_exec(vcdist + ' setup /passive', admin=True)
    os.remove(vcdist)
    return True

## 檢查群益 API 元件是否已註冊
def check_skcom():
    cmd = r'reg query HKLM\SOFTWARE\Classes\TypeLib /s /f SKCOM.dll'
    result = ps_exec(cmd)

    version = '0.0.0.0'
    if result is not None:
        lines = result.split('\r\n')
        for line in lines:
            # 找 DLL 檔案路徑
            m = re.match(r'.+REG_SZ\s+(.+SKCOM.dll)', line)
            if m is not None:
                dll_path = m.group(1)
                # 取檔案摘要內容裡版本號碼
                fso = comtypes.client.CreateObject('Scripting.FileSystemObject')
                version = fso.GetFileVersion(dll_path)

    return version

## 安裝群益 API 元件
def install_skcom():
    url = 'https://www.capital.com.tw/Service2/download/api_zip/CapitalAPI_2.13.16.zip'

    # 建立元件目錄
    com_path = check_dir(r'~\.skcom\lib')

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
def download_file(url, save_path):
    abs_path = check_dir(save_path)
    file_path = r'%s\%s' % (abs_path, url.split('/')[-1])

    with requests.get(url, stream=True) as r:
        r.raise_for_status()
        with open(file_path, 'wb') as f:
            for chunk in r.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
            f.flush()

    return file_path

# 檢查目錄, 不存在就建立目錄, 完成後回傳絕對路徑
def check_dir(usr_path):
    rel_path = os.path.expanduser(usr_path)
    if not os.path.isdir(rel_path):
        os.makedirs(rel_path)
    abs_path = os.path.realpath(rel_path)
    return abs_path
