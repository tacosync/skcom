"""
quicksk.helper
"""

import os.path
import re
import shutil
import site
import subprocess
import time
import winreg
import zipfile

import comtypes.client
from comtypes import COMError
import requests
from packaging import version


_SLIENT_MODE = False

def set_silent(mode):
    """
    控制 console_print 是否輸出
    """
    # pylint: disable=superfluous-parens, unidiomatic-typecheck,global-statement
    assert(type(mode) != 'bool')

    global _SLIENT_MODE
    _SLIENT_MODE = mode

def console_print(msg):
    """
    可以控制開關的 console 輸出
    """
    if not _SLIENT_MODE:
        print(msg)

def ps_exec(cmd, admin=False):
    """
    用 Power Shell 執行一個指令
    * 可以使用系統管理者執行
    * 參數含有空白時會發生問題, 例如 tasklist 的 /fi
    """
    if isinstance(cmd, list):
        tokens = cmd
    else:
        tokens = cmd.split(' ')
    prog = tokens[0]
    spcmd = [
        'powershell.exe', 'Start-Process',
        '-FilePath', prog
    ]
    if len(tokens) > 1:
        quoted_tokens = list(map(lambda i: '"%s"' % i, tokens[1:]))
        args = ','.join(quoted_tokens)
        spcmd.append('-ArgumentList')
        spcmd.append(args)

    spcmd.append('-NoNewWindow')
    spcmd.append('-Wait')
    # 用系統管理員身分執行
    if admin:
        # TODO: 需要想一下系統管理員模式怎麼取得 stdout
        spcmd.append('-Verb')
        spcmd.append('RunAs')

    # Python 3.7 才能用這個寫法
    # completed = subprocess.run(spcmd, capture_output=True)
    completed = subprocess.run(spcmd, stdout=subprocess.PIPE)
    if completed.returncode == 0:
        return completed.stdout.decode('cp950')

    return None

def cmd_exec(cmd):
    """
    用 cmd 執行一個指令
    * 適用指令較廣泛
    * 無法使用系統管理者身分執行
    """
    if isinstance(cmd, list):
        tokens = cmd
    else:
        tokens = cmd.split(' ')

    spcmd = ['cmd', '/C'] + tokens
    # Python 3.7 才能用這個寫法
    # completed = subprocess.run(spcmd, capture_output=True)
    completed = subprocess.run(spcmd, stdout=subprocess.PIPE)
    if completed.returncode == 0:
        # 因為被 cmd 包了一層, 不管怎樣都是 return 0
        return completed.stdout.decode('cp950')

    return None

def verof_vcredist():
    """
    取得 Visual C++ 2010 可轉發套件版本資訊
    """
    try:
        # 版本字串格式: v10.0.40219.325
        keyname = r'SOFTWARE\WOW6432Node\Microsoft\VisualStudio\10.0\VC\VCRedist\x64'
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, keyname)
        pkg_ver = winreg.QueryValueEx(key, 'Version')[0].strip('v')
        if not re.match(r'(\d+\.){3}\d+', pkg_ver):
            pkg_ver = '0.0.0.0'
    except FileNotFoundError:
        pkg_ver = '0.0.0.0'

    return version.parse(pkg_ver)

def install_vcredist():
    """
    安裝 Visual C++ 2010 x64 Redistributable 10.0.40219.325
    """

    # 下載與安裝
    url = 'https://download.microsoft.com/download/1/6/5/' + \
          '165255E7-1014-4D0A-B094-B6A430A6BFFC/vcredist_x64.exe'
    vcdist = download_file(url, check_dir('~/.skcom'))
    ps_exec(vcdist + r' setup /passive', admin=True)

    # 等待安裝完成
    args = ['tasklist', '/fi', 'imagename eq vcredist_x64.exe', '/fo', 'csv']
    output = cmd_exec(args)
    while output.count('\r\n') == 2:
        time.sleep(0.5)
        output = cmd_exec(args)

    # 移除安裝包
    os.remove(vcdist)

def remove_vcredist():
    """
    TODO: 移除 Visual C++ 2010 x64 Redistributable 10.0.40219.325
    這部分先確定是否會阻斷, 如果不利自動化就放棄實作
    """

def verof_skcom():
    """
    檢查群益 API 元件是否已註冊
    """
    cmd = r'reg query HKLM\SOFTWARE\Classes\TypeLib /s /f SKCOM.dll'
    result = ps_exec(cmd)

    skcom_ver = '0.0.0.0'
    if result is not None:
        lines = result.split('\r\n')
        for line in lines:
            # 找 DLL 檔案路徑
            match = re.match(r'.+REG_SZ\s+(.+SKCOM.dll)', line)
            if match is not None:
                # 取檔案摘要內容裡版本號碼
                dll_path = match.group(1)
                fso = comtypes.client.CreateObject('Scripting.FileSystemObject')
                try:
                    skcom_ver = fso.GetFileVersion(dll_path)
                    # skcom_ver = fso.GetFileVersion(r'C:\makeexception.txt')
                except COMError:
                    pass

    return version.parse(skcom_ver)

def install_skcom(install_ver):
    """
    安裝群益 API 元件
    """
    url = 'https://www.capital.com.tw/Service2/download/api_zip/CapitalAPI_%s.zip' % install_ver

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
            if re.search(r'元件/x64/.+\.dll$', name950):
                dest_path = r'%s\%s' % (com_path, name950.split('/')[-1])
                with archive.open(name437, 'r') as cmpf, \
                     open(dest_path, 'wb') as extf:
                    extf.write(cmpf.read())

    # 註冊元件
    cmd = r'regsvr32 %s\SKCOM.dll' % com_path
    ps_exec(cmd, admin=True)

    return True

def remove_skcom():
    """
    TODO: 解除註冊與移除 skcom 元件
    """
    com_path = os.path.expanduser(r'~\.skcom\lib')
    com_file = com_path + r'\SKCOM.dll'
    if not os.path.isfile(com_file):
        return

    console_print('移除群益 API 元件')
    console_print('  路徑: ' + com_path)
    console_print('  解除註冊: ' + com_file)
    cmd = 'regsvr32 /u ' + com_file
    ps_exec(cmd, admin=True)

    console_print('  移除元件目錄')
    shutil.rmtree(com_path)

def has_valid_mod():
    r"""
    檢測 comtypes\gen\SKCOMLib.py 是否已生成, 如果已生成則檢查是否連結到正確的 dll 檔

    關於例外
      import comtypes.gen.SKCOMLib as sk
      ImportError: No module named 'comtypes.gen.SKCOMLib'
    原因是 COM 對應的 comtypes\gen\{UUID}.py 檔案尚未產生
      ...\site-packages\comtypes\gen\SKCOMLib.py
      ...\site-packages\comtypes\gen\_75AAD71C_8F4F_4F1F_9AEE_3D41A8C9BA5E_0_1_0.py
    如果要重現這個問題, 只要把 comtypes 移除再安裝就會出現
      pip uninstall comtypes
      pip install comtypes
    """
    result = False

    pkgdirs = site.getsitepackages()
    for pkgdir in pkgdirs:
        if pkgdir.endswith('site-packages'):
            name_mod = pkgdir + r'\comtypes\gen\SKCOMLib.py'
            # uuid_mod = pkgdir + r'\comtypes\gen\_75AAD71C_8F4F_4F1F_9AEE_3D41A8C9BA5E_0_1_0.py'
            break

    if os.path.isfile(name_mod):
        result = True
        # pylint: disable=pointless-string-statement
        '''
        dll_path = os.path.expanduser(r'~\\.skcom\\lib\\SKCOM.dll')
        import comtypes.gen.SKCOMLib as sk
        if sk.typelib_path == dll_path:
            result = True
        else:
            console_print('移除 site-packages\\comtypes\\gen\\SKCOMLib.py, 因為連結的 DLL 不正確')
            del comtypes.gen.SKCOMLib
            del comtypes.gen._75AAD71C_8F4F_4F1F_9AEE_3D41A8C9BA5E_0_1_0
            os.remove(name_mod)
            os.remove(uuid_mod)

        TODO: 如果跑了這一段, 會導致接下來的 comtypes.client.GetModule(dll_path) 沒有作用,
              使用者會誤以為安裝完成, 在找到解決辦法之前, 先不導入這個設計
        '''

    return result

def generate_mod():
    r"""
    產生 COM 元件的 comtypes.gen 對應模組
    """
    console_print(r'生成 site-packages\comtypes\gen\SKCOMLib.py')
    dll_path = os.path.expanduser(r'~\.skcom\lib\SKCOM.dll')
    comtypes.client.GetModule(dll_path)

def clean_mod():
    r"""
    TODO: 清除已產生的 site-packages\comtypes\gen\*.py
    """
    pkgdirs = site.getsitepackages()
    for pkgdir in pkgdirs:
        if not pkgdir.endswith('site-packages'):
            continue

        gendir = pkgdir + r'\comtypes\gen'
        if not os.path.isdir(gendir):
            continue

        console_print('移除 comtypes 套件自動生成檔案')
        console_print('  路徑 ' + gendir)

        for item in os.listdir(gendir):
            if item.endswith('.py'):
                console_print('  移除 %s' % item)
                os.remove(gendir + '\\' + item)

        cache_dir = gendir + '\\' + '__pycache__'
        if os.path.isdir(cache_dir):
            console_print('  移除 __pycache__')
            shutil.rmtree(cache_dir)

def download_file(url, save_path):
    """
    使用 8K 緩衝下載檔案
    """
    abs_path = check_dir(save_path)
    file_path = r'%s\%s' % (abs_path, url.split('/')[-1])

    with requests.get(url, stream=True) as resp:
        resp.raise_for_status()
        with open(file_path, 'wb') as dlf:
            for chunk in resp.iter_content(chunk_size=8192):
                if chunk:
                    dlf.write(chunk)
            dlf.flush()

    return file_path

def check_dir(usr_path):
    """
    檢查目錄, 不存在就建立目錄, 完成後回傳絕對路徑
    """
    rel_path = os.path.expanduser(usr_path)
    if not os.path.isdir(rel_path):
        os.makedirs(rel_path)
    abs_path = os.path.realpath(rel_path)
    return abs_path

def get_dll_abs_path():
    """
    取得 SKCOM.dll 絕對路徑
    """
    return os.path.expanduser(r'~\.skcom\lib\SKCOM.dll')
