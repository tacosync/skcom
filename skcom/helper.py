"""
quicksk.helper
"""

import logging
import os.path
import random
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
from skcom.exception import ShellException

def win_exec(cmd, admin_priv=False):
    charset = 'cp950'

    # ~ 自動轉換家目錄
    path = os.path.expanduser(cmd[0])
    cmd[0] = path

    if admin_priv:
        # 產生隨機檔名, 用來儲存 stdout
        existed = True
        while existed:
            fsn = random.randint(0, 65535)
            stdout_path = os.path.expanduser(r'~\stdout-%04x.txt' % fsn)
            existed = os.path.isfile(stdout_path)
        try:
            open(stdout_path, 'w').close()
        except IOError:
            return (-1, '')

        deep_args = pack_arglist(cmd[1:])

        # 產生底層指令與表層參數
        surface_args = [
            'Start-Process',
            '-FilePath', cmd[0],
            '-ArgumentList', deep_args,
            '-RedirectStandardOutput', stdout_path,
            '-NoNewWindow'
        ]

        # 表層參數加上雙引號套入 ArgumentList
        # TODO: 底層參數的雙引號逸脫
        surface_args = ','.join(surface_args)

        # 產生表層指令
        cmd = [
            'powershell.exe', 'Start-Process',
            '-FilePath', 'powershell.exe',
            '-ArgumentList', surface_args,
            '-Verb', 'RunAs',
            '-Wait'
        ]

        # 取 stdout
        stdout = ''
        comp = subprocess.run(cmd, check=True)
        with open(stdout_path, 'r') as stdout_file:
            stdout = stdout_file.read()

        # 移除暫存檔
        os.remove(stdout_path)
    else:
        # 執行指令
        comp = subprocess.run(cmd, shell=True, capture_output=True)
        if comp.returncode != 0:
            try:
                stderr = comp.stderr.decode(charset)
            except:
                stderr = 'Cannot decode stderr.'
            raise ShellException(comp.returncode, stderr)

        # stdout 解碼
        try:
            stdout = comp.stdout.decode(charset)
        except:
            raise ShellException(-1, 'Cannot decode stdout')

    return stdout

def pack_arglist(args):
    """
    處理 Start-Process -ArgumentList 的參數內容
    """
    packed = list(map(lambda arg: '"{}"'.format(arg), args))
    packed = ' '.join(packed)
    packed = "'{}'".format(packed)
    return packed

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

    # 下載與安裝 (安裝程式會自動切換系統管理員身分)
    url = 'https://download.microsoft.com/download/1/6/5/' + \
          '165255E7-1014-4D0A-B094-B6A430A6BFFC/vcredist_x64.exe'
    vcexe = download_file(url, check_dir('~/.skcom'))
    cmd = [vcexe, 'setup', '/passive']
    win_exec(cmd)

    # 等待安裝完成
    cmd = ['tasklist', '/fi', 'imagename eq vcredist_x64.exe', '/fo', 'csv']
    stdout = win_exec(cmd)
    while stdout.count('\r\n') == 2:
        print(stdout)
        time.sleep(0.5)
        stdout = win_exec(cmd)
    print(stdout)

    # 移除安裝包
    os.remove(vcexe)

def remove_vcredist():
    """
    TODO: 移除 Visual C++ 2010 x64 Redistributable 10.0.40219.325
    這部分先確定是否會阻斷, 如果不利自動化就放棄實作
    """

def verof_skcom():
    """
    檢查群益 API 元件是否已註冊
    """
    skcom_ver = '0.0.0.0'

    # TODO: 比照 install_vcredist 用 winreg 處理
    cmd = ['reg', 'query', r'HKLM\SOFTWARE\Classes\TypeLib', '/s', '/f', 'SKCOM.dll']

    try:
        stdout = win_exec(cmd)
        lines = stdout.split('\n')
        for line in lines:
            # 找 DLL 檔案路徑
            match = re.match(r'.+REG_SZ\s+(.+SKCOM.dll)', line)
            # print(match)
            if match is not None:
                # 取檔案摘要內容裡版本號碼
                dll_path = match.group(1)
                fso = comtypes.client.CreateObject('Scripting.FileSystemObject')
                skcom_ver = fso.GetFileVersion(dll_path)
                # skcom_ver = fso.GetFileVersion(r'C:\makeexception.txt')
    except subprocess.CalledProcessError:
        pass
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
    cmd = ['regsvr32', r'%s\SKCOM.dll' % com_path]
    win_exec(cmd, admin_priv=True)

    return True

def remove_skcom():
    """
    TODO: 解除註冊與移除 skcom 元件
    """
    com_path = os.path.expanduser(r'~\.skcom\lib')
    com_file = com_path + r'\SKCOM.dll'
    if not os.path.isfile(com_file):
        return

    logger = logging.getLogger('helper')
    logger.info('移除群益 API 元件')
    logger.info('  路徑: %s', com_path)
    logger.info('  解除註冊: %s', com_file)
    cmd = ['regsvr32', '/u', com_file]
    win_exec(cmd, admin_priv=True)

    logger.info('  移除元件目錄')
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
            logger.info('移除 site-packages\\comtypes\\gen\\SKCOMLib.py, 因為連結的 DLL 不正確')
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
    logger = logging.getLogger('helper')
    logger.info(r'生成 site-packages\comtypes\gen\SKCOMLib.py')
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

        logger = logging.getLogger('helper')
        logger.info('移除 comtypes 套件自動生成檔案')
        logger.info('  路徑 %s', gendir)

        for item in os.listdir(gendir):
            if item.endswith('.py'):
                logger.info('  移除 %s', item)
                os.remove(gendir + '\\' + item)

        cache_dir = gendir + '\\' + '__pycache__'
        if os.path.isdir(cache_dir):
            logger.info('  移除 __pycache__')
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
