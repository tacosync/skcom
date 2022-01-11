"""
skcom.helper
"""
# pylint: disable=broad-except, bare-except, pointless-string-statement

import logging
import logging.config
import os.path
import platform
import random
import re
import shutil
import site
import subprocess
import winreg
import zipfile
from getpass import getpass

import comtypes.client
import win32com.client
import requests
from packaging import version
from requests.exceptions import ConnectionError as RequestsConnectionError
import yaml
import base64
import busm

from skcom.crypto import decrypt_text
from skcom.exception import ShellException
from skcom.exception import NetworkException
from skcom.exception import InstallationException
from skcom.exception import ConfigException

def powershell_base64(script):
    '''
    Powershell Script 轉換為 Start-Process -EncodedCommand 參數
    '''
    return base64.b64encode(script.encode('utf-16le')).decode('ascii')

def powershell_start(cmd, admin_priv=False):
    '''
    在 PowerShell 環境以 Start-Process 執行應用程式
    '''
    # 參數加一層引號, 避免空白字元分割
    quoted_args = list(map(lambda a: '"%s"' % a if ' ' not in a else '"`"%s`""' % a, cmd[1:]))

    # 生成 stdout, stderr 暫存檔
    existed = True
    while existed:
        fsn = random.randint(0, 65535)
        stdout_path = os.path.expanduser(r'~\%04x-stdout.txt' % fsn)
        stderr_path = os.path.expanduser(r'~\%04x-stderr.txt' % fsn)
        exitcode_path = os.path.expanduser(r'~\%04x-exitcode.txt' % fsn)
        existed = os.path.isfile(stdout_path) or os.path.isfile(stderr_path) or os.path.isfile(exitcode_path)

    try:
        open(stdout_path, 'w').close()
        open(stderr_path, 'w').close()
        open(exitcode_path, 'w').close()
    except IOError:
        raise ShellException(-1, '無法產生 stdout 暫存檔')

    # 組合 PowerShell 指令
    # 注意!! -Verb 與下列參數不能並用
    # * -NoNewWindow
    # * -RedirectStandardOutput
    # * -RedirectStandardError
    cmd_sp = 'Start-Process "%s" ' % cmd[0]
    if len(quoted_args) > 0:
        cmd_sp += '-ArgumentList $arglist '
    cmd_sp += '-RedirectStandardOutput "%s" ' % stdout_path
    cmd_sp += '-RedirectStandardError "%s" ' % stderr_path
    cmd_sp += '-Wait -NoNewWindow'
    cmd_ex = 'Write-Output $LastExitCode > "%s"' % exitcode_path
    if len(quoted_args) > 0:
        cmd_arg = '$arglist = @(%s)' % ','.join(quoted_args)
        ps_script = [cmd_arg, cmd_sp, cmd_ex]
    else:
        ps_script = [cmd_sp, cmd_ex]

    # 指令以 unicode 編碼後轉換為 base64
    base64_script = powershell_base64('\n'.join(ps_script))

    # 以管理員身分執行時, 需要多一層 powershell 啟動
    if admin_priv:
        base64_script = powershell_base64('Start-Process powershell -ArgumentList "-EncodedCommand %s" -Wait -Verb RunAs' % base64_script)

    # TODO: 如果系統有安裝 Powershell 6+ 的版本, 優先使用新版 (pwsh.exe)
    ps_command = 'powershell -EncodedCommand ' + base64_script

    # TODO: 需要加強錯誤處理以及回傳值蒐集
    try:
        comp = subprocess.run(ps_command)
    except Exception as ex:
        print(ex)

    # 讀取 stdout, stderr, exitcode
    # 檔案開頭會多出一個 0xfeff 字元需要跳過
    stdout = ''
    stderr = ''
    exitcode = ''
    with open(stdout_path, 'r') as stdout_file:
        stdout = stdout_file.read()
    with open(stderr_path, 'r') as stderr_file:
        stderr = stderr_file.read()
    with open(exitcode_path, 'r', encoding='utf-16le') as exitcode_file:
        exitcode = exitcode_file.read()[1:]

    # 移除暫存檔
    os.remove(stdout_path)
    os.remove(stderr_path)
    os.remove(exitcode_path)

    # TODO: stdout 擷取
    # * 實測 ping unknown, 錯誤時會輸出在 stdout 而不是在 stderr, 不論是否以管理員模式執行都是如此
    # * 實測 tasklist /FI "IMAGENAME eq unknown.exe", 錯誤時會輸出在 stdout 而不是在 stderr, 不論是否以管理員模式執行都是如此
    # TODO: stderr 擷取
    # * 目前測不出實際作用
    # TODO: 研究 $LastExitCode 用法
    # * 實測發現自己執行 Write-Output $LastExitCode 有值, 腳本執行則沒有值, 所以暫時無法得到回傳值
    # print('STDOUT:')
    # print(stdout)
    # print('STDERR:')
    # print(stderr)
    # print('RETURN:')
    # print(exitcode)

def pack_arglist(args):
    """
    處理 Start-Process -ArgumentList 的參數內容
    """
    packed = list(map('"{}"'.format, args))
    packed = ' '.join(packed)
    packed = "'{}'".format(packed)
    return packed

def reg_read_value(node, root=winreg.HKEY_LOCAL_MACHINE):
    """
    讀取單一值
    """
    (key, name) = node.split(':')
    handle = winreg.OpenKey(root, key)
    (value, _) = winreg.QueryValueEx(handle, name)
    winreg.CloseKey(handle)
    return value

def reg_list_value(key, root=winreg.HKEY_LOCAL_MACHINE):
    """
    列舉機碼下的所有值
    """
    i = 0
    values = {}
    handle = winreg.OpenKey(root, key)

    while True:
        try:
            (vname, value, _) = winreg.EnumValue(handle, i)
            if vname == '':
                vname = '(default)'
            values[vname] = value
            i += 1
        except OSError:
            # winreg.EnumValue(handle, i) 的 i 超出範圍
            break

    winreg.CloseKey(handle)
    return values

def reg_find_value(key, value, root=winreg.HKEY_LOCAL_MACHINE):
    """
    遞迴搜尋機碼下的值, 回傳一組 tuple (所在位置, 數值)
    字串資料採用局部比對, 其餘型態採用完整比對
    """
    i = 0
    handle = winreg.OpenKey(root, key)
    vtype = type(value)

    while True:
        try:
            values = reg_list_value(key)
            for vname in values:
                # 除錯用
                # if key == r'SOFTWARE\Classes\TypeLib\{75AAD71C-8F4F-4F1F-9AEE-3D41A8C9BA5E}\1.0\0\win32':
                #     print(vname, values[vname], value)

                # 型態不同忽略
                if not isinstance(values[vname], vtype):
                    continue

                leaf_node = key + ':' + vname
                if isinstance(value, str):
                    # 字串採用局部比對, 不分大小寫
                    if value.lower() in values[vname].lower():
                        return (leaf_node, values[vname])
                else:
                    # 其餘型態採用完整比對
                    if value == values[vname]:
                        return (leaf_node, values[vname])

            subkey = key + '\\' + winreg.EnumKey(handle, i)
            result = reg_find_value(subkey, value)
            if result[0] != '':
                return result

            i += 1
        except OSError:
            # winreg.EnumKey(handle, i) 的 i 超出範圍
            break

    winreg.CloseKey(handle)
    return ('', '')

def verof_vcredist():
    """
    取得 Visual C++ 2010 可轉發套件版本資訊
    注意!! 即使套件沒安裝也不可以噴例外, 否則會中斷安裝流程
    """
    try:
        (width, _) = platform.architecture()
        if width == '32bit':
            node = r'SOFTWARE\WOW6432Node\Microsoft\VisualStudio\10.0\VC\VCRedist\x86:Version'
        else:
            node = r'SOFTWARE\WOW6432Node\Microsoft\VisualStudio\10.0\VC\VCRedist\x64:Version'
        reg_val = reg_read_value(node)
        pkg_ver = reg_val.strip('v')
        if not re.match(r'(\d+\.){3}\d+', pkg_ver):
            pkg_ver = '0.0.0.0'
    except FileNotFoundError:
        pkg_ver = '0.0.0.0'

    return version.parse(pkg_ver)

def install_vcredist():
    """
    安裝 Visual C++ 2010 Redistributable 10.0.40219.325
    """
    # pylint: disable=invalid-name
    DEBUG_MODE = False

    # 下載 (失敗會觸發 NetworkException)
    (width, _) = platform.architecture()
    url = 'https://download.microsoft.com/download/1/6/5/165255E7-1014-4D0A-B094-B6A430A6BFFC/'
    if width == '32bit':
        url += 'vcredist_x86.exe'
    else:
        url += 'vcredist_x64.exe'
    vcexe = download_file(url, check_dir('~/.skcom'))

    # 安裝 (安裝程式會自動切換系統管理員身分)
    # https://www.microsoft.com/en-US/download/details.aspx?id=26999
    cmd = [vcexe, 'setup', '/passive']
    powershell_start(cmd, True)

    # 移除安裝包, 暫不處理檔案系統例外
    if not DEBUG_MODE:
        os.remove(vcexe)

    # 版本檢查
    expected_ver = version.parse('10.0.40219.325')
    current_ver = verof_vcredist()
    if current_ver != expected_ver:
        raise InstallationException('Visual C++ 2010 可轉發套件', expected_ver)

    return current_ver

def remove_vcredist():
    """
    移除 Visual C++ 2010 x64 Redistributable  x64 - 10.0.40219.325

    其他作法與不採用原因:
      * 方法1: Get-CimInstance ... | Invoke-CimMethod ...
        * 缺點: 無法自動切換系統管理員身分
      * 方法2: Uninstall-Package ...
        * 缺點: UAC 對話方塊不會聚焦
    """
    (width, _) = platform.architecture()
    if width == '32bit':
        uuid = '/X{F0C3E5D1-1ADE-321E-8167-68EF0DE699A5}'
    else:
        uuid = '/X{1D8E6291-B0D5-35EC-8441-6616F567A0F7}'
    cmd = ['msiexec.exe', uuid, '/passive']
    powershell_start(cmd)

def verof_skcom():
    """
    檢查群益 API 元件是否已註冊
    注意!! 即使套件沒安裝也不可以噴例外, 否則會中斷安裝流程
    """
    skcom_ver = '0.0.0.0'
    try:
        (_, dll_path) = reg_find_value(r'SOFTWARE\Classes\TypeLib', 'SKCOM.dll')
        if dll_path != '':
            fso = win32com.client.Dispatch('Scripting.FileSystemObject')
            skcom_ver = fso.GetFileVersion(dll_path)
        else:
            # TODO: 找不到 SKCOM.dll 字串 (先前因為區分大小寫會發生)
            pass
    except Exception as ex:
        # TODO: 忘了什麼時候會炸掉
        pass

    return version.parse(skcom_ver)

def install_skcom(install_ver):
    """
    安裝群益 API 元件
    """
    url = 'https://www.capital.com.tw/Service2/download/api_zip/CapitalAPI_%s.zip' % install_ver

    # 建立元件目錄
    # 暫不處理檔案系統例外
    com_path = check_dir(r'~\.skcom\lib')

    # 下載
    # 如果下載失敗, 會觸發 NetworkException
    file_path = download_file(url, com_path)

    # 解壓縮
    # 只解壓縮 64-bits DLL 檔案, 其他非必要檔案不處理
    # 讀檔時需要用 CP437 編碼處理檔名, 寫檔時需要用 CP950 處理檔名
    # 暫不處理解壓縮例外
    (width, _) = platform.architecture()
    if width == '32bit':
        pattern = r'元件/x86/.+\.dll$'
    else:
        pattern = r'元件/x64/.+\.dll$'
    with zipfile.ZipFile(file_path, 'r') as archive:
        for name437 in archive.namelist():
            name950 = name437.encode('cp437').decode('cp950')
            if re.search(pattern, name950):
                dest_path = r'%s\%s' % (com_path, name950.split('/')[-1])
                with archive.open(name437, 'r') as cmpf, \
                     open(dest_path, 'wb') as extf:
                    extf.write(cmpf.read())

    # 註冊元件
    # TODO: 使用者拒絕提供系統管理員權限時, 會觸發 ShellException,
    #       但是登錄檔依然能寫入成功, 原因需要再調查
    cmd = ['regsvr32', '/s', r'%s\SKCOM.dll' % com_path]
    powershell_start(cmd, admin_priv=True)

    # 版本檢查
    expected_ver = version.parse(install_ver)
    current_ver = verof_skcom()
    if current_ver != expected_ver:
        raise InstallationException('群益 API COM 元件', expected_ver)

    return current_ver

def remove_skcom():
    r"""
    解除註冊與移除 skcom 元件
    注意!! 解除註冊後, 這些登錄檔的值仍然會殘留:

    * HKLM\SOFTWARE\Classes\TypeLib\{75AAD71C-8F4F-4F1F-9AEE-3D41A8C9BA5E}
    * HKLM\SOFTWARE\Classes\WOW6432Node\TypeLib\{75AAD71C-8F4F-4F1F-9AEE-3D41A8C9BA5E}
    * HKLM\SOFTWARE\WOW6432Node\Classes\TypeLib\{75AAD71C-8F4F-4F1F-9AEE-3D41A8C9BA5E}
    * HKLM\SOFTWARE\Classes\CLSID\{54FE0E28-89B6-43A7-9F07-BE988BB40299}
    * HKLM\SOFTWARE\Classes\CLSID\{72D98963-03E9-42AB-B997-BB2E5CCE78DD}
    * HKLM\SOFTWARE\Classes\CLSID\{853EC706-F437-46E2-80E0-896901A5B490}
    * HKLM\SOFTWARE\Classes\CLSID\{AC30BAB5-194A-4515-A8D3-6260749F8577}
    * HKLM\SOFTWARE\Classes\CLSID\{E3CB8A7C-896F-4828-85FC-8975E56BA2C4}
    * HKLM\SOFTWARE\Classes\CLSID\{E7BCB8BB-E1F0-4F6F-A944-2679195E5807}

    殘留值可能會導致下次安裝時誤判已註冊
    """
    com_path = os.path.expanduser(r'~\.skcom\lib')
    com_file = com_path + r'\SKCOM.dll'
    if not os.path.isfile(com_file):
        return

    logger = logging.getLogger('helper')
    logger.info('  路徑: %s', com_path)
    logger.info('  解除註冊: %s', com_file)
    cmd = ['regsvr32', '/s', '/u', com_file]
    powershell_start(cmd, admin_priv=True)

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
    """
    產生 COM 元件的 comtypes.gen 對應模組
    """
    logger = logging.getLogger('helper')
    logger.info(r'生成 site-packages\comtypes\gen\SKCOMLib.py')
    dll_path = os.path.expanduser(r'~\.skcom\lib\SKCOM.dll')
    comtypes.client.GetModule(dll_path)

def clean_mod():
    r"""
    清除已產生的 site-packages\comtypes\gen\*.py
    """
    pkgdirs = site.getsitepackages()
    for pkgdir in pkgdirs:
        if not pkgdir.endswith('site-packages'):
            continue

        gendir = pkgdir + r'\comtypes\gen'
        if not os.path.isdir(gendir):
            continue

        logger = logging.getLogger('helper')
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

    try:
        with requests.get(url, stream=True) as resp:
            resp.raise_for_status()
            with open(file_path, 'wb') as dlf:
                for chunk in resp.iter_content(chunk_size=8192):
                    if chunk:
                        dlf.write(chunk)
                dlf.flush()
    except RequestsConnectionError:
        # 拔網路線可以測試這段
        raise NetworkException('無法下載: {}'.format(url))

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

def reset_logging(cfg_skcom=None):
    """
    重新載入 logging 設定
    """
    # pylint: disable=unused-argument
    cfg_logging_path = '{}/conf/logging.yaml'.format(os.path.dirname(__file__))
    with open(cfg_logging_path, 'r', encoding='utf-8') as cfg_logging_file:
        cfg_logging = yaml.load(cfg_logging_file, Loader=yaml.SafeLoader)

        # 自動替換 ~ 為家目錄, 以及自動產生 log 目錄
        for name in cfg_logging['handlers']:
            handler = cfg_logging['handlers'][name]
            if 'filename' in handler:
                if handler['filename'].startswith('~'):
                    handler['filename'] = os.path.expanduser(handler['filename'])
                dirname = os.path.dirname(handler['filename'])
                if not os.path.isdir(dirname):
                    os.makedirs(dirname)

        logging.config.dictConfig(cfg_logging)

def load_config():
    """
    載入設定檔
    """
    cfg_path = os.path.expanduser(r'~\.skcom\skcom.yaml')
    enc_path = cfg_path + '.enc'
    config = None
    load_plain = False
    logger = logging.getLogger('helper')
    message = ''

    # 嘗試讀取加密設定
    try:
        with open(enc_path, 'rb') as enc_file:
            secret = enc_file.read()
            password = getpass('請輸入設定檔密碼: ')
            plain = decrypt_text(secret, password)
            config = yaml.load(plain, Loader=yaml.SafeLoader)
        logger.info('已載入加密設定')
        logger.info('如果需要變更設定檔, 執行下列指令可以解密:')
        logger.info('  python -m skcom.tools.cfdec')
    except FileNotFoundError as ex:
        load_plain = True
    except Exception as ex:
        # TODO: 這裡掛掉的可能性很多種, 有改善空間
        message = str(ex)

    # 嘗試讀取明文設定
    if load_plain:
        try:
            with open(cfg_path, 'r', encoding='utf-8') as cfg_file:
                config = yaml.load(cfg_file, Loader=yaml.SafeLoader)
            logger.warning('目前設定檔沒有加密, 建議您加密避免帳號外流')
            logger.warning('執行下列指令即可加密:')
            logger.warning('  python -m skcom.tools.cfenc')
        except FileNotFoundError as ex:
            # 複製設定檔模板
            tpl_path = os.path.dirname(os.path.realpath(__file__)) + r'\conf\skcom.yaml'
            shutil.copy(tpl_path, cfg_path)
            with open(cfg_path, 'r', encoding='utf-8') as cfg_file:
                config = yaml.load(cfg_file, Loader=yaml.SafeLoader)
        except Exception as ex:
            message = str(ex)

    if config is not None:
        # 檢查設定檔是不是沒改過的模板
        if config['account'] != 'A123456789':
            logger = logging.getLogger('bot')
            try:
                token = config['telegram']['token']
                master = config['telegram']['master']
                bh = busm.BusmHandler()
                bh.setup_telegram(token, master)
                logger.addHandler(bh)
            except Exception as ex:
                logger.error('Cannot setup BusmHandler.')
                logger.error(ex)
        else:
            logger.warning('請開啟設定檔, 將帳號密碼改為您的證券帳號')
            logger.warning('設定檔路徑: %s', cfg_path)
            raise ConfigException('設定檔尚未修改', loaded=True)
    else:
        # 檢查設定檔是否載入失敗
        raise ConfigException('設定檔讀取失敗:\n%s' % message)

    return config
