import re
import subprocess

import comtypes.client

def run_command(cmd):
    completed = subprocess.run(cmd.split(' '), shell=True, capture_output=True)
    if completed.returncode == 0:
        return completed.stdout.decode('cp950')
    else:
        return None

## 檢查 Visual Studio 2010 可轉發套件是否已安裝
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
    result = run_command(cmd)

    version = None
    if result is not None:
        lines = result.split('\r\n')
        for line in lines:
            m = re.match(r'.+REG_SZ\s+(.+SKCOM.dll)', line)
            if m is not None:
                dll_path = m.group(1)
                fso = comtypes.client.CreateObject('Scripting.FileSystemObject')
                version = fso.GetFileVersion(dll_path)

    return version
