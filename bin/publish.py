#!/usr/bin/env python3
#
# 必要工具套件:
# * pylint
# * setuptools
# * wheel
# * virtualenv
# * twine
# * pyenv or pyenv-win

import configparser
import os
import re
import shutil
import subprocess
import sys
import platform

def get_wheel(platname='any'):
    """ 製作 wheel 檔案與取得檔名 """
    wheel = False
    comp = subprocess.run(
        [
            'python', 'setup.py', 'bdist_wheel',
            '--plat-name', platname
        ],
        check=True,
        capture_output=True
    )
    stdout = comp.stdout.decode('utf-8').split('\n')
    for line in stdout:
        # 留意 Windows 斜線
        # creating 'dist\busm-0.9.4-py3-none-any.whl' and adding 'build\bdist.win-amd64\wheel' to it
        match = re.search(r'dist[\\/].+\.whl', line)
        if match is not None:
            wheel = match.group(0)
            break

    return wheel

def get_installed_python():
    """ 選擇 pyenv 環境中 3.6 ~ 3.8 的最新版本 """
    detected = {
        '3.6': { 'patch': -1, 'suffix': '' },
        '3.7': { 'patch': -1, 'suffix': '' },
        '3.8': { 'patch': -1, 'suffix': '' }
    }

    # 偵測已安裝版本, 同一個 minor 版採用最新的 patch 版
    # TODO: 這段在 Windows 會噴 "The batch file cannot be found." 不過實際上沒什麼大礙
    comp = subprocess.run(['pyenv', 'versions'], shell=True, check=True, capture_output=True)
    stdout = comp.stdout.decode('utf-8').strip().split('\n')
    for line in stdout:
        #   3.7.6
        # * 3.8.1 (...)
        # * 3.8.1-amd64 (...)
        match = re.search(r'^[^\d]*(\d\.\d)\.(\d+)([^\s]*)', line)
        if match is not None:
            minor = match[1]
            patch = int(match[2])
            suffix = match[3]
            if minor in detected and \
                patch > detected[minor]['patch']:
                detected[minor]['patch'] = patch
                detected[minor]['suffix'] = suffix

    # 偵測到的清單轉換為陣列
    selected = []
    for minor in detected:
        if detected[minor]['patch'] != -1:
            full_ver = '%s.%s%s' % (minor, detected[minor]['patch'], detected[minor]['suffix'])
            selected.append(full_ver)

    return selected

def test_in_virtualenv(pyver, wheel):
    """ 配置 virtualenv 並執行測試程式 """

    # 安裝 virtualenv
    if platform.system() == 'Windows':
        src = os.path.expanduser(r'~\.pyenv\pyenv-win\versions\%s\python.exe' % pyver)
    else:
        src = os.path.expanduser('~/.pyenv/versions/%s/bin/python' % pyver)
    dst = 'sandbox/%s' % pyver
    comp = subprocess.run(
        ['python', '-m', 'virtualenv', '-p', src, dst],
        check=True,
        shell=True,
        # capture_output=True
    )

    # 測試 wheel 包
    # 1. 安裝 wheel
    # 2. 安裝 green
    # 3. 執行測試程式
    wheel = '../../' + wheel
    if platform.system() == 'Windows':
        pip = r'Scripts\pip.exe'
    else:
        pip = 'bin/pip'
    comp = subprocess.run(
        [pip, 'install', wheel, 'green'],
        cwd=dst,
        check=True,
        shell=True,
        # capture_output=True
    )

    """
    if platform.system() == 'Windows':
        green = r'Scripts\green.exe'
    else:
        green = 'bin/pip'
    comp = subprocess.run(
        [green, '-vv', 'twnews'],
        cwd=dst,
        check=True,
        shell=True,
        capture_output=True
    )
    """

def wheel_check():
    """ 檢查 wheel 是否能正常運作在各個 Python 版本環境上 """

    print('檢查程式碼品質')
    ret = os.system('python -m pylint -f colorized busm')
    if ret != 0:
        print('檢查沒通過，停止封裝')
        exit(ret)

    #print('檢查 README.md')
    # TODO:

    print('偵測可用的測試環境')
    if os.path.isdir('sandbox'):
        shutil.rmtree('sandbox')

    wheel32 = get_wheel('win32')
    wheel64 = get_wheel('win_amd64')
    installed_py = get_installed_python()
    if len(installed_py) == 0:
        print('沒有任何可用的測試環境')
        exit(1)

    for pyver in installed_py:
        print('-', pyver)
    print('')

    for pyver in installed_py:
        print('測試 Python %s' % pyver)
        if pyver.endswith('-amd64'):
            test_in_virtualenv(pyver, wheel64)
        else:
            test_in_virtualenv(pyver, wheel32)
        print('')

def upload_to_pypi(test=False):
    """ 上傳 wheel 到 PyPi """

    # 檢查 ~/.pypirc 是否存在
    if not os.path.isfile(os.path.expanduser('~/.pypirc')):
        print('缺少 pypi 設定檔 ~/.pypirc')
        print('參考: https://gist.github.com/ibrahim12/c6a296c1e8f409dbed2f')

    # 重新產生 wheel 與上傳前確認
    wheel32 = get_wheel('win32')
    wheel64 = get_wheel('win_amd64')
    print('準備上傳兩個檔案:')
    print('  -', wheel32)
    print('  -', wheel64)
    ans = input('確定上傳嗎 [y/n]? ')
    if ans != 'y':
        print('取消上傳')
        return

    # 上傳 wheel
    cmd = ['python', '-m', 'twine', 'upload']
    if test:
        cmd.append('--repository')
        cmd.append('testpypi')
    cmd.append('--verbose')
    cmd.append(wheel32)
    comp = subprocess.run(cmd)
    if comp.returncode == 0:
        print(wheel32, '上傳成功')
    else:
        print(wheel32, '上傳失敗')

    cmd.pop()
    cmd.append(wheel64)
    comp = subprocess.run(cmd)
    if comp.returncode == 0:
        print(wheel64, '上傳成功')
    else:
        print(wheel64, '上傳失敗')

def main():
    # 確保不在 repo 目錄也能正常執行
    HOME = os.path.realpath(os.path.dirname(__file__) + '/..')
    os.chdir(HOME)

    action = 'wheel'
    if len(sys.argv) > 1:
        action = sys.argv[1]

    try:
        if action == 'release':
            upload_to_pypi()
        elif action == 'test':
            upload_to_pypi(True)
        elif action == 'wheel':
            wheel_check()
        else:
            print('Unknown action "%s".' % action)
            exit(1)
    except subprocess.CalledProcessError as ex:
        print(ex)

if __name__ == '__main__':
    main()
