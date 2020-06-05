import os
import logging
from getpass import getpass

from skcom.crypto import encrypt_text

def main():
    logger = logging.getLogger('helper')
    cfg_path = os.path.expanduser(r'~\.skcom\skcom.yaml')

    try:
        # 加密
        secret = b''
        with open(cfg_path, 'r', encoding='utf-8') as cfg_file:
            plain = cfg_file.read()
            password = getpass('請輸入密碼, 至少 8 個字元: ')
            if len(password) < 8:
                raise Exception('密碼長度太短, 取消加密作業')
            passchk = getpass('   再輸入一次密碼進行確認: ')
            if password != passchk:
                raise Exception('確認失敗, 取消加密作業')
            secret = encrypt_text(plain, password)

        # 儲存加密設定檔, 刪除明文設定檔
        with open(cfg_path + '.enc', 'wb') as enc_file:
            enc_file.write(secret)
            os.remove(cfg_path)

        logger.info('加密完成')
    except Exception as ex:
        logger.error(ex)

if __name__ == '__main__':
    main()
