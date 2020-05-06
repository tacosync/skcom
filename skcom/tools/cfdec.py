import os
import logging
from getpass import getpass

from skcom.crypto import decrypt_text

def main():
    logger = logging.getLogger('helper')
    cfg_path = os.path.expanduser(r'~\.skcom\skcom.yaml')

    try:
        # 解密
        with open(cfg_path + '.enc', 'rb') as enc_file:
            secret = enc_file.read()
            password = getpass('請輸入設定檔密碼: ')
            plain = decrypt_text(secret, password)

        # 儲存明文設定檔, 刪除加密設定檔
        with open(cfg_path, 'w', encoding='utf-8') as cfg_file:
            cfg_file.write(plain)
            os.remove(cfg_path + '.enc')

        logger.info('解密完成')
    except Exception as ex:
        logger.error(ex)

if __name__ == '__main__':
    main()
