import logging

import skcom.helper
from skcom.exception import SkcomException

def main():
    logger = logging.getLogger('helper')
    try:
        logger.info('移除 comtypes 套件自動生成檔案')
        skcom.helper.clean_mod()
        logger.info('移除群益 API 元件')
        skcom.helper.remove_skcom()
        logger.info('移除 Visual C++ 2010 可轉發套件')
        skcom.helper.remove_vcredist()
    except SkcomException as ex:
        logger.error(str(ex))

if __name__ == '__main__':
    main()
