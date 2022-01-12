"""
安裝程式範例
"""
import logging
from packaging import version

import skcom.helper
from skcom.exception import SkcomException

def main():
    """
    安裝流程
    """
    logger = logging.getLogger('helper')
    try:
        required_ver = version.parse('10.0.40219.325')
        current_ver = skcom.helper.verof_vcredist()
        if current_ver < required_ver:
            logger.info('安裝 Visual C++ 2010 可轉發套件')
            current_ver = skcom.helper.install_vcredist()
        logger.info('Visual C++ 2010 可轉發套件已安裝, 版本: %s', current_ver)

        skcom_verstr = '2.13.37'
        required_ver = version.parse(skcom_verstr)
        current_ver = skcom.helper.verof_skcom()
        if current_ver < required_ver:
            logger.info('安裝群益 API')
            current_ver = skcom.helper.install_skcom(skcom_verstr)
        logger.info('群益 API 元件已安裝, 版本: %s', current_ver)

        if not skcom.helper.has_valid_mod():
            logger.info('生成群益 API 元件的 comtypes 模組')
            skcom.helper.generate_mod()
        logger.info('群益 API 元件的 comtypes 模組已生成')
    except SkcomException as ex:
        logger.error(str(ex))

if __name__ == '__main__':
    main()
