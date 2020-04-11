"""
安裝程式範例
"""
import logging
from packaging import version

import skcom.helper

def main():
    """
    安裝流程
    """
    logger = logging.getLogger('helper')

    # 檢查與更新 Visual C++ 2010 (>= 10.0.40219.325)
    required_ver = version.parse('10.0.40219.325')
    current_ver = skcom.helper.verof_vcredist()
    if current_ver < required_ver:
        logger.info('安裝 Visual C++ 2010 可轉發套件')
        skcom.helper.install_vcredist()
        current_ver = skcom.helper.verof_vcredist()
    if current_ver < required_ver:
        logger.error('安裝失敗, 可能是您取消了安裝動作')
        return
    logger.info('Visual C++ 2010 可轉發套件已安裝, 版本: %s', current_ver)

    # 檢查與更新 API 元件 (>= 2.13.21)
    skcom_verstr = '2.13.21'
    required_ver = version.parse(skcom_verstr)
    current_ver = skcom.helper.verof_skcom()
    if current_ver < required_ver:
        logger.info('安裝與註冊群益 API 元件')
        skcom.helper.install_skcom(skcom_verstr)
        current_ver = skcom.helper.verof_skcom()
    if current_ver < required_ver:
        logger.error('安裝失敗, 可能是您取消了安裝動作')
        return
    logger.info('群益 API 元件已安裝, 版本: %s', current_ver)

    if not skcom.helper.has_valid_mod():
        skcom.helper.generate_mod()

    logger.info('群益 API 元件模組已生成')

if __name__ == '__main__':
    main()
