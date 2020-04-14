"""
skcom
"""
# pylint: disable=invalid-name

import os.path
import logging.config

import yaml

cfg_yaml = '{}/conf/logging.yaml'.format(os.path.dirname(__file__))
with open(cfg_yaml, 'r', encoding='utf-8') as cfg_file:
    cfg_dict = yaml.load(cfg_file, Loader=yaml.SafeLoader)

    # 自動替換 ~ 為家目錄, 以及自動產生 log 目錄
    for name in cfg_dict['handlers']:
        handler = cfg_dict['handlers'][name]
        if 'filename' in handler:
            if handler['filename'].startswith('~'):
                handler['filename'] = os.path.expanduser(handler['filename'])
            dirname = os.path.dirname(handler['filename'])
            if not os.path.isdir(dirname):
                os.makedirs(dirname)

    # 檢查 Telegram 機器人設定是否有效
    enable_tgbot = False
    blank_token = '1234567890:-----------------------------------'
    cfg_skcom_path = os.path.expanduser(cfg_dict['handlers']['telegram']['config'])
    if os.path.isfile(cfg_skcom_path):
        with open(cfg_skcom_path, 'r', encoding='utf-8') as cfg_skcom_stream:
            cfg_skcom_dict = yaml.load(cfg_skcom_stream, Loader=yaml.SafeLoader)
            enable_tgbot = cfg_skcom_dict['telegram']['token'] != blank_token

    # 如果 Telegram 機器人的設定無效, 就抽掉 logging handler
    if not enable_tgbot:
        del cfg_dict['handlers']['telegram']
        cfg_dict['loggers']['busm']['handlers'].remove('telegram')

    logging.config.dictConfig(cfg_dict)
