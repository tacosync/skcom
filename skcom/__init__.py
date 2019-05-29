"""
skcom
"""
import os.path
import logging.config

cfg_file = os.path.dirname(__file__) + '/conf/logging.ini'
logging.config.fileConfig(cfg_file)
logger = logging.getLogger('skcom')
