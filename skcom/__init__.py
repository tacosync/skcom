"""
skcom
"""
# pylint: disable=invalid-name

import os.path
import logging.config

import yaml

cfg_yaml = '{}/conf/logging.yaml'.format(os.path.dirname(__file__))
with open(cfg_yaml, 'r') as cfg_file:
    cfg_dict = yaml.load(open(cfg_yaml, 'r'), Loader=yaml.SafeLoader)

    for name in cfg_dict['handlers']:
        handler = cfg_dict['handlers'][name]
        if 'filename' in handler:
            # Replace ~ (home) as absolute path.
            if handler['filename'].startswith('~'):
                handler['filename'] = os.path.expanduser(handler['filename'])
            # Create dirs for log file
            dirname = os.path.dirname(handler['filename'])
            if not os.path.isdir(dirname):
                os.makedirs(dirname)

    logging.config.dictConfig(cfg_dict)
