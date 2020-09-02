# 確認預設編碼器:
#   python encoding.py
#   > cp950
#   python -Xutf8=0 encoding.py
#   > cp950
#   python -Xutf8=1 encoding.py
#   > UTF-8
#
# See: https://docs.python.org/3.8/library/locale.html

import sys
import locale

locale.setlocale(locale.LC_ALL, 'en_US')

print('---------------')
print('   encoding ')
print('---------------')
print('Console:', sys.stdout.encoding)
print('   File:', locale.getpreferredencoding())
print('---------------')