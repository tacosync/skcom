import unittest
import subprocess
from skcom.helper import *

class TestHelper(unittest.TestCase):

    def setUp(self):
        pass

    def tearDown(self):
        pass

    def test_PsExec(self):
        # 測試一般使用者身分執行
        cmd = ['tasklist', '/fi', 'imagename eq svchost.exe', '/fo', 'csv']
        stdout = win_exec(cmd)
        self.assertIn('svchost.exe', stdout)

        # 測試系統管理員身分執行
        dll_path = os.path.expanduser(r'~\.skcom\lib\SKCOM.dll')
        cmd = ['regsvr32', '/u', dll_path]
        stdout = win_exec(cmd, admin_priv=True)
