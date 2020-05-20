"""
skcom 例外模組
"""

class SkcomException(Exception):
    """ SKCOM 通用例外 """

class ShellException(SkcomException):
    """ 在 PowerShell 環境內執行失敗 """

    def __init__(self, return_code, stderr):
        super().__init__()
        self.return_code = return_code
        self.stderr = stderr.strip()

    def get_return_code(self):
        """
        TODO
        """
        return self.return_code

    def get_stderr(self):
        """
        TODO
        """
        return self.stderr

    def __str__(self):
        return '(%d) %s' % (self.return_code, self.stderr)

class NetworkException(SkcomException):
    """ 在 PowerShell 環境內執行失敗 """

    def __init__(self, message):
        super().__init__()
        self.message = message

    def get_message(self):
        """
        TODO
        """
        return self.message

class InstallationException(SkcomException):
    """ 套件安裝失敗 """

    def __init__(self, package, required_ver):
        super().__init__()
        self.message = '%s %s 安裝失敗' % (package, required_ver)

    def __str__(self):
        return self.message

class ConfigException(SkcomException):
    """ 設定值無法使用 """

    def __init__(self, message, loaded=False):
        super().__init__()
        self.message = message
        self.loaded = loaded

    def __str__(self):
        return self.message
