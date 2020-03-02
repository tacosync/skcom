class SkcomException(Exception):
    """ SKCOM 通用例外 """

class ShellException(SkcomException):
    """ 在 PowerShell 環境內執行失敗 """

    def __init__(self, return_code, stderr):
        self.return_code = return_code
        self.stderr = stderr

    def getReturnCode():
        return self.return_code

    def getStdErr():
        return self.stderr
