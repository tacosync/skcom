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

class NetworkException(SkcomException):
    """ 在 PowerShell 環境內執行失敗 """

    def __init__(self, message):
        self.message = message

    def getMessage(self):
        return self.message
