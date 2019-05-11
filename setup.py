import quicksk.DepHelper
from packaging import version

def main():
    not_installed = version.parse('0.0.0.0')
    required_ver = version.parse('2.13.11')
    current_ver = version.parse(quicksk.DepHelper.check_skcom())
    if current_ver >= required_ver:
        print('群益 API 元件已註冊, 版本:', current_ver)
    else:
        if current_ver == not_installed:
            print('群益 API 元件未安裝, 現在為您安裝')
        else:
            print('群益 API 元件版本太舊, 現在為您更新')
        quicksk.DepHelper.install_skcom()
        current_ver = version.parse(quicksk.DepHelper.check_skcom())
        print('安裝完成, 版本:', current_ver)

if __name__ == '__main__':
    main()
    # print(quicksk.DepHelper.check_vsredist())
    # quicksk.DepHelper.install_vsredist()
    # print(quicksk.DepHelper.check_skcom())
    # quicksk.DepHelper.install_skcom()
