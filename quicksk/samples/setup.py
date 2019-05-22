import quicksk.helper
from packaging import version

def main():
    # 檢查與更新 Visual C++ 2010 (>= 10.0.40219.325)
    required_ver = version.parse('10.0.40219.325')
    current_ver = quicksk.helper.verof_vcredist()
    if current_ver < required_ver:
        print('安裝 Visual C++ 2010 可轉發套件')
        quicksk.helper.install_vcredist()
        current_ver = quicksk.helper.verof_vcredist()
        if current_ver < required_ver:
            print('安裝失敗, 可能是您取消了安裝動作')
            return

    print('Visual C++ 2010 可轉發套件已安裝, 版本:', current_ver)

    # 檢查與更新 API 元件 (>= 2.13.11)
    required_ver = version.parse('2.13.11')
    current_ver = quicksk.helper.verof_skcom()
    if current_ver < required_ver:
        print('安裝與註冊群益 API 元件')
        quicksk.helper.install_skcom()
        current_ver = quicksk.helper.verof_skcom()

    print('群益 API 元件已安裝, 版本:', current_ver)

if __name__ == '__main__':
    main()
