import sys

import skcom.helper

def main():
    skcom.helper.clean_mod()
    skcom.helper.remove_skcom()
    skcom.helper.remove_vcredist()

if __name__ == '__main__':
    main()
