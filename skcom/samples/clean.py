import sys
sys.path.append('.')

import skcom.helper
# import site

def main():
    skcom.helper.clean_mod()
    skcom.helper.remove_skcom()

if __name__ == '__main__':
    main()
