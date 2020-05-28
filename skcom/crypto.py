"""
加解密模組
"""

from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives import hashes

def hash_string(srcstr, length):
    """
    TODO
    """
    salt_list = [
        b':salt-0:Plan^^^^^^^^^^^^^^^^^^^$',
        b':salt-1:Do%%%%%%%%%%%%%%%%%%%%%$',
        b':salt-2:Cancel&&&&&&&&&&&&&&&&&$',
        b':salt-3:Apologizeeeeeeeeeeeeeee$',
        b':salt-4:Salt~~~~~~~~~~~~~~~~~~~$',
        b':salt-5:Suger!!!!!!!!!!!!!!!!!!$',
        b':salt-6:Basil##################$',
        b':salt-7:Garlic@@@@@@@@@@@@@@@@@$',
        b':salt-8:Diesel>>>>>>>>>>>>>>>>>$',
        b':salt-9:Turbo==================$',
        b':salt-a:Broken?????????????????$',
        b':salt-b:Blocked................$',
        b':salt-c:Ray Tracing------------$',
        b':salt-d:CloseGL++++++++++++++++$',
        b':salt-e:Mvidia*****************$',
        b':salt-f:WinerWinerDinnerChicken$'
    ]
    digest = hashes.Hash(hashes.SHA256(), backend=default_backend())
    raw = srcstr.encode('utf-8')
    digest.update(raw)
    digest.update(salt_list[raw[-1] % 16])
    hashdata = digest.finalize()
    return hashdata[0:length]

def get_cipher(password):
    """
    TODO
    """
    key = hash_string(password, 256)
    inv = hash_string('{What The Fuck?}', 16)
    return Cipher(algorithms.AES(key), modes.CBC(inv), backend=default_backend())

def encrypt_text(plain, password):
    """
    TODO
    """
    payload = 'L%s:%s:' % (len(plain), plain)
    raw = payload.encode('utf-8')
    # 明文長度湊足 16 倍數
    pdl = (16 - len(raw)) % 16
    raw += b'-' * pdl
    # 明文長度湊足 1M
    pdl = max(1048576 - len(raw), 0)
    raw += b'+' * pdl
    enc = get_cipher(password).encryptor()
    # 密文前標記加密政策, 確保未來可以變更加密政策與向下相容
    return b'P01:' + enc.update(raw) + enc.finalize()

def decrypt_text(secret, password):
    """
    TODO
    """
    # 檢查加密政策
    if secret[0:4] != b'P01:':
        raise Exception('Unknown policy.')
    # 解密
    dec = get_cipher(password).decryptor()
    raw = dec.update(secret[4:]) + dec.finalize()
    payload = raw.decode('utf-8')
    # 取得明文實際長度
    smpos = payload.find(':')
    plen = int(payload[1:smpos])
    # 取得明文內容
    plain = payload[smpos+1:smpos+plen+1]
    return plain
