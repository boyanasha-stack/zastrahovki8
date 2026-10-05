#!/usr/bin/env python3
"""Криптира data.json с парола (AES-256-GCM) → data.enc.json
Формат: {salt, iv, ct} в base64. Ключ = PBKDF2(password, salt, 100000, SHA256)."""
import json, os, sys, base64
from hashlib import pbkdf2_hmac
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.backends import default_backend

SRC = os.path.expanduser('~/zastrahovki-app/data.json')
OUT = os.path.expanduser('~/zastrahovki-app/data.enc.json')
PASSWORD = sys.argv[1] if len(sys.argv) > 1 else 'Money888'

def main():
    raw = open(SRC, 'rb').read()
    salt = os.urandom(16)
    iv = os.urandom(12)
    key = pbkdf2_hmac('sha256', PASSWORD.encode('utf-8'), salt, 100000, dklen=32)
    aes = AESGCM(key)
    ct = aes.encrypt(iv, raw, None)
    out = {
        'salt': base64.b64encode(salt).decode(),
        'iv': base64.b64encode(iv).decode(),
        'ct': base64.b64encode(ct).decode(),
        'v': 1
    }
    json.dump(out, open(OUT, 'w'))
    print('Криптирано:', OUT, '| размер', len(ct), 'байта')

if __name__ == '__main__':
    main()
