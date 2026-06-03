import re, base64, json, sys

try:
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    has_crypto = True
except ImportError:
    has_crypto = False
    print("cryptography not installed")

html = open(r'pl_dashboard.html', encoding='utf-8').read()

# Find ENC line (it's all on one line at line 445)
for line in html.split('\n'):
    if line.strip().startswith('const ENC'):
        enc_line = line.strip()
        break

print("ENC line found:", enc_line[:80], "...")

salt_m = re.search(r'"salt"\s*:\s*"([^"]+)"', enc_line) or re.search(r"'salt'\s*:\s*'([^']+)'", enc_line) or re.search(r'salt\s*:\s*"([^"]+)"', enc_line)
iv_m   = re.search(r'"iv"\s*:\s*"([^"]+)"', enc_line)   or re.search(r"'iv'\s*:\s*'([^']+)'", enc_line)   or re.search(r'iv\s*:\s*"([^"]+)"', enc_line)
ct_m   = re.search(r'"ct"\s*:\s*"([^"]+)"', enc_line)   or re.search(r"'ct'\s*:\s*'([^']+)'", enc_line)   or re.search(r'ct\s*:\s*"([^"]+)"', enc_line)

print("salt found:", bool(salt_m))
print("iv found:",   bool(iv_m))
print("ct found:",   bool(ct_m))

if salt_m and iv_m and ct_m and has_crypto:
    salt = base64.b64decode(salt_m.group(1))
    iv   = base64.b64decode(iv_m.group(1))
    ct   = base64.b64decode(ct_m.group(1))

    passphrase = b'karimkabir!'
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=250000)
    key = kdf.derive(passphrase)

    aesgcm = AESGCM(key)
    plaintext_b64 = aesgcm.decrypt(iv, ct, None)
    print("Raw decrypt length:", len(plaintext_b64))
    print("Raw first 200 bytes:", plaintext_b64[:200])
    # Save raw to file for inspection
    with open('raw_decrypt.bin', 'wb') as f:
        f.write(plaintext_b64)
    print("Saved raw_decrypt.bin")
    sys.exit(0)
    print("\nDecryption SUCCESS!")
    print("Top-level keys:", list(data.keys()) if isinstance(data, dict) else type(data))
    if isinstance(data, dict):
        for k, v in list(data.items())[:5]:
            print(f"  {k}: {str(v)[:100]}")
    elif isinstance(data, list):
        print("List length:", len(data))
        print("First item keys:", list(data[0].keys()) if data else "empty")
        print("First item:", str(data[0])[:300])
    # Save decrypted data for inspection
    with open('decrypted_data.json', 'w') as f:
        json.dump(data, f, indent=2)
    print("\nFull data saved to decrypted_data.json")
