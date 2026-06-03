import re, base64, json, sys
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

html = open(r'pl_dashboard.html', encoding='utf-8').read()

for line in html.split('\n'):
    if line.strip().startswith('const ENC'):
        enc_line = line.strip()
        break

salt = base64.b64decode(re.search(r'"salt"\s*:\s*"([^"]+)"', enc_line).group(1))
iv   = base64.b64decode(re.search(r'"iv"\s*:\s*"([^"]+)"', enc_line).group(1))
ct   = base64.b64decode(re.search(r'"ct"\s*:\s*"([^"]+)"', enc_line).group(1))

kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=250000)
key = kdf.derive(b'karimkabir!')

aesgcm = AESGCM(key)
raw_b64 = aesgcm.decrypt(iv, ct, None)

# Decode base64 to get the JS bundle string
js_bundle = base64.b64decode(raw_b64).decode('utf-8')

# Extract the RAW array JSON from `const RAW=[...]`
raw_match = re.search(r'const RAW=(\[[\s\S]*?\]);?\s*(?:const |var |let |//|$)', js_bundle)
if not raw_match:
    # Try just finding the JSON array
    raw_match = re.search(r'const RAW=(\[[\s\S]+)', js_bundle)
    if raw_match:
        arr_str = raw_match.group(1)
        # Find the closing bracket
        depth = 0
        end = 0
        for i, c in enumerate(arr_str):
            if c == '[': depth += 1
            elif c == ']':
                depth -= 1
                if depth == 0:
                    end = i + 1
                    break
        arr_str = arr_str[:end]
        print("Extracted array length:", len(arr_str))
    else:
        print("Could not find RAW array")
        print("Bundle start:", js_bundle[:500])
        sys.exit(1)
else:
    arr_str = raw_match.group(1)

try:
    data = json.loads(arr_str)
except json.JSONDecodeError as e:
    # The array ends with }]; and then more JS follows — find last ] and truncate
    last_bracket = arr_str.rfind(']', 0, e.pos + 10)
    arr_str = arr_str[:last_bracket + 1]
    data = json.loads(arr_str)
print(f"Total records: {len(data)}")

# Show unique periods
periods = sorted(set((r.get('year'), r.get('pnum')) for r in data))
print("Periods present:", periods)

# Show unique stores
stores = sorted(set(r.get('store') for r in data))
print(f"\nStores ({len(stores)}):", stores)

# Show all keys for one record
print("\nRecord keys:", list(data[0].keys()))
print("\nSample record (P4 2026 or latest):")
latest = sorted(data, key=lambda r: (r.get('year',0), r.get('pnum',0)))[-1]
print(json.dumps(latest, indent=2))

# Show all records from the latest period
latest_year = latest['year']
latest_pnum = latest['pnum']
print(f"\nAll records from P{latest_pnum} {latest_year}:")
for r in data:
    if r.get('year') == latest_year and r.get('pnum') == latest_pnum:
        print(f"  {r['store']}: sales={r.get('sales')}, cogs={r.get('cogs')}, labor={r.get('labor')}, rent={r.get('rent')}, other={r.get('other')}")

# Save the full decoded bundle for re-encryption later
with open('bundle.txt', 'w', encoding='utf-8') as f:
    f.write(js_bundle)
with open('raw_data.json', 'w') as f:
    json.dump(data, f, indent=2)
print("\nSaved bundle.txt and raw_data.json")
