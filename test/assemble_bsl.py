import os, base64

# Read all parts
parts = []
for i in [1, 2, 3, 4, 5]:
    p = os.path.join(r"C:\Configuration_downloads\BASERP25\test", f"_p{i}.b64")
    if os.path.exists(p):
        with open(p, "r") as f:
            encoded = f.read()
        parts.append(base64.b64decode(encoded).decode("utf-8"))

content = "".join(parts)

filepath = os.path.join(
    r"C:\Configuration_downloads\BASERP25\_Rarzrabotki",
    "\u041e\u0442\u0447\u0435\u0442\u044b",
    "\u0410_\u0421\u0440\u0430\u0432\u043d\u0438\u0442\u044c\u041e\u0441\u0442\u0430\u0442\u043a\u0438\u041f\u043e\u0434\u041e\u0442\u0447\u0435\u0442\u043d\u044b\u0445\u041b\u0438\u0446\u0415\u0420\u041f\u0441BAS\u0411\u0443\u0445\u0433\u0430\u043b\u0442\u0435\u0440\u0438\u044f",
    "Ext",
    "ObjectModule.bsl"
)
os.makedirs(os.path.dirname(filepath), exist_ok=True)
with open(filepath, "w", encoding="utf-8-sig", newline="") as f:
    f.write(content)
print(f"Created: {len(content)} chars, {os.path.getsize(filepath)} bytes")

# Verify
with open(filepath, "rb") as f:
    raw = f.read()
print(f"BOM: {raw[:3] == bytes([0xEF, 0xBB, 0xBF])}")
print(f"CRLF: {b'\r\n' in raw}")
print(f"TAB: {b'\t' in raw}")
print(f"Total lines: {raw.count(b'\r\n')}")
