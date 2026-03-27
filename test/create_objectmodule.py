import base64, os, sys
# Read base64 data from file and decode
b64_path = r"C:\Configuration_downloads\BASERP25	est\_content.b64"
with open(b64_path, "r") as bf:
    data = bf.read()
content = base64.b64decode(data).decode("utf-8")
filepath = r"C:\Configuration_downloads\BASERP25\_RarzrabotkiОтчетыА_СравнитьОстаткиПодОтчетныхЛицЕРПсBASБухгалтерия\Ext\ObjectModule.bsl"
os.makedirs(os.path.dirname(filepath), exist_ok=True)
with open(filepath, "w", encoding="utf-8-sig", newline="") as f:
    f.write(content)
print(f"Created: {len(content)} chars, {os.path.getsize(filepath)} bytes")
# Verify
with open(filepath, "rb") as f:
    raw = f.read()
print(f"BOM: {raw[:3] == bytes([0xEF, 0xBB, 0xBF])}")
print(f"CRLF: {chr(13).encode() + chr(10).encode() in raw}")
print(f"TAB: {chr(9).encode() in raw}")
