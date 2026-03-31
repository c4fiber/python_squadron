# diagnose.py
import re, sys

ansi_escape = re.compile(r'\x1b\[[0-9;]*m|\[[0-9]+m')
event_pattern = re.compile(
    r"\[(?P<ts>\d{4}-\d{2}-\d{2}T[\d:.]+)\]"
    r".*\[(?P<id>[^\]]+)\]"
    r".+(?P<type>READ|WRITE):\s*(?P<size>\d+)B"
)
hex_line_pattern = re.compile(r"^\|[0-9a-f]+\|(?P<hex>[^|]+)\|")

def parse_id(raw_id):
    id_part = raw_id.split(",")[0].strip()
    m = re.match(r"^(?P<base>.+)-\d+$", id_part)
    if m:
        return m.group("base"), True, id_part
    return id_part, False, id_part

def hex_to_text(lines):
    buf = bytearray()
    for line in lines:
        m = hex_line_pattern.match(line)
        if m:
            for b in m.group("hex").strip().split():
                try: buf.append(int(b, 16))
                except: pass
    return buf.decode("utf-8", errors="ignore")

log_file = sys.argv[1]
collecting = None
hex_lines = []
printed = 0

with open(log_file, "r", encoding="utf-8", errors="ignore") as f:
    for raw_line in f:
        clean = ansi_escape.sub("", raw_line).rstrip()

        if collecting is not None:
            is_hex = (hex_line_pattern.match(clean)
                      or clean.startswith("+-")
                      or clean.startswith("|  "))
            if is_hex:
                hex_lines.append(clean)
                continue
            else:
                # suffix 없는 READ만 출력
                if not collecting["has_suffix"] and collecting["type"] == "READ":
                    text = hex_to_text(hex_lines)
                    print(f"[{collecting['full_id']}]")
                    print(f"  hex_lines: {len(hex_lines)}개")
                    print(f"  text 앞 200자: {repr(text[:200])}")
                    print(f"  '/produce' 포함: {'/produce' in text}")
                    print()
                    printed += 1
                    if printed >= 3:
                        sys.exit(0)
                collecting = None
                hex_lines = []

        m = event_pattern.search(clean)
        if not m:
            continue

        base_id, has_suffix, full_id = parse_id(m.group("id"))
        collecting = {
            "full_id": full_id,
            "base_id": base_id,
            "has_suffix": has_suffix,
            "type": m.group("type"),
        }
        hex_lines = []