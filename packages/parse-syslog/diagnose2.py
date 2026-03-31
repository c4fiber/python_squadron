# diagnose2.py - hex dump 앞뒤 원본 라인 그대로 출력
import re, sys

ansi_escape = re.compile(r'\x1b\[[0-9;]*m|\[[0-9]+m')
event_pattern = re.compile(
    r"\[(?P<ts>\d{4}-\d{2}-\d{2}T[\d:.]+)\]"
    r".*\[(?P<id>[^\]]+)\]"
    r".+(?P<type>READ|WRITE):\s*(?P<size>\d+)B"
)

log_file = sys.argv[1]
found = False

with open(log_file, "r", encoding="utf-8", errors="ignore") as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    clean = ansi_escape.sub("", line).rstrip()
    if event_pattern.search(clean):
        # 이 이벤트 라인 이후 10줄 원본 그대로 출력
        print(f"=== 이벤트 라인 (index {i}) ===")
        print(repr(clean))
        print(f"--- 이후 10줄 raw repr ---")
        for j in range(i+1, min(i+11, len(lines))):
            print(f"  [{j}] {repr(lines[j].rstrip())}")
        print()
        if found:
            sys.exit(0)
        found = True