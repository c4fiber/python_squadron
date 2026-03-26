import re
import json
import csv
from datetime import datetime
from dataclasses import dataclass, field
from pathlib import Path
from collections import defaultdict

log_files = ["secret/174_sys.log.2026-02-10.txt", "secret/175_sys.log.2026-02-10.txt"]

ansi_escape = re.compile(r'\x1b\[[0-9;]*m|\[[0-9]+m')
event_pattern = re.compile(
    r"\[(?P<ts>\d{4}-\d{2}-\d{2}T[\d:.]+)\]"
    r".*\[(?P<id>[^\]]+)\]"
    r".+(?P<type>READ|WRITE):\s*(?P<size>\d+)B"
)
hex_line_pattern = re.compile(r"^\|[0-9a-f]+\|(?P<hex>[^|]+)\|")

# ── 유틸 ──────────────────────────────────────────────────────────────────────

def parse_id(raw_id: str) -> tuple[str, bool, str]:
    """(base_id, has_suffix, full_id)"""
    id_part = raw_id.split(",")[0].strip()
    m = re.match(r"^(?P<base>.+)-\d+$", id_part)
    if m:
        return m.group("base"), True, id_part
    return id_part, False, id_part

def parse_hex_to_bytes(lines: list[str]) -> bytes:
    buf = bytearray()
    for line in lines:
        m = hex_line_pattern.match(line)
        if m:
            for b in m.group("hex").strip().split():
                try:
                    buf.append(int(b, 16))
                except ValueError:
                    pass
    return bytes(buf)

def parse_packet(raw: bytes) -> dict:
    text = raw.decode("utf-8", errors="ignore")
    result = {}
    sep = "\r\n\r\n" if "\r\n\r\n" in text else "\n\n"
    if sep in text:
        header_part, body_part = text.split(sep, 1)
        result["http_header"] = header_part.strip()
    else:
        body_part = text
    j = body_part.find("{")
    if j == -1:
        j = body_part.find("[")
    if j != -1:
        try:
            result["body_json"] = json.loads(body_part[j:])
        except json.JSONDecodeError:
            result["body_raw"] = body_part[j:j+200]
    return result

def get_http_path(packet: dict) -> str:
    """'POST /produce HTTP/1.1' → '/produce'"""
    first_line = packet.get("http_header", "").split("\n")[0].strip()
    parts = first_line.split()
    return parts[1] if len(parts) >= 2 else ""

# ── 데이터 클래스 ─────────────────────────────────────────────────────────────

@dataclass
class RawEvent:
    ts: datetime
    full_id: str
    base_id: str
    has_suffix: bool
    event_type: str
    hex_lines: list[str] = field(default_factory=list)
    packet: dict = field(default_factory=dict)

@dataclass
class Result:
    stream_id: str
    request_time: datetime
    response_time: datetime
    request_packet: dict
    response_packet: dict

    def time_taken_ms(self) -> float:
        return round(
            (self.response_time - self.request_time).total_seconds() * 1000, 3
        )

# ── Step 1: 전체 파일 읽기 ────────────────────────────────────────────────────

all_events: list[RawEvent] = []
collecting: RawEvent | None = None

for log_file in log_files:
    print(f"=== 읽는 중: {log_file} ===", flush=True)
    with open(log_file, "r", encoding="utf-8", errors="ignore") as f:
        for line in f:
            clean = ansi_escape.sub("", line).rstrip()

            if collecting is not None:
                stripped = clean.strip()
                is_hex = (
                    hex_line_pattern.match(clean)
                    or stripped.startswith("+-")
                    or stripped.startswith("|  0")
                )
                if is_hex:
                    collecting.hex_lines.append(clean)
                    continue
                else:
                    all_events.append(collecting)
                    collecting = None

            m = event_pattern.search(clean)
            if not m:
                continue

            try:
                ts = datetime.fromisoformat(m.group("ts"))
            except ValueError:
                continue

            base_id, has_suffix, full_id = parse_id(m.group("id"))
            collecting = RawEvent(
                ts=ts,
                full_id=full_id,
                base_id=base_id,
                has_suffix=has_suffix,
                event_type=m.group("type"),
            )

if collecting is not None:
    all_events.append(collecting)

print(f"수집된 이벤트: {len(all_events)}개")
print(f"  suffix 없음: {sum(1 for e in all_events if not e.has_suffix)}개")
print(f"  suffix 있음: {sum(1 for e in all_events if e.has_suffix)}개")

# ── Step 2: hex dump → packet 파싱 ───────────────────────────────────────────

for ev in all_events:
    raw = parse_hex_to_bytes(ev.hex_lines)
    ev.packet = parse_packet(raw)

# ── Step 3: /produce full_id 허용 목록 구성 ──────────────────────────────────
# 헤더(suffix 없는) READ에서 경로 확인 → 직후 suffix 있는 READ의 full_id를 허용

# base_id별로 suffix 있는 READ를 타임스탬프 순으로 그룹화
suffixed_reads_by_base: defaultdict[str, list[RawEvent]] = defaultdict(list)
for ev in all_events:
    if ev.has_suffix and ev.event_type == "READ":
        suffixed_reads_by_base[ev.base_id].append(ev)
for v in suffixed_reads_by_base.values():
    v.sort(key=lambda e: e.ts)

produce_full_ids: set[str] = set()  # 수집 대상 full_id

for ev in all_events:
    if ev.has_suffix or ev.event_type != "READ":
        continue
    # suffix 없는 헤더 READ
    path = get_http_path(ev.packet)
    if "/produce" not in path:
        continue
    # 이 헤더 직후에 오는 같은 base_id의 suffix READ → 허용
    for candidate in suffixed_reads_by_base.get(ev.base_id, []):
        if candidate.ts >= ev.ts:
            produce_full_ids.add(candidate.full_id)
            break

print(f"/produce full_id 수: {len(produce_full_ids)}개  샘플: {list(produce_full_ids)[:5]}")

# ── Step 4: 타임스탬프 순 정렬 후 READ↔WRITE 페어링 ──────────────────────────

target_events = [
    ev for ev in all_events
    if ev.has_suffix and (
        (ev.event_type == "READ" and ev.full_id in produce_full_ids)
        or ev.event_type == "WRITE"
    )
]
target_events.sort(key=lambda e: e.ts)

pending: dict[str, RawEvent] = {}
results: list[Result] = []

for ev in target_events:
    if ev.event_type == "READ":
        pending[ev.full_id] = ev
    elif ev.event_type == "WRITE":
        read_ev = pending.pop(ev.full_id, None)
        if read_ev:
            results.append(Result(
                stream_id=ev.full_id,
                request_time=read_ev.ts,
                response_time=ev.ts,
                request_packet=read_ev.packet,
                response_packet=ev.packet,
            ))

print(f"완성된 페어: {len(results)}개")

# ── Step 5: CSV 저장 ──────────────────────────────────────────────────────────

csv_path = Path("produce_requests.csv")
with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
    writer = csv.writer(f)
    writer.writerow([
        "stream_id", "request_time", "response_time",
        "time_taken_ms", "request_payload", "response_payload",
    ])
    for r in results:
        req_json = json.dumps(
            r.request_packet.get("body_json"), ensure_ascii=False
        ) if r.request_packet.get("body_json") is not None else ""
        res_json = json.dumps(
            r.response_packet.get("body_json"), ensure_ascii=False
        ) if r.response_packet.get("body_json") is not None else ""

        writer.writerow([
            r.stream_id,
            r.request_time.isoformat(),
            r.response_time.isoformat(),
            r.time_taken_ms(),
            req_json,
            res_json,
        ])

print(f"✅ {len(results)}건 저장 → {csv_path.resolve()}")