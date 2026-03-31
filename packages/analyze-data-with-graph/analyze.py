import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import numpy as np
from pathlib import Path
import sys
import warnings

# ─────────────────────────────────────────────
# [추가] 한글 폰트 설정 (Windows: 맑은 고딕)
# ─────────────────────────────────────────────
import matplotlib as mpl

# 1. 시스템 폰트 설정
plt.rcParams['font.family'] = 'Malgun Gothic'

# 2. 마이너스 기호(-) 깨짐 방지
plt.rcParams['axes.unicode_minus'] = False

# (선택 사항) Retina 디스플레이 지원 (그래프 선명도 향상)
# %matplotlib inline # 주피터 노트북 사용 시
# ─────────────────────────────────────────────

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────
# 1. 파일 로드 (csv / xlsx 자동 감지)
# ─────────────────────────────────────────────
FILE_PATH = "K:\\GoogleDrive\\00. Quick Share\\AIA생명\\Application_Insight_021013_021016_original.xlsx"  # 파일 경로를 여기에 입력하거나 인자로 전달

# 인자로 파일 경로 받기
if len(sys.argv) > 1:
    FILE_PATH = sys.argv[1]

# 파일이 없으면 샘플 데이터 생성
if FILE_PATH is None or not Path(FILE_PATH).exists():
    print("📋 파일 없음 → 샘플 데이터로 실행합니다")
    import random
    random.seed(42)
    
    # 샘플: 7일 * 24시간, FALSE가 특정 시간대에 몰리도록 설계
    timestamps = pd.date_range("2024/01/01", periods=2000, freq="10min")
    values = []
    for ts in timestamps:
        hour = ts.hour
        # 새벽(0~6시)과 점심(11~13시)에 FALSE 확률 높임
        if 0 <= hour < 6:
            values.append("FALSE" if random.random() < 0.6 else "TRUE")
        elif 11 <= hour < 14:
            values.append("FALSE" if random.random() < 0.4 else "TRUE")
        else:
            values.append("FALSE" if random.random() < 0.1 else "TRUE")
    
    df = pd.DataFrame({
        "A": range(len(timestamps)),
        "B": [ts.strftime("%-m/%d/%Y, %-I:%M:%S.%f")[:-3] + " " + ts.strftime("%p") for ts in timestamps],
        "C": 0,
        "D": 0,
        "E": values
    })
else:
    ext = Path(FILE_PATH).suffix.lower()
    if ext == ".csv":
        df = pd.read_csv(FILE_PATH, header=0)
    elif ext in [".xlsx", ".xls"]:
        df = pd.read_excel(FILE_PATH, header=0)
    else:
        raise ValueError(f"지원하지 않는 파일 형식: {ext}")
    print(f"✅ 파일 로드 완료: {FILE_PATH} ({len(df):,}행)")

# ─────────────────────────────────────────────
# 2. 컬럼 이름 통일 (실제 파일은 컬럼명이 다를 수 있음)
# ─────────────────────────────────────────────
# B컬럼(인덱스 1), E컬럼(인덱스 4)로 접근
timestamp_col = df.columns[0]   # B컬럼
status_col    = df.columns[4]   # E컬럼

df = df.rename(columns={timestamp_col: "timestamp", status_col: "status"})

# ─────────────────────────────────────────────
# 3. 타입 변환
# ─────────────────────────────────────────────
# timestamp: "d/mm/yyyy h:mm:ss.000 AM" 형식 파싱
df["timestamp"] = pd.to_datetime(
    df["timestamp"],
    format="%m/%d/%Y, %I:%M:%S.%f %p",
    errors="coerce"
)

# 파싱 실패 행 확인
failed = df["timestamp"].isna().sum()
if failed > 0:
    print(f"⚠️  timestamp 파싱 실패: {failed}행 (제외됩니다)")
    df = df.dropna(subset=["timestamp"])

# status: 대소문자 통일
df["status"] = df["status"].astype(str).str.strip().str.upper()
df_false = df[df["status"] == "FALSE"].copy()

total   = len(df)
n_false = len(df_false)
print(f"\n📊 전체: {total:,}건 | FALSE: {n_false:,}건 ({n_false/total*100:.1f}%)")

# ─────────────────────────────────────────────
# 4. 13~16시 필터링 + 파생 컬럼
# ─────────────────────────────────────────────
TARGET_START = 13
TARGET_END   = 16   # 16:59까지 포함

df_range       = df[(df["timestamp"].dt.hour >= TARGET_START) &
                    (df["timestamp"].dt.hour <= TARGET_END)].copy()
df_false_range = df_range[df_range["status"] == "FALSE"].copy()

total_range   = len(df_range)
n_false_range = len(df_false_range)
print(f"🕐 {TARGET_START}~{TARGET_END}시 범위 — 전체: {total_range:,}건 | FALSE: {n_false_range:,}건 ({n_false_range/max(total_range,1)*100:.1f}%)")

# 파생 컬럼
# hhmm_abs: 13:05 → 785 (분 절대값) — 타임라인 X축용
df_false_range["hour"]     = df_false_range["timestamp"].dt.hour
df_false_range["minute"]   = df_false_range["timestamp"].dt.minute
df_false_range["hhmm_abs"] = df_false_range["hour"] * 60 + df_false_range["minute"]
df_false_range["date"]     = df_false_range["timestamp"].dt.date

df_range["hour"]     = df_range["timestamp"].dt.hour
df_range["minute"]   = df_range["timestamp"].dt.minute
df_range["hhmm_abs"] = df_range["hour"] * 60 + df_range["minute"]

# 분 단위 집계 인덱스 (13:00 ~ 16:59 → 780~1019)
all_minutes = range(TARGET_START * 60, (TARGET_END + 1) * 60)  # 240분

# FALSE 건수 (분별)
false_by_min = df_false_range.groupby("hhmm_abs").size().reindex(all_minutes, fill_value=0)

# 전체 건수 (분별) — FALSE 비율 계산용
total_by_min = df_range.groupby("hhmm_abs").size().reindex(all_minutes, fill_value=0)

# FALSE 비율 (분별)
ratio_by_min = (false_by_min / total_by_min.replace(0, np.nan) * 100).fillna(0)

# X축 레이블: 780→"13:00", 800→"13:20" ...
def min_to_hhmm(m):
    return f"{m // 60:02d}:{m % 60:02d}"

x_vals   = list(all_minutes)
x_labels = [min_to_hhmm(m) for m in x_vals]

# 10분 간격 tick 위치
tick_positions = [m for m in x_vals if m % 10 == 0]
tick_labels    = [min_to_hhmm(m) for m in tick_positions]

# ─────────────────────────────────────────────
# 5. 그래프 4종
# ─────────────────────────────────────────────
fig, axes = plt.subplots(4, 1, figsize=(18, 25),)
fig.suptitle(f"FALSE 분포 분석 — {TARGET_START}:00 ~ {TARGET_END}:59 (분 단위)",
             fontsize=15, fontweight="bold", y=1.02)

COLORS = {"bar": "#E74C3C", "line": "#C0392B", "ratio": "#8E44AD",
          "bg": "#F8F9FA", "grid": "#BDC3C7"}

# ── 그래프 1: 분별 FALSE 건수 (막대) ──────────
ax1 = axes[0]
ax1.bar(x_vals, false_by_min.values, color=COLORS["bar"], alpha=0.75, width=0.8)
ax1.set_title("① 분별 FALSE 발생 건수", fontweight="bold", fontsize=12)
ax1.set_ylabel("건수")
ax1.set_facecolor(COLORS["bg"])
ax1.grid(axis="y", linestyle="--", alpha=0.4, color=COLORS["grid"])
ax1.set_xticks(tick_positions)
ax1.set_xticklabels(tick_labels, rotation=45, ha="right", fontsize=7)
ax1.set_xlim(min(x_vals) - 1, max(x_vals) + 1)

# 시간 경계선 (14:00, 15:00, 16:00)
for h in range(TARGET_START + 1, TARGET_END + 1):
    ax1.axvline(h * 60, color="#2C3E50", linewidth=1.2, linestyle=":", alpha=0.6)
    ax1.text(h * 60 + 0.5, ax1.get_ylim()[1] * 0.9, f"{h}:00",
             fontsize=8, color="#2C3E50", alpha=0.8)

# ── 그래프 3: 분별 FALSE 비율(%) ──────────────
ax3 = axes[1]
ax3.fill_between(x_vals, ratio_by_min.values, alpha=0.3, color=COLORS["ratio"])
ax3.plot(x_vals, ratio_by_min.values, color=COLORS["ratio"], linewidth=1.2)
ax3.axhline(ratio_by_min.mean(), color="#E67E22", linewidth=1.5,
            linestyle="--", label=f"평균 {ratio_by_min.mean():.1f}%")
ax3.set_title("③ 분별 FALSE 비율 (%)", fontweight="bold", fontsize=12)
ax3.set_ylabel("FALSE 비율 (%)")
ax3.set_ylim(0, 105)
ax3.set_facecolor(COLORS["bg"])
ax3.grid(linestyle="--", alpha=0.4, color=COLORS["grid"])
ax3.set_xticks(tick_positions)
ax3.set_xticklabels(tick_labels, rotation=45, ha="right", fontsize=7)
ax3.set_xlim(min(x_vals) - 1, max(x_vals) + 1)
ax3.legend(fontsize=9)
for h in range(TARGET_START + 1, TARGET_END + 1):
    ax3.axvline(h * 60, color="#2C3E50", linewidth=1.2, linestyle=":", alpha=0.6)

# ── 그래프 4: 분별 FALSE 건수 + FALSE 비율 (이중 Y축) ──
ax4 = axes[2]

# 막대: 분별 FALSE 건수
ax4.bar(x_vals, false_by_min.values, color=COLORS["bar"], alpha=0.75, width=0.8, label="FALSE 건수")
ax4.set_title("④ 분별 FALSE 건수 및 비율", fontweight="bold", fontsize=12)
ax4.set_ylabel("건수", color=COLORS["bar"])
ax4.tick_params(axis="y", labelcolor=COLORS["bar"])
ax4.set_facecolor(COLORS["bg"])
ax4.grid(axis="y", linestyle="--", alpha=0.4, color=COLORS["grid"])
ax4.set_xticks(tick_positions)
ax4.set_xticklabels(tick_labels, rotation=45, ha="right", fontsize=7)
ax4.set_xlim(min(x_vals) - 1, max(x_vals) + 1)

# 오른쪽 Y축: 분별 FALSE 비율 꺾은선
ax4r = ax4.twinx()
ax4r.plot(x_vals, ratio_by_min.values, color=COLORS["ratio"],
          linewidth=1.0, alpha=0.8, label="FALSE 비율")
ax4r.axhline(ratio_by_min.mean(), color="#E67E22", linewidth=1.2,
             linestyle="--", label=f"평균 {ratio_by_min.mean():.1f}%")
ax4r.set_ylabel("FALSE 비율 (%)", color=COLORS["ratio"])
ax4r.tick_params(axis="y", labelcolor=COLORS["ratio"])
ax4r.set_ylim(0, 110)

# 시간 경계선
for h in range(TARGET_START + 1, TARGET_END + 1):
    ax4.axvline(h * 60, color="#2C3E50", linewidth=1.2, linestyle=":", alpha=0.6)
    ax4.text(h * 60 + 0.5, ax4.get_ylim()[1] * 0.9, f"{h}:00",
             fontsize=8, color="#2C3E50", alpha=0.8)

# 범례 합치기
lines1, labels1 = ax4.get_legend_handles_labels()
lines2, labels2 = ax4r.get_legend_handles_labels()
ax4r.legend(lines1 + lines2, labels1 + labels2, loc="upper right", fontsize=9)

# ── 그래프 5: AP열 Result Code 분별 선그래프 ──
ax5 = axes[3]

AP_COL_INDEX = 41
CODE_ORDER  = ["0 [not sent in full (see exception telemetries)]", "200", "500", "값없음"]
CODE_COLORS = {"0 [not sent in full (see exception telemetries)]": "#3498DB", "200": "#2ECC71", "500": "#E74C3C", "값없음": "#95A5A6"}
LINE_STYLES = {"0 [not sent in full (see exception telemetries)]": "-", "200": "-", "500": "-", "값없음": "--"}

if AP_COL_INDEX < len(df.columns):
    ap_col_name = df.columns[AP_COL_INDEX]

    # AP열 정규화
    df_range["result_code"] = (
        df_range[ap_col_name]
        .astype(str).str.strip()
        .apply(lambda x: x if x in ["0 [not sent in full (see exception telemetries)]", "200", "500"] else "값없음")
    )

    # 분별 × 코드별 건수 집계
    min_code_counts = (
        df_range.groupby(["hhmm_abs", "result_code"])
        .size()
        .unstack(fill_value=0)
        .reindex(columns=CODE_ORDER, fill_value=0)
        .reindex(all_minutes, fill_value=0)
    )

    # 코드별 선 그리기
    for code in CODE_ORDER:
        if code in min_code_counts.columns:
            ax5.plot(
                x_vals,
                min_code_counts[code].values,
                color=CODE_COLORS[code],
                linestyle=LINE_STYLES[code],
                linewidth=1.5,
                alpha=0.85,
                label=f"code {code}"
            )

    ax5.set_title("⑤ (APIM)분별 Result Code 분포 (AP열, 선그래프)", fontweight="bold", fontsize=12)
    ax5.set_ylabel("건수")
    ax5.set_facecolor(COLORS["bg"])
    ax5.grid(linestyle="--", alpha=0.4, color=COLORS["grid"])
    ax5.set_xticks(tick_positions)
    ax5.set_xticklabels(tick_labels, rotation=45, ha="right", fontsize=7)
    ax5.set_xlim(min(x_vals) - 1, max(x_vals) + 1)
    ax5.legend(loc="upper right", fontsize=9, title="Result Code")

    for h in range(TARGET_START + 1, TARGET_END + 1):
        ax5.axvline(h * 60, color="#2C3E50", linewidth=1.2, linestyle=":", alpha=0.6)
        ax5.text(h * 60 + 0.5, ax5.get_ylim()[1] * 0.9,
                 f"{h}:00", fontsize=8, color="#2C3E50", alpha=0.8)
else:
    ax5.text(0.5, 0.5, f"AP열(인덱스 {AP_COL_INDEX}) 없음\n전체 컬럼 수: {len(df.columns)}",
             ha="center", va="center", transform=ax5.transAxes, fontsize=12, color="red")
    ax5.set_title("⑤ Result Code 분포 — 데이터 없음", fontweight="bold", fontsize=12)

plt.tight_layout()
plt.savefig("K:\\GoogleDrive\\00. Quick Share\\AIA생명\\outputs\\false_distribution.png", dpi=150, bbox_inches="tight")
print("\n✅ 저장 완료: false_distribution.png")

# ─────────────────────────────────────────────
# 6. 요약 통계 출력
# ─────────────────────────────────────────────
top5 = false_by_min.nlargest(5)
print("\n─── FALSE 최다 발생 TOP 5 (분 단위) ───")
for m, cnt in top5.items():
    print(f"  {min_to_hhmm(m)}  →  {cnt}건")

print("\n─── 시간별 소계 ───")
for h in range(TARGET_START, TARGET_END + 1):
    h_false = df_false_range[df_false_range["hour"] == h].shape[0]
    h_total = df_range[df_range["hour"] == h].shape[0]
    h_ratio = h_false / h_total * 100 if h_total > 0 else 0
    print(f"  {h}시: FALSE {h_false}건 / 전체 {h_total}건 ({h_ratio:.1f}%)")
