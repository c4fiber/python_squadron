"""
app.py — UI + 설정 + 템플릿 쓰기
GUI, settings.ini 관리, 판매가 계산, onecell 템플릿 출력을 담당한다.
"""
from __future__ import annotations

import configparser
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

from onecell.parser import ProductRecord
from onecell.preset import REGISTRY, PresetConfig


# ─────────────────────────────────────────────
# 경로 헬퍼
# ─────────────────────────────────────────────
def _resource(relative: str) -> str:
    """PyInstaller 패키징 대응 리소스 경로."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative)


TEMPLATE_PATH = _resource("onecell_template.xlsx")
SETTINGS_PATH = os.path.join(
    os.path.dirname(os.path.abspath(sys.argv[0])), "settings.ini"
)


# ─────────────────────────────────────────────
# Settings
# ─────────────────────────────────────────────
def _default_cfg() -> dict[str, dict[str, str]]:
    """설정 파일 기본값. REGISTRY 기반으로 자동 생성."""
    defaults: dict[str, dict[str, str]] = {
        "general": {"tag": ""},
        "pricing": {"margin_rate": "15"},
    }
    for preset in REGISTRY.values():
        defaults[preset.settings_key] = {"shipping_fee": str(preset.default_fee)}
    return defaults


def load_settings() -> configparser.ConfigParser:
    cfg = configparser.ConfigParser()
    if os.path.exists(SETTINGS_PATH):
        cfg.read(SETTINGS_PATH, encoding="utf-8")
    for section, kvs in _default_cfg().items():
        if not cfg.has_section(section):
            cfg[section] = kvs
        else:
            for k, v in kvs.items():
                if not cfg.has_option(section, k):
                    cfg.set(section, k, v)
    return cfg


def save_settings(cfg: configparser.ConfigParser) -> None:
    with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
        cfg.write(f)


# ─────────────────────────────────────────────
# 판매가 계산
# ─────────────────────────────────────────────
def calc_sell_price(buy_price: float, shipping_fee: float, margin_rate: float) -> int:
    """판매가 = round((매입가 + 배송비) × 1.1 × (1 + 마진율/100) / 10) × 10"""
    raw = (buy_price + shipping_fee) * 1.1 * (1 + margin_rate / 100)
    return int(round(raw / 10) * 10)


# ─────────────────────────────────────────────
# 템플릿 쓰기
# ─────────────────────────────────────────────
def write_to_template(
    records: list[ProductRecord],
    tag: str,
    margin_rate: float,
    shipping_fee: float,
    save_path: str,
) -> None:
    """ProductRecord 리스트를 onecell_template.xlsx에 써서 저장."""
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb["기초상품정보"]

    for i, rec in enumerate(records):
        row = 8 + i  # 데이터 시작 행
        sell_price = calc_sell_price(rec.buy_price, shipping_fee, margin_rate) if rec.buy_price > 0 else 0
        base_name  = f"[{tag}] {rec.product_name}" if tag else rec.product_name

        row_data: dict[str, object] = {
            "A":  base_name,
            "B":  rec.attr_name1,
            "C":  rec.attr_val1,
            "D":  rec.attr_name2,
            "E":  rec.attr_val2,
            "H":  sell_price,
            "J":  rec.stock,
            "K":  rec.seller_code,
            "M":  rec.brand,
            "N":  rec.manufacturer,
            "O":  "상세설명참조",
            "AC": rec.detail_html,
            "AG": rec.notice_category,
            "BC": rec.image_url,
        }
        for col_letter, value in row_data.items():
            ws.cell(row=row, column=column_index_from_string(col_letter), value=value)

    wb.save(save_path)


# ─────────────────────────────────────────────
# Mock: 업로드 클래스 (추후 자동화용)
# ─────────────────────────────────────────────
class OnecellUploader:
    """
    onecell 파일 업로드 Mock. 추후 실제 로직으로 교체.
    Usage:
        uploader = OnecellUploader(credentials={"id": "...", "pw": "..."})
        result = uploader.upload("output.xlsx")
    """
    def __init__(self, credentials: dict | None = None) -> None:
        self.credentials = credentials or {}
        self._logged_in  = False

    def login(self) -> bool:
        print(f"[Mock] login: {self.credentials.get('id', 'unknown')}")
        self._logged_in = True
        return True

    def upload(self, filepath: str) -> dict:
        if not self._logged_in:
            self.login()
        print(f"[Mock] upload: {filepath}")
        return {
            "success": True,
            "message": f"[Mock] {os.path.basename(filepath)} uploaded",
            "uploaded_count": 0,
        }

    def logout(self) -> None:
        self._logged_in = False


# ─────────────────────────────────────────────
# GUI
# ─────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("onecell 자동 입력 도구")
        self.resizable(False, False)
        self.cfg = load_settings()
        self._fee_vars: dict[str, tk.StringVar] = {}
        self._build_ui()

    # ── UI 구성 ─────────────────────────────────
    def _build_ui(self) -> None:
        pad = {"padx": 12, "pady": 6}

        # 설정 프레임
        frm_cfg = ttk.LabelFrame(self, text="설정 (settings.ini)")
        frm_cfg.pack(fill="x", **pad)

        ttk.Label(frm_cfg, text="태그:").grid(row=0, column=0, sticky="w", padx=8, pady=4)
        self.var_tag = tk.StringVar(value=self.cfg.get("general", "tag", fallback=""))
        ttk.Entry(frm_cfg, textvariable=self.var_tag, width=16).grid(row=0, column=1, sticky="w", padx=4)
        ttk.Label(frm_cfg, text="예) 26신상  →  [26신상] 상품명").grid(
            row=0, column=2, columnspan=4, sticky="w", padx=4)

        ttk.Label(frm_cfg, text="마진율 (%):").grid(row=1, column=0, sticky="w", padx=8, pady=4)
        self.var_margin = tk.StringVar(value=self.cfg.get("pricing", "margin_rate", fallback="15"))
        ttk.Entry(frm_cfg, textvariable=self.var_margin, width=8).grid(row=1, column=1, sticky="w", padx=4)
        ttk.Label(frm_cfg, text="판매가 = (매입가+배송비)×1.1×(1+마진율%)").grid(
            row=1, column=2, columnspan=4, sticky="w", padx=4)

        # 배송비: REGISTRY 기반으로 동적 생성
        col_offset = 0
        for preset in REGISTRY.values():
            ttk.Label(frm_cfg, text=f"{preset.label} 배송비:").grid(
                row=2, column=col_offset, sticky="w", padx=8, pady=4)
            var = tk.StringVar(value=self.cfg.get(
                preset.settings_key, "shipping_fee", fallback=str(preset.default_fee)
            ))
            self._fee_vars[preset.label] = var
            ttk.Entry(frm_cfg, textvariable=var, width=8).grid(
                row=2, column=col_offset + 1, sticky="w", padx=4)
            col_offset += 2

        ttk.Button(frm_cfg, text="설정 저장", command=self._save_settings).grid(
            row=3, column=1, sticky="w", padx=4, pady=6)

        # 프리셋 선택 (REGISTRY 기반 동적 생성)
        frm_preset = ttk.LabelFrame(self, text="데이터 종류 선택")
        frm_preset.pack(fill="x", **pad)
        self.var_preset = tk.StringVar(value=next(iter(REGISTRY)))
        for i, label in enumerate(REGISTRY):
            ttk.Radiobutton(
                frm_preset, text=label, variable=self.var_preset,
                value=label, command=self._on_preset_change,
            ).grid(row=0, column=i, padx=16, pady=6, sticky="w")
        self.lbl_preset_info = ttk.Label(frm_preset, text="", foreground="gray")
        self.lbl_preset_info.grid(row=0, column=len(REGISTRY), padx=8)
        self._on_preset_change()

        # 파일 선택
        frm_file = ttk.LabelFrame(self, text="상품 파일 선택")
        frm_file.pack(fill="x", **pad)
        self.var_filepath = tk.StringVar()
        ttk.Entry(frm_file, textvariable=self.var_filepath, width=48, state="readonly").grid(
            row=0, column=0, padx=8, pady=8)
        ttk.Button(frm_file, text="파일 선택...", command=self._browse_file).grid(
            row=0, column=1, padx=4)

        # 실행
        frm_btn = ttk.Frame(self)
        frm_btn.pack(fill="x", padx=12, pady=8)
        self.btn_run = ttk.Button(frm_btn, text="▶  자동 입력", command=self._run)
        self.btn_run.pack(side="left", ipadx=16, ipady=4)

        # 상태 / 진행 바
        self.var_status = tk.StringVar(value="파일을 선택하고 자동 입력을 누르세요.")
        self.lbl_status = ttk.Label(self, textvariable=self.var_status, foreground="gray")
        self.lbl_status.pack(anchor="w", padx=14)
        self.progress = ttk.Progressbar(self, mode="indeterminate", length=530)
        self.progress.pack(padx=12, pady=(4, 10))

    # ── 콜백 ────────────────────────────────────
    def _on_preset_change(self) -> None:
        label = self.var_preset.get()
        fee_str = self._fee_vars.get(label, tk.StringVar()).get()
        try:
            self.lbl_preset_info.config(text=f"배송비: {int(float(fee_str)):,}원")
        except ValueError:
            self.lbl_preset_info.config(text="")

    def _save_settings(self) -> None:
        self.cfg["general"]["tag"]         = self.var_tag.get().strip()
        self.cfg["pricing"]["margin_rate"] = self.var_margin.get().strip()
        for preset in REGISTRY.values():
            var = self._fee_vars.get(preset.label)
            if var:
                self.cfg[preset.settings_key]["shipping_fee"] = var.get().strip()
        save_settings(self.cfg)
        self._set_status("설정이 저장되었습니다.", "green")

    def _browse_file(self) -> None:
        preset = REGISTRY[self.var_preset.get()]
        path = filedialog.askopenfilename(
            title="상품 파일 선택", filetypes=preset.file_types)
        if path:
            self.var_filepath.set(path)
            self._set_status(f"선택됨: {os.path.basename(path)}", "gray")

    def _run(self) -> None:
        filepath = self.var_filepath.get()
        if not filepath:
            messagebox.showwarning("파일 없음", "상품 파일을 먼저 선택해 주세요.")
            return
        try:
            margin = float(self.var_margin.get())
        except ValueError:
            messagebox.showerror("입력 오류", "마진율은 숫자로 입력해 주세요.")
            return

        label  = self.var_preset.get()
        preset = REGISTRY[label]
        tag    = self.var_tag.get().strip()
        try:
            shipping_fee = float(self._fee_vars[label].get())
        except ValueError:
            messagebox.showerror("입력 오류", "배송비는 숫자로 입력해 주세요.")
            return

        save_path = filedialog.asksaveasfilename(
            title="결과 파일 저장", defaultextension=".xlsx",
            filetypes=[("Excel 파일", "*.xlsx")],
            initialfile="onecell_output.xlsx",
        )
        if not save_path:
            return

        self.btn_run.config(state="disabled")
        self.progress.start(10)
        self._set_status("처리 중...", "blue")
        self.update()

        try:
            records = preset.make_parser().parse(filepath)
            write_to_template(records, tag, margin, shipping_fee, save_path)

            self.progress.stop()
            self.btn_run.config(state="normal")
            self._set_status(f"완료: {len(records)}개 상품 → {os.path.basename(save_path)}", "green")
            messagebox.showinfo("완료", f"{len(records)}개 상품이 저장되었습니다.\n\n{save_path}")
        except Exception as exc:
            self.progress.stop()
            self.btn_run.config(state="normal")
            self._set_status(f"오류: {exc}", "red")
            messagebox.showerror("오류 발생", str(exc))

    def _set_status(self, msg: str, color: str = "gray") -> None:
        self.var_status.set(msg)
        self.lbl_status.config(foreground=color)
        self.update_idletasks()