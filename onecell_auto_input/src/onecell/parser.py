"""
parser.py — Mechanism
파일 형식별 파싱 엔진.

각 파서는 파일 경로를 받아 ProductRecord 리스트를 반환한다.
어떤 파서를 언제 쓸지(Policy)는 preset.py가 결정한다.
"""
from __future__ import annotations

import re
from abc import ABC, abstractmethod
from dataclasses import dataclass

# 최상단 import: PyInstaller 정적 분석이 lazy import를 놓치지 않도록
import xlrd                                    # ModuParser
from bs4 import BeautifulSoup                  # SudoParser
from bs4 import NavigableString, Tag as BSTag  # SudoParser._cell_text


# ─────────────────────────────────────────────
# 공통 데이터 모델
# ─────────────────────────────────────────────
@dataclass
class ProductRecord:
    """파서 출력 정규화 모델. 프리셋/파일형식에 무관하게 동일한 구조."""
    product_name:    str
    buy_price:       float
    attr_name1:      str
    attr_val1:       str
    attr_name2:      str
    attr_val2:       str
    stock:           str
    seller_code:     str
    brand:           str
    manufacturer:    str
    detail_html:     str
    notice_category: str
    image_url:       str


# ─────────────────────────────────────────────
# 공통 유틸리티
# ─────────────────────────────────────────────
ATTR_NAME1 = "색상"
ATTR_NAME2 = "사이즈"
DEFAULT_VAL1 = "ONECOLOR"
DEFAULT_VAL2 = "ONESIZE"
DEFAULT_STOCK = "999"

# 제품명 끝 관리코드 패턴: 공백 + 대문자알파벳-숫자 (예: NF-74, SD-251014, URK-274)
_TRAILING_CODE = re.compile(r'\s+[A-Z]+(-[A-Z0-9]+)+$')


def _clean_product_name(name: str) -> str:
    """제품명 끝의 관리코드 패턴([A-Z]+-[0-9]+(-[0-9]+)*)을 제거한다."""
    return _TRAILING_CODE.sub('', name).strip()


def _extract_color_size(raw_opt_name: str, raw_opt_val: str) -> tuple[str, str]:
    """
    옵션명/옵션값 원본에서 색상값·사이즈값을 추출한다.
    - 속성명1은 항상 "색상", 속성명2는 항상 "사이즈"로 고정
    - 옵션명이 개행 구분(\'색상\n사이즈\')이면 순서대로 색상/사이즈 매핑
    - 옵션값이 개행 구분이면 첫 번째 → 색상값, 두 번째 → 사이즈값
    - 값이 없으면 ONECOLOR / ONESIZE 기본값 적용
    """
    names = [n.strip() for n in raw_opt_name.split("\n") if n.strip()] if raw_opt_name else []
    vals  = [v.strip() for v in raw_opt_val.split("\n")  if v.strip()] if raw_opt_val  else []

    # 옵션명 위치로 색상/사이즈 인덱스 결정
    color_idx = next((i for i, n in enumerate(names) if "색상" in n), None)
    size_idx  = next((i for i, n in enumerate(names) if "사이즈" in n), None)

    color_val = vals[color_idx] if color_idx is not None and color_idx < len(vals) else ""
    size_val  = vals[size_idx]  if size_idx  is not None and size_idx  < len(vals) else ""

    # 옵션명 구분 없이 값만 있는 경우 (단일 값): 첫 번째를 색상으로
    if not names and vals:
        color_val = vals[0]
        size_val  = vals[1] if len(vals) > 1 else ""

    return (color_val or DEFAULT_VAL1, size_val or DEFAULT_VAL2)


# ─────────────────────────────────────────────
# 추상 기반 클래스
# ─────────────────────────────────────────────
class BaseParser(ABC):
    def __init__(self, seller_prefix: str = "") -> None:
        """seller_prefix: 판매자 관리코드 앞에 붙일 프리셋 키워드 (예: "스도", "모두")
        preset.make_parser()가 주입 — 파서는 Policy를 알 필요 없음.
        """
        self.seller_prefix = seller_prefix

    @abstractmethod
    def parse(self, filepath: str) -> list[ProductRecord]:
        """파일을 읽어 ProductRecord 리스트를 반환한다."""


# ─────────────────────────────────────────────
# SudoParser — HTML 위장 XLS (UTF-8, BeautifulSoup)
# ─────────────────────────────────────────────
class SudoParser(BaseParser):
    """
    스도 파일 파서.
    파일 형식: HTML 위장 XLS (확장자는 .xls, 내용은 HTML UTF-8).
    옵션명/옵션값: 단일 셀에 <br/>로 구분된 복수 행.
    """

    @staticmethod
    def _cell_text(td) -> str:
        """<br>/<br/> → \\n 변환 후 텍스트 추출. strip=True는 \\n을 제거하므로 직접 순회."""
        parts = []
        for child in td.children:
            if isinstance(child, NavigableString):
                t = child.strip()
                if t:
                    parts.append(t)
            elif isinstance(child, BSTag) and child.name == "br":
                parts.append("\n")
        return "".join(parts).strip()

    def parse(self, filepath: str) -> list[ProductRecord]:
        with open(filepath, encoding="utf-8", errors="replace") as f:
            content = f.read()
        soup  = BeautifulSoup(content, "html.parser")
        table = soup.find("table")
        if not table:
            raise ValueError("스도 파일에서 테이블을 찾을 수 없습니다.")
        rows = table.find_all("tr")
        if len(rows) < 2:
            raise ValueError("스도 파일에 데이터가 없습니다.")

        header = [self._cell_text(td) for td in rows[0].find_all(["td", "th"])]

        def col(cells: list[str], name: str) -> str:
            return cells[header.index(name)] if name in header else ""

        records = []
        for tr in rows[1:]:
            cells = [self._cell_text(td) for td in tr.find_all(["td", "th"])]
            while len(cells) < len(header):
                cells.append("")

            name = col(cells, "상품명")
            if not name:
                continue
            name = _clean_product_name(name)

            av1, av2 = _extract_color_size(
                col(cells, "옵션명"), col(cells, "옵션값")
            )

            try:
                buy = float(col(cells, "판매가") or 0)
            except ValueError:
                buy = 0.0

            records.append(ProductRecord(
                product_name    = name,
                buy_price       = buy,
                attr_name1      = ATTR_NAME1,
                attr_val1       = av1,
                attr_name2      = ATTR_NAME2,
                attr_val2       = av2,
                stock           = DEFAULT_STOCK,
                seller_code     = f"{self.seller_prefix} {col(cells, '판매자 상품코드')}" if self.seller_prefix and col(cells, "판매자 상품코드") else col(cells, "판매자 상품코드"),
                brand           = col(cells, "브랜드") or "상세설명참조",
                manufacturer    = col(cells, "제조사") or "상세설명참조",
                detail_html     = col(cells, "상품 상세정보"),
                notice_category = col(cells, "상품정보제공고시 품명") or "의류",
                image_url       = col(cells, "대표 이미지 파일명"),
            ))
        return records


# ─────────────────────────────────────────────
# ModuParser — 바이너리 XLS (CP949, xlrd)
# ─────────────────────────────────────────────
class ModuParser(BaseParser):
    """
    모두 파일 파서.
    파일 형식: 진짜 바이너리 XLS (Microsoft Compound Document).
    옵션명/옵션값: I/J(옵션이름1/옵션값1), K/L(옵션이름2/옵션값2) 열에 분리되어 있음.
    """

    def parse(self, filepath: str) -> list[ProductRecord]:
        wb = xlrd.open_workbook(filepath)
        ws = wb.sheet_by_index(0)
        if ws.nrows < 2:
            raise ValueError("모두 파일에 데이터가 없습니다.")

        raw_header = [str(ws.cell_value(0, c)) for c in range(ws.ncols)]

        def find_col(keyword: str) -> int | None:
            for i, h in enumerate(raw_header):
                if keyword in h:
                    return i
            return None

        idx = {
            "name":   find_col("상품명"),
            "price":  find_col("판매가"),
            "detail": find_col("상품상세설명"),
            "image":  find_col("대표이미지"),
            "code":   find_col("자체상품코드"),
            "opt_n1": find_col("옵션이름1"),
            "opt_v1": find_col("옵션값1"),
            "opt_n2": find_col("옵션이름2"),
            "opt_v2": find_col("옵션값2"),
        }

        def cv(row: int, key: str) -> str:
            col_i = idx.get(key)
            if col_i is None:
                return ""
            val = ws.cell_value(row, col_i)
            return str(val).strip() if val != "" else ""

        records = []
        for r in range(1, ws.nrows):
            name = _clean_product_name(cv(r, "name").strip())
            if not name:
                continue
            name = _clean_product_name(name)

            # 모두 파일은 옵션이름1/2가 이미 분리되어 있음 → 이름 무시하고 값만 추출
            raw_opt_name = f"{cv(r, 'opt_n1')}\n{cv(r, 'opt_n2')}"
            raw_opt_val  = f"{cv(r, 'opt_v1')}\n{cv(r, 'opt_v2')}"
            av1, av2 = _extract_color_size(raw_opt_name, raw_opt_val)

            try:
                price_col = idx.get("price")
                buy = float(ws.cell_value(r, price_col)) if price_col is not None else 0.0
            except (ValueError, TypeError):
                buy = 0.0

            records.append(ProductRecord(
                product_name    = name,
                buy_price       = buy,
                attr_name1      = ATTR_NAME1,
                attr_val1       = av1,
                attr_name2      = ATTR_NAME2,
                attr_val2       = av2,
                stock           = DEFAULT_STOCK,
                seller_code     = f"{self.seller_prefix} {cv(r, 'code')}" if self.seller_prefix and cv(r, "code") else cv(r, "code"),
                brand           = "상세설명참조",
                manufacturer    = "상세설명참조",
                detail_html     = cv(r, "detail"),
                notice_category = "의류",
                image_url       = cv(r, "image"),
            ))
        return records