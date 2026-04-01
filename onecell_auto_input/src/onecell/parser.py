"""
parser.py — Mechanism
파일 형식별 파싱 엔진.

각 파서는 파일 경로를 받아 ProductRecord 리스트를 반환한다.
어떤 파서를 언제 쓸지(Policy)는 preset.py가 결정한다.
"""
from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass


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
_OPTION_KEYWORDS = ["색상", "사이즈", "타입", "품목"]


def _parse_option_names(raw: str) -> list[str]:
    """
    옵션명 문자열 → 속성명 리스트.
    1) 개행 구분: '색상\n사이즈' → ['색상', '사이즈']
    2) 붙은 형태 fallback: '색상사이즈' → 키워드 위치 기반 분리
    3개 이상이면 색상·사이즈 우선 2개 선택.
    """
    if not raw:
        return []
    if "\n" in raw:
        parts = [p.strip() for p in raw.split("\n") if p.strip()]
        if len(parts) >= 3:
            pri    = [k for k in ["색상", "사이즈"] if k in parts]
            others = [k for k in parts if k not in pri]
            parts  = (pri + others)[:2]
        return parts[:2]
    found = sorted([(raw.find(kw), kw) for kw in _OPTION_KEYWORDS if kw in raw])
    result = [kw for _, kw in found]
    if len(result) >= 3:
        pri    = [k for k in ["색상", "사이즈"] if k in result]
        others = [k for k in result if k not in pri]
        result = (pri + others)[:2]
    return result[:2]


def _default_val(attr_name: str, val: str) -> str:
    """빈 속성값 → ONE COLOR / ONE SIZE 기본값."""
    if val:
        return val
    return "ONE COLOR" if attr_name == "색상" else "ONE SIZE"


# ─────────────────────────────────────────────
# 추상 기반 클래스
# ─────────────────────────────────────────────
class BaseParser(ABC):
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
        from bs4 import NavigableString, Tag as BSTag
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
        from bs4 import BeautifulSoup
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

            raw_opt_name = col(cells, "옵션명")
            raw_opt_val  = col(cells, "옵션값")

            opt_names = _parse_option_names(raw_opt_name)
            opt_vals  = (
                [p.strip() for p in raw_opt_val.split("\n") if p.strip()]
                if "\n" in raw_opt_val
                else ([raw_opt_val.strip()] if raw_opt_val.strip() else [])
            )
            while len(opt_vals) < len(opt_names):
                opt_vals.append("")

            an1 = opt_names[0] if opt_names else "색상"
            av1 = _default_val(an1, opt_vals[0] if opt_vals else "")
            an2 = opt_names[1] if len(opt_names) > 1 else ""
            av2 = _default_val(an2, opt_vals[1] if len(opt_vals) > 1 else "") if an2 else ""

            try:
                buy = float(col(cells, "판매가") or 0)
            except ValueError:
                buy = 0.0

            records.append(ProductRecord(
                product_name    = name,
                buy_price       = buy,
                attr_name1      = an1,
                attr_val1       = av1,
                attr_name2      = an2,
                attr_val2       = av2,
                stock           = col(cells, "재고수량") or "999",
                seller_code     = col(cells, "판매자 상품코드"),
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
        import xlrd
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
            name = cv(r, "name").strip()
            if not name:
                continue

            an1 = cv(r, "opt_n1") or "색상"
            av1 = _default_val(an1, cv(r, "opt_v1"))
            an2 = cv(r, "opt_n2")
            av2 = _default_val(an2, cv(r, "opt_v2")) if an2 else ""

            try:
                price_col = idx.get("price")
                buy = float(ws.cell_value(r, price_col)) if price_col is not None else 0.0
            except (ValueError, TypeError):
                buy = 0.0

            records.append(ProductRecord(
                product_name    = name,
                buy_price       = buy,
                attr_name1      = an1,
                attr_val1       = av1,
                attr_name2      = an2,
                attr_val2       = av2,
                stock           = "999",
                seller_code     = cv(r, "code"),
                brand           = "상세설명참조",
                manufacturer    = "상세설명참조",
                detail_html     = cv(r, "detail"),
                notice_category = "의류",
                image_url       = cv(r, "image"),
            ))
        return records