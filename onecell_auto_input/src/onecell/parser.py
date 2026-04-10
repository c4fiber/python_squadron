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

# 관리코드 패턴: 대문자알파벳으로 시작하고 -로 연결된 세그먼트 (예: NF-74, LED-RSIS-198)
_CODE_PATTERN = re.compile(r'[A-Z]+(-[A-Z0-9]+)+')


def _clean_product_name(name: str) -> str:
    """제품명 맨 앞 또는 맨 뒤의 관리코드 패턴을 제거한다.
    앞: '^[A-Z]+(-[A-Z0-9]+)+\\s+', 뒤: '\\s+[A-Z]+(-[A-Z0-9]+)+$'
    """
    # 뒤에서 제거
    name = re.sub(r'\s+[A-Z]+(-[A-Z0-9]+)+$', '', name)
    # 앞에서 제거
    name = re.sub(r'^[A-Z]+(-[A-Z0-9]+)+\s+', '', name)
    return name.strip()


def _extract_color_size(raw_opt_name: str, raw_opt_val: str) -> tuple[str, str]:
    """
    옵션명/옵션값에서 색상값·사이즈값을 추출한다.
    - 옵션명 키워드(색상/사이즈)로 vals 인덱스를 매핑
    - vals 개수 > names 개수이면 초과분을 색상값으로 처리
      예) names=['사이즈'], vals=['딥블루','S,M,L'] → 색상=딥블루, 사이즈=S,M,L
    - 값이 없으면 ONECOLOR / ONESIZE 기본값 적용
    """
    names = [n.strip() for n in raw_opt_name.split("\n") if n.strip()] if raw_opt_name else []
    vals  = [v.strip() for v in raw_opt_val.split("\n")  if v.strip()] if raw_opt_val  else []

    color_idx = next((i for i, n in enumerate(names) if "색상" in n), None)
    size_idx  = next((i for i, n in enumerate(names) if "사이즈" in n), None)

    # vals가 names보다 많으면 초과분이 앞쪽에 색상값으로 존재
    offset = max(0, len(vals) - len(names))

    if color_idx is not None:
        color_val = vals[color_idx + offset] if (color_idx + offset) < len(vals) else ""
    elif offset > 0:
        # 옵션명에 색상 없지만 vals가 더 많음 → 앞 초과분을 색상으로
        color_val = vals[0]
    else:
        color_val = ""

    if size_idx is not None:
        size_val = vals[size_idx + offset] if (size_idx + offset) < len(vals) else ""
    elif not names and len(vals) > 1:
        size_val = vals[1]
    else:
        size_val = ""

    # 옵션명 없고 vals만 있는 경우
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
# SudoParser — HTML 위장 XLS 또는 바이너리 XLS 자동 감지
# ─────────────────────────────────────────────
_XLS_MAGIC = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'


class SudoParser(BaseParser):
    """
    스도 파일 파서.
    파일 형식: HTML 위장 XLS(UTF-8) 또는 바이너리 XLS(CP949) 자동 감지.
    컬럼 구조는 두 형식 모두 동일 — 읽기 엔진만 분기.
    """

    @staticmethod
    def _cell_text(td) -> str:
        """<br>/<br/> → \\n 변환 후 텍스트 추출."""
        parts = []
        for child in td.children:
            if isinstance(child, NavigableString):
                t = child.strip()
                if t:
                    parts.append(t)
            elif isinstance(child, BSTag) and child.name == "br":
                parts.append("\n")
        return "".join(parts).strip()

    def _make_record(self, col) -> ProductRecord | None:
        """컬럼 접근 함수 col을 받아 ProductRecord를 생성한다."""
        name = col("상품명")
        if not name:
            return None
        name = _clean_product_name(name)
        av1, av2 = _extract_color_size(col("옵션명"), col("옵션값"))
        try:
            buy = float(col("판매가") or 0)
        except ValueError:
            buy = 0.0
        code = col("판매자 상품코드")
        return ProductRecord(
            product_name    = name,
            buy_price       = buy,
            attr_name1      = ATTR_NAME1,
            attr_val1       = av1,
            attr_name2      = ATTR_NAME2,
            attr_val2       = av2,
            stock           = DEFAULT_STOCK,
            seller_code     = f"{self.seller_prefix} {code}" if self.seller_prefix and code else code,
            brand           = col("브랜드") or "상세설명참조",
            manufacturer    = col("제조사") or "상세설명참조",
            detail_html     = col("상품 상세정보"),
            notice_category = col("상품정보제공고시 품명") or "의류",
            image_url       = col("대표 이미지 파일명"),
        )

    def _parse_html(self, filepath: str) -> list[ProductRecord]:
        """HTML 위장 XLS 파싱."""
        with open(filepath, encoding="utf-8", errors="replace") as f:
            content = f.read()
        soup  = BeautifulSoup(content, "html.parser")
        table = soup.find("table")
        if not table:
            raise ValueError("스도 HTML 파일에서 테이블을 찾을 수 없습니다.")
        rows = table.find_all("tr")
        if len(rows) < 2:
            raise ValueError("스도 HTML 파일에 데이터가 없습니다.")
        header = [self._cell_text(td) for td in rows[0].find_all(["td", "th"])]

        records = []
        for tr in rows[1:]:
            cells = [self._cell_text(td) for td in tr.find_all(["td", "th"])]
            while len(cells) < len(header):
                cells.append("")
            def col(name, _cells=cells, _header=header):
                return _cells[_header.index(name)] if name in _header else ""
            rec = self._make_record(col)
            if rec:
                records.append(rec)
        return records

    def _parse_binary(self, filepath: str) -> list[ProductRecord]:
        """바이너리 XLS 파싱 (xlrd)."""
        wb = xlrd.open_workbook(filepath)
        ws = wb.sheet_by_index(0)
        if ws.nrows < 2:
            raise ValueError("스도 바이너리 파일에 데이터가 없습니다.")
        header = [str(ws.cell_value(0, c)) for c in range(ws.ncols)]

        records = []
        for r in range(1, ws.nrows):
            def col(name, _r=r, _ws=ws, _header=header):
                if name not in _header:
                    return ""
                v = _ws.cell_value(_r, _header.index(name))
                return str(v).strip() if v != "" else ""
            rec = self._make_record(col)
            if rec:
                records.append(rec)
        return records

    def parse(self, filepath: str) -> list[ProductRecord]:
        """magic byte로 형식 자동 감지 후 적절한 파서로 분기."""
        with open(filepath, "rb") as f:
            magic = f.read(8)
        if magic == _XLS_MAGIC:
            return self._parse_binary(filepath)
        return self._parse_html(filepath)


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