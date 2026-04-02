"""
preset.py — Policy
데이터 제공 출처별 파싱 방법 정의.

"어떤 파서를 쓰고, 배송비는 얼마이며, 파일 형식은 무엇인가"를 결정한다.
Mechanism(parser.py)은 이 파일을 알지 못한다 — 단방향 의존.

새 업체 추가 방법:
  1. PresetConfig 인스턴스를 아래에 정의
  2. REGISTRY에 등록
"""
from __future__ import annotations

from dataclasses import dataclass, field

from onecell.parser import BaseParser, SudoParser, ModuParser


# ─────────────────────────────────────────────
# 프리셋 설정 모델
# ─────────────────────────────────────────────
@dataclass
class PresetConfig:
    """단일 프리셋의 모든 정책 정보."""
    label:           str              # UI 표시명
    parser_class:    type[BaseParser] # 사용할 파서 클래스
    settings_key:    str              # settings.ini 섹션명 (배송비 등)
    default_fee:     float            # 배송비 기본값 (원)
    file_types:      list[tuple[str, str]] = field(default_factory=lambda: [
        ("XLS 파일", "*.xls"), ("모든 파일", "*.*")
    ])

    def make_parser(self) -> BaseParser:
        """label을 seller_prefix로 내장한 파서 인스턴스를 반환.
        app.py는 prefix를 알 필요 없음 — Policy가 Mechanism에 주입.
        """
        return self.parser_class(seller_prefix=self.label)


# ─────────────────────────────────────────────
# 프리셋 정의
# ─────────────────────────────────────────────
SUDO = PresetConfig(
    label        = "스도",
    parser_class = SudoParser,
    settings_key = "preset_sudo",
    default_fee  = 3000.0,
)

MODU = PresetConfig(
    label        = "모두",
    parser_class = ModuParser,
    settings_key = "preset_modu",
    default_fee  = 3300.0,
)


# ─────────────────────────────────────────────
# 레지스트리 (UI가 참조하는 단일 진입점)
# ─────────────────────────────────────────────
REGISTRY: dict[str, PresetConfig] = {
    SUDO.label: SUDO,
    MODU.label: MODU,
}