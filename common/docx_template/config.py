"""
配置文件 - 管理占位符格式和替换行为
"""
from dataclasses import dataclass, field
from typing import Optional, List


@dataclass
class PlaceholderConfig:
    """占位符配置"""
    prefix: str = "{{"           # 占位符前缀
    suffix: str = "}}"           # 占位符后缀
    case_sensitive: bool = True  # 是否区分大小写


@dataclass
class FormatConfig:
    """格式保留配置"""
    preserve_run_format: bool = True    # 保留run级别格式
    preserve_paragraph_format: bool = True  # 保留段落级别格式

    # 要保留的格式属性列表
    run_format_attrs: List[str] = field(default_factory=list)
    paragraph_format_attrs: List[str] = field(default_factory=list)

    def __post_init__(self):
        if not self.run_format_attrs:
            self.run_format_attrs = [
                'bold', 'italic', 'underline', 'strike',
                'font_name', 'font_size', 'color', 'highlight_color',
                'all_caps', 'small_caps', 'superscript', 'subscript'
            ]
        if not self.paragraph_format_attrs:
            self.paragraph_format_attrs = [
                'alignment', 'line_spacing', 'space_before',
                'space_after', 'first_line_indent', 'left_indent',
                'right_indent'
            ]


@dataclass
class ReplacerConfig:
    """替换器主配置"""
    placeholder: PlaceholderConfig = field(default_factory=PlaceholderConfig)
    format: FormatConfig = field(default_factory=FormatConfig)
    replace_all: bool = True          # 是否全量替换（True=替换所有匹配项，False=仅替换首次出现）
    strict_mode: bool = False         # 严格模式:缺失键时报错
    ignore_missing: bool = True       # 忽略缺失的键
    log_level: str = "INFO"           # 日志级别
