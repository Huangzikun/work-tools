"""
配置管理
"""
import os
from dataclasses import dataclass, field
from typing import Optional
from pathlib import Path


@dataclass
class AppConfig:
    """应用配置"""

    # AI配置
    api_key: str = ""
    model: str = "doubao-seed-2-0-mini-260428"  # 支持Responses API的模型
    timeout: int = 1800  # 30分钟超时（秒）
    max_retries: int = 2  # 失败重试次数

    # 模板和输出配置
    template_path: str = ""
    output_dir: str = "generated_plans"

    # 解析配置
    batch_size: int = 2  # 每批生成的教案数

    # 日志配置
    log_level: str = "INFO"

    @classmethod
    def from_env(cls) -> 'AppConfig':
        """从环境变量加载配置"""
        api_key = os.environ.get("ARK_API_KEY", "")
        return cls(api_key=api_key)

    def validate(self) -> None:
        """验证配置"""
        if not self.api_key:
            raise ValueError("API Key未设置，请设置ARK_API_KEY环境变量")

        if not self.template_path:
            raise ValueError("模板路径未设置")

        if not Path(self.template_path).exists():
            raise ValueError(f"模板文件不存在: {self.template_path}")

    def update_from_args(self, args) -> None:
        """从命令行参数更新配置

        Args:
            args: argparse解析后的参数对象
        """
        if hasattr(args, 'template') and args.template:
            self.template_path = args.template

        if hasattr(args, 'output_dir') and args.output_dir:
            self.output_dir = args.output_dir

        if hasattr(args, 'batch_size') and args.batch_size:
            self.batch_size = args.batch_size

        if hasattr(args, 'model') and args.model:
            self.model = args.model

        if hasattr(args, 'log_level') and args.log_level:
            self.log_level = args.log_level
