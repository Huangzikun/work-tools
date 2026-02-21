"""
日志配置
"""
import logging
import sys
from pathlib import Path
from typing import Optional


def setup_logging(
    level: str = "INFO",
    log_file: Optional[str] = None,
    log_dir: str = "logs"
):
    """
    配置日志系统

    Args:
        level: 日志级别
        log_file: 日志文件名（如果为None则不输出到文件）
        log_dir: 日志文件目录
    """
    # 转换日志级别
    numeric_level = getattr(logging, level.upper(), logging.INFO)

    # 创建根logger
    root_logger = logging.getLogger()
    root_logger.setLevel(numeric_level)

    # 清除已有的handlers
    root_logger.handlers.clear()

    # 创建格式化器
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

    # 控制台handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(numeric_level)
    console_handler.setFormatter(formatter)
    root_logger.addHandler(console_handler)

    # 文件handler（如果指定）
    if log_file:
        log_path = Path(log_dir)
        log_path.mkdir(parents=True, exist_ok=True)

        file_handler = logging.FileHandler(
            log_path / log_file,
            encoding='utf-8'
        )
        file_handler.setLevel(numeric_level)
        file_handler.setFormatter(formatter)
        root_logger.addHandler(file_handler)


def get_logger(name: str) -> logging.Logger:
    """
    获取logger实例

    Args:
        name: logger名称

    Returns:
        logger实例
    """
    return logging.getLogger(name)
