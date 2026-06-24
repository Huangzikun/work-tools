import os
import sys
from pathlib import Path

# 让 backend/ 目录可被 import（测试从 backend/tests/ 下运行时）
BACKEND_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BACKEND_DIR))

FIXTURES_DIR = Path(__file__).resolve().parent / "fixtures"
ROSTER_FIXTURE = FIXTURES_DIR / "test_roster.xls"
