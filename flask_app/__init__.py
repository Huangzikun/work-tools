# init.py
from flask import Flask
from flask_sqlalchemy import SQLAlchemy
from flask_cors import CORS
import logging
from flask_app.config import Config

print("init.py 导入成功")

app = Flask(__name__)
app.config.from_object(Config)  # 加载配置



# 配置日志
logging.basicConfig(
    level=logging.DEBUG,  # 设置日志级别
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',  # 设置日志格式
    handlers=[
        logging.FileHandler('app.log'),  # 输出到文件
        logging.StreamHandler()  # 输出到控制台
    ]
)

# 获取日志记录器
logger = logging.getLogger(__name__)

db = SQLAlchemy()  # 全局数据库实例（未绑定应用）
db.init_app(app)  # 绑定应用到数据库实例

from flask_app.routes.routes import template
from flask_app.obe.routes import obe_bp

CORS(obe_bp, resources={r"/*": {"origins": "*"}})
CORS(template, resources={r"/*": {"origins": "*"}})

# 注册蓝图（示例）
app.register_blueprint(template)
app.register_blueprint(obe_bp)

