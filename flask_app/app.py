from flask import Flask
from obe.routes import obe_bp
import logging
from flask_cors import CORS

# 创建 Flask 应用实例
app = Flask(__name__)

CORS(obe_bp, resources={r"/*": {"origins": "*"}})

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


# 注册蓝图
app.register_blueprint(obe_bp)

if __name__ == '__main__':
    app.run(debug=True)