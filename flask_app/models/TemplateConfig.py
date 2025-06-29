# models.py
from flask_app import db  # 从 init.py 导入数据库实例

class TemplateConfig(db.Model):
    __tablename__ = 'template_config'  # 对应数据库表名
    id = db.Column(db.Integer, primary_key=True)  # 主键
    # 其他字段（根据实际表结构添加）
    template_name = db.Column(db.String())
    template_config = db.Column(db.Text)
    template_file_path = db.Column(db.Text)
    create_time = db.Column(db.DateTime)
    update_time = db.Column(db.DateTime)

    def to_dict(self):
        """将模型对象转为字典（方便返回 JSON）"""
        # 过滤掉 SQLAlchemy 内部使用的属性
        to_dict = {key: value for key, value in self.__dict__.items() if not key.startswith('_')}
        to_dict['create_time'] = self.create_time.strftime('%Y-%m-%d %H:%M:%S')
        to_dict['update_time'] = self.update_time.strftime('%Y-%m-%d %H:%M:%S')
        return to_dict