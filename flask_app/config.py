# config.py
import os

class Config:
    # 数据库 URI（格式：数据库类型+驱动://用户:密码@主机:端口/数据库名）
    SQLALCHEMY_DATABASE_URI = ''
    SQLALCHEMY_TRACK_MODIFICATIONS = False  # 关闭修改追踪（减少性能消耗）