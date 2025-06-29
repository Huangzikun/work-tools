# config.py
import os

class Config:
    # 数据库 URI（格式：数据库类型+驱动://用户:密码@主机:端口/数据库名）
    SQLALCHEMY_DATABASE_URI = 'mysql+pymysql://glc_work_tools:Glc11223344@101.37.164.15:3306/glc_work_tools'
    SQLALCHEMY_TRACK_MODIFICATIONS = False  # 关闭修改追踪（减少性能消耗）