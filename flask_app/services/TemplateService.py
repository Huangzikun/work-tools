import json

from flask_app.models.TemplateConfig import TemplateConfig
from flask_app.models.TemplateInfo import TemplateInfo
from flask_app import db  # 从 init.py 导入数据库实例
from flask_app.utils.WordReplacer import WordReplacer


class TemplateService:

    @classmethod
    def templateConfig(cls, templateName):
        # 使用 SQLAlchemy ORM 查询（自动管理连接）
        config = TemplateConfig.query.filter(TemplateConfig.template_name == templateName).first()
        return config

    @classmethod
    def templateInfo(cls, infoId):
        templateInfo = TemplateInfo.query.filter(TemplateInfo.id == infoId).first()
        templateConfig = TemplateConfig.query.filter(TemplateConfig.id == templateInfo.template_id).first()

        key_config = json.loads(templateConfig.template_config, strict=False)
        info = json.loads(templateInfo.template_info, strict=False)

        return_list = []
        for key in key_config:
            return_list.append({
                'key': key['key'],
                'field_type': key['field_type'],
                'value': info[key['key']] if key['key'] in info else '',
                'show_key': key['show_key'],
            })

        return {
            'templateId': templateInfo.template_id,
            'templateName': templateConfig.template_name,
            'course_id': templateInfo.course_id,
            'key_map': return_list,
        }

    @classmethod
    def saveTemplateInfo(cls, infoId, info):
        templateInfo = TemplateInfo.query.filter(TemplateInfo.id == infoId).first()
        if not templateInfo:
            return None

        templateInfo.template_info = info
        db.session.commit()
        return templateInfo

    @classmethod
    def outputDoc(cls, infoId):
        templateInfo = TemplateInfo.query.filter(TemplateInfo.id == infoId).first()
        if not templateInfo:
            return None
        templateConfig = TemplateConfig.query.filter(TemplateConfig.id == templateInfo.template_id).first()
        if not templateConfig:
            return None

        key_config = json.loads(templateInfo.template_info, strict=False)
        wordReplacer = WordReplacer(templateConfig.template_file_path)
        wordReplacer.replace_text(key_config)

        filePath = f"/tmp/{templateConfig.template_name}_{templateInfo.course_id}.docx"
        wordReplacer.save(filePath)
        return filePath

    @classmethod
    def configList(cls):
        # 使用 SQLAlchemy ORM 查询（自动管理连接）
        configs = TemplateConfig.query.all()
        # 提取每个配置的 id 和 template_name 字段
        template_info = [config.to_dict() for config in configs]
        return template_info

    @classmethod
    def infoList(cls):
        # 使用 SQLAlchemy ORM 查询（自动管理连接）
        configs = TemplateInfo.query.all()

        # 提取所有 template_id
        template_ids = {config.template_id for config in configs}

        # 根据 template_id 查询对应的 TemplateConfig
        template_configs = TemplateConfig.query.filter(TemplateConfig.id.in_(template_ids)).all()

        # 构建 template_id 到 template_name 的映射
        template_name_map = {config.id: config.template_name for config in template_configs}

        # 提取每个配置的信息，并添加 template_name 字段
        template_info = []
        for config in configs:
            config_dict = config.to_dict()
            config_dict['template_name'] = template_name_map.get(config.template_id, '')
            template_info.append(config_dict)

        return template_info
