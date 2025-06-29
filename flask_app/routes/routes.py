import os

from flask import Blueprint, jsonify, request, send_file
from flask_app.services.TemplateService import TemplateService

template = Blueprint('template', __name__)

@template.route('/template/config/list', methods=['POST', 'GET'])
def list_templates():
    configs = TemplateService.configList()
    # 将模型对象转为字典（方便返回 JSON）
    return jsonify(code=0, data=configs, msg='配置获取成功')

@template.route('/template/info/list', methods=['POST', 'GET'])
def list_info():
    configs = TemplateService.infoList()
    # 将模型对象转为字典（方便返回 JSON）
    return jsonify(code=0, data=configs, msg='配置获取成功')


@template.route('/template/info/get', methods=['POST'])
def getInfo():
    infoId = request.form.get('info_id')
    # 检查参数是否存在
    if  not infoId:
        return jsonify(code=1, msg='缺少必要参数infoId')

    config = TemplateService.templateInfo(infoId)
    if config is None:
        return jsonify({'code': 1, 'msg': '配置不存在'})
    # 将模型对象转为字典（方便返回 JSON）
    return jsonify(code=0, data=config, msg='配置获取成功')

@template.route('/template/info/save', methods=['POST'])
def saveInfo():
    infoId = request.form.get('info_id')
    info = request.form.get('info')
    # 检查参数是否存在
    if not infoId:
        return jsonify(code=1, msg='缺少必要参数 infoId')

    config = TemplateService.saveTemplateInfo(infoId, info)
    if config is None:
        return jsonify({'code': 1, 'msg': '配置保存失败'})
    # 将模型对象转为字典（方便返回 JSON）
    return jsonify(code=0, data={}, msg='配置保存成功')

@template.route('/template/info/output', methods=['POST', 'GET'])
def outputDoc():
    infoId = request.values.get('info_id')
    # 检查参数是否存在
    if not infoId:
        return jsonify(code=1, msg='缺少必要参数 infoId')

    config = TemplateService.outputDoc(infoId)
    if config is None:
        return jsonify({'code': 1,'msg': '配置导出失败'})
    # 获取文件名
    file_name = os.path.basename(config)
    # 返回文件供用户下载
    return send_file(config, as_attachment=True, download_name=file_name)



