import json
from docx import Document
from volcenginesdkarkruntime import Ark
import os
need_AI = True

AI_keys = ['课堂内容']

def AI_help(key, value):
    if not need_AI:
        return value

    if key not in AI_keys:
        return value

    completion = client.chat.completions.create(
        # 指定您创建的方舟推理接入点 ID，此处已帮您修改为您的推理接入点 ID
        model="doubao-seed-1-6-250615",
        messages=[
            {"role": "system",
             "content": "你是桂林学院信息工程学院的一名计算机专任教师，你拥有丰富的教学经验。"},
            {"role": "user",
             "content": f"改写以下内容以完善其内容并符合学术规范，明确教学知识重点，选择合适的教学方法，300字以下，不需要对你的决策进行解释，只需要陈述结论。内容如下：\n{value}\n。"},
        ]
    )

    return completion.choices[0].message.content

def replace_doc_placeholders(doc_name, data, num):

    for key in data.keys():
        data[key] = AI_help(key, data[key])

    data['授课学时'] = 2
    data['课堂导入时间分配'] = 5
    data['课堂内容时间分配'] = 70
    data['课堂小结时间分配'] = 5
    """
    替换Word文档中的占位符（格式：{{key}}）为对应的值
    :param doc_name: Document对象
    :param data: 教案数据字典（单个教案）
    """

    count = len(data)

    data['课次'] = num

    # 替换表格中的占位符
    for table in doc_name.tables:
        for row in table.rows:
            for cell in row.cells:
                original_text = cell.text
                for key, value in data.items():
                    placeholder = f"{{{key}}}"  # 匹配{{课次}}格式的占位符
                    if placeholder in original_text:
                        cell.text = original_text.replace(placeholder, str(value), 1)
                        count -= 1
                        if count == 0:
                            return




# 新增：读取模板文档并处理（假设使用第一个教案填充，可根据需要调整）
if __name__ == "__main__":

    # 加载模板文档（请确保文件路径正确）
    template_path = "/Users/huangzikun/PycharmProjects/work-tools/teaching_plan/数据结构实验教案-V1.docx"
    doc = Document(template_path)

    # 请确保您已将 API Key 存储在环境变量 ARK_API_KEY 中
    # 初始化Ark客户端，从环境变量中读取您的API Key
    client = Ark(
        # 此为默认路径，您可根据业务所在地域进行配置
        base_url="https://ark.cn-beijing.volces.com/api/v3",
        # 从环境变量中获取您的 API Key。此为默认方式，您可根据需要进行修改
        api_key=os.environ.get("ARK_API_KEY"),
    )

    # 新增：解析JSON字符串为Python对象
    with open('/Users/huangzikun/PycharmProjects/work-tools/teaching_plan/temp_plan.txt', 'r') as file:
        lesson_plans = json.load(file, strict=False)

    count = 1
    for lesson_plan in lesson_plans:
        replace_doc_placeholders(doc, lesson_plan, count)
        count += 1


    # 保存新文档
    doc.save(f"new_v1_{count}.docx")
    print(f"new_v1_{count}.docx")
    count += 1
