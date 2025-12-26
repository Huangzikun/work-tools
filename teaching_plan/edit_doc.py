import json
from docx import Document
from volcenginesdkarkruntime import Ark
import os

need_AI = True

AI_keys = ['课堂内容']

temp_file_path = '/Users/huangzikun/Desktop/桂林学院/AI大数据在城乡规划中的应用（上）/temp_plan.txt'
save_path = '/Users/huangzikun/Desktop/桂林学院/AI大数据在城乡规划中的应用（上）/'
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
             "content": f"改写以下内容以完善其内容并符合学术规范，明确教学知识，重难点，教学方法，教学内容，确定教学步骤，融合思政元素，300字以下，应减少使用括号描述具体项目或举例的描述。"
                        f"如果有多项列表、添加换行符以保证格式。"
                        f"避免使用markdown的标签，如*、**、#等"
                        f"不需要对你的决策进行解释，只需要陈述结论。内容如下：\n{value}\n。"},
        ],
        max_tokens=32000,
    )

    return completion.choices[0].message.content

def replace_doc_placeholders(doc_name, data, num):

    for key in data.keys():
        data[key] = AI_help(key, data[key])

    data['授课学时'] = 5
    data['课堂导入时间分配'] = 5
    data['课堂内容时间分配'] = 190
    data['课堂小结时间分配'] = 5
    data['课次'] = num

    print(data['课次'])
    """
    替换Word文档中的占位符（格式：{{key}}）为对应的值，保留原有格式
    :param doc_name: Document对象
    :param data: 教案数据字典（单个教案）
    """

    count = len(data)


    def replace_text_in_paragraph(paragraph, placeholder, replacement):
        """在段落中替换文本，保留格式"""
        if placeholder not in paragraph.text:
            return False
            
        # 遍历运行(run)来替换文本
        full_text = ""
        runs_info = []
        
        # 收集所有文本和格式信息
        for run in paragraph.runs:
            runs_info.append({
                'text': run.text,
                'bold': run.font.bold,
                'italic': run.font.italic,
                'underline': run.font.underline,
                'font_name': run.font.name,
                'font_size': run.font.size,
                'color': run.font.color.rgb if run.font.color else None
            })
            full_text += run.text
        
        if placeholder not in full_text:
            return False
            
        # 清除原有运行
        for run in paragraph.runs:
            run.clear()
            
        # 重新构建文本，保留格式
        new_text = full_text.replace(placeholder, str(replacement), 1)
        
        # 创建新的运行
        if runs_info:
            # 使用第一个运行的格式作为默认
            first_run = paragraph.add_run(new_text)
            if runs_info[0]['bold'] is not None:
                first_run.font.bold = runs_info[0]['bold']
            if runs_info[0]['italic'] is not None:
                first_run.font.italic = runs_info[0]['italic']
            if runs_info[0]['underline'] is not None:
                first_run.font.underline = runs_info[0]['underline']
            if runs_info[0]['font_name']:
                first_run.font.name = runs_info[0]['font_name']
            if runs_info[0]['font_size']:
                first_run.font.size = runs_info[0]['font_size']
            if runs_info[0]['color']:
                first_run.font.color.rgb = runs_info[0]['color']
        else:
            # 如果没有格式信息，直接添加文本
            paragraph.add_run(new_text)
            
        return True

    # 替换表格中的占位符
    for table in doc_name.tables:
        for row in table.rows:
            for cell in row.cells:
                # 处理单元格中的每个段落
                for paragraph in cell.paragraphs:
                    for key, value in data.items():
                        placeholder = f"{{{key}}}"  # 匹配{{课次}}格式的占位符
                        if replace_text_in_paragraph(paragraph, placeholder, str(value)):
                            count -= 1
                            if count == 0:
                                return

    # 处理文档主体中的占位符（非表格部分）
    for paragraph in doc_name.paragraphs:
        for key, value in data.items():
            placeholder = f"{{{key}}}"
            replace_text_in_paragraph(paragraph, placeholder, str(value))




# 新增：读取模板文档并处理（假设使用第一个教案填充，可根据需要调整）
if __name__ == "__main__":

    # 加载模板文档（请确保文件路径正确）
    template_path = "/Users/huangzikun/Desktop/桂林学院/AI大数据在城乡规划中的应用（上）/理论课教案-V1.docx"
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
    with open(temp_file_path, 'r') as file:
        lesson_plans = json.load(file, strict=False)

    count = 1
    for lesson_plan in lesson_plans:
        replace_doc_placeholders(doc, lesson_plan, count)
        print(f"课次: {count}")
        count += 1


    # 保存新文档
    doc.save(f"{save_path}/new_v1_{count}.docx")
    print(f"{save_path}/new_v1_{count}.docx")
    count += 1
