import os
from volcenginesdkarkruntime import Ark
from docx import Document
import json

def read_docx(file_path):
    """读取docx文件内容并返回文本"""
    doc = Document(file_path)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    # 处理表格内容
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                full_text.append(cell.text)
    return '\n'.join(full_text)


# 请确保您已将 API Key 存储在环境变量 ARK_API_KEY 中
# 初始化Ark客户端，从环境变量中读取您的API Key
client = Ark(
    # 此为默认路径，您可根据业务所在地域进行配置
    base_url="https://ark.cn-beijing.volces.com/api/v3",
    # 从环境变量中获取您的 API Key。此为默认方式，您可根据需要进行修改
    api_key=os.environ.get("ARK_API_KEY"),
)

type = "{'课次,应为一个整数，如1','授课内容,应精简描述知识点','授课学时,应为一个整数，如3','知识目标','能力目标','情感与价值观目标','重点','难点','教学工具或资源','课堂导入','课堂导入时间分配','课堂内容，应尽可能的详细、准确描述授课内容和知识点','课堂内容时间分配','教学方法与设计','课堂小结','课堂小结时间分配','课后作业','课后作业时间分配','教学反思'}"

file_dir = "/Users/huangzikun/Desktop/桂林学院/面向对象程序设计/教学大纲/2024级/《面向对象程序设计》课程教学大纲.docx"
teaching_plan_count = 17
step = 4
res_list = []


doc_content = read_docx(file_dir)

for i in range(1, teaching_plan_count+1, step):
    completion = client.chat.completions.create(
       # 指定您创建的方舟推理接入点 ID，此处已帮您修改为您的推理接入点 ID
        model="doubao-1-5-pro-256k-250115",
        messages=[
            {"role": "system", "content": "你是桂林学院信息工程学院的一名计算机专任教师，你拥有丰富的教学经验。你的任务是按照读取教学大纲，根据教学大纲中的章节和实验安排，合理分布到用户提供的课时安排中。从文件中读取以下列表中的每一项，生成json格式的内容。json应为一个数组，其中每一个对象包括后续列表内的每一项。"},
            {"role": "user", "content": f"根据文档内容中的教学时数/3确定教案的数量，如教学时数为9，即你应该给我提供9/3=3个教案的内容。总共应提供{teaching_plan_count}个教案。当前输出第{i}~{i+step-1}个教案。列表内容如下：\n{type}\n。文档内容如下：\n{doc_content}\n如果文档未提及，则对象内的这一项设置为空字符串。如果一项中包含一个列表，则每一个条目后面增加换行。"},
        ],
        response_format={
            "type": "json_object",
        },
    )

    print(completion.choices[0].message.content)
    res_list.extend(json.loads(completion.choices[0].message.content, strict=False))


# 保存为JSON格式到temp.txt（中文友好格式）
with open("temp_plan.txt", "w", encoding="utf-8") as f:
    json.dump(res_list, f, ensure_ascii=False, indent=2)