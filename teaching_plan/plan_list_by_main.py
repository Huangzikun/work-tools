import os
from common.llm_client import LLMClient
from docx import Document
import json

system_content = '''
你是桂林学院信息工程学院的一名计算机专任教师，你拥有丰富的教学经验。你的任务是按照读取教学大纲，根据教学大纲中的章节和实验安排，合理分布到用户提供的课时安排中。遵守以下限制：
1. 每个课时为40分钟，如果每次课时为5个课时，则总共可使用的时间分配为5*40=200分钟。
2. 课堂导入、课后总结时间不得超过10分钟。
3. 课堂内容只描述内容，不要用括号举例的形式，不要使用代码。
从文件中读取以下列表中的每一项，生成json格式的内容。json应为一个数组，其中每一个对象包括后续列表内的每一项。
'''
file_dir = "/Users/huangzikun/Desktop/桂林学院/AI大数据在城乡规划中的应用（上）/城乡规划-AI大数据在城乡规划中的应用（上）-理论教学大纲.docx"
teaching_plan_count = 7
step = 2


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


client = LLMClient()

type = "{'课次,应为一个整数，如1','授课内容,应精简描述知识点','授课学时,应为一个整数，如3','知识目标','能力目标','情感与价值观目标','重点','难点','教学工具或资源','课堂导入','课堂导入时间分配','课堂内容，应尽可能的详细、准确描述授课内容和知识点','课堂内容时间分配','教学方法与设计','课堂小结','课堂小结时间分配','课后作业','课后作业时间分配','教学反思'}"


res_list = []


doc_content = read_docx(file_dir)

for i in range(1, teaching_plan_count+1, step):
    print(i)
    output_text = client.generate(
        system_prompt=system_content,
        user_prompt=f"根据文档内容中的教学时数/5确定教案的数量，如教学时数为5，即你应该给我提供10/5=2个教案的内容。总共应提供{teaching_plan_count}个教案。当前输出第{i}~{i+step-1}个教案。列表内容如下：\n{type}\n。文档内容如下：\n{doc_content}\n如果文档未提及，则对象内的这一项设置为空字符串。如果一项中包含一个列表，则每一个条目后面增加换行。",
        json_output=True,
        max_tokens=32000,
    )

    print(output_text)
    res_list.extend(json.loads(output_text, strict=False))


# 保存为JSON格式到temp.txt（中文友好格式）
with open("/Users/huangzikun/Desktop/桂林学院/AI大数据在城乡规划中的应用（上）/temp_plan.txt", "w", encoding="utf-8") as f:
    json.dump(res_list, f, ensure_ascii=False, indent=2)