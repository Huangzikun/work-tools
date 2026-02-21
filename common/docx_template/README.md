# DOCX模板替换通用组件

一个基于python-docx的Word模板替换工具，支持保留原格式的批量占位符替换。

## 功能特性

- 支持 `{{key}}` 占位符语法
- 支持段落和表格中的占位符替换
- 完整的格式保留（15+字体属性 + 段落格式）
- 中文字体支持
- 支持全量替换和首次替换两种模式
- 详细的日志和统计信息
- 灵活的配置选项

## 安装依赖

```bash
pip install -r requirements.txt
```

## 快速开始

### 基础用法（全量替换，默认）

```python
from common.docx_template import DocxTemplateReplacer

# 创建替换器（默认全量替换）
replacer = DocxTemplateReplacer('template.docx')

# 执行替换
data = {
    '姓名': '张三',
    '学号': '202213008001',
    '课程': '数据结构',
    '课时': 64
}
replacer.replace(data, 'output.docx')

# 查看统计信息
print(replacer.get_stats())
```

### 首次替换模式

```python
from common.docx_template import DocxTemplateReplacer, ReplacerConfig

# 配置为首次替换模式
config = ReplacerConfig(
    replace_all=False  # 仅替换首次出现的占位符
)

replacer = DocxTemplateReplacer('template.docx', config)
replacer.replace(data, 'output.docx')
```

### 高级用法（带配置）

```python
from common.docx_template import DocxTemplateReplacer, ReplacerConfig

config = ReplacerConfig(
    replace_all=True,      # 全量替换
    strict_mode=True,      # 缺失键时报错
    log_level="DEBUG"      # 日志级别
)

replacer = DocxTemplateReplacer('template.docx', config)
replacer.replace(data, 'output.docx')

# 获取统计信息
stats = replacer.get_stats()
print(f"替换了{stats['placeholders_replaced']}个占位符")
```

## 模板格式

模板使用双大括号占位符:

```
学生姓名: {{姓名}}
学号: {{学号}}
课程名称: {{课程}}
课时: {{课时}}
授课教师: {{姓名}}
```

## 全量替换 vs 首次替换

### 全量替换（replace_all=True，默认）

模板内容：
```
姓名：{{姓名}}
课程：{{课程}}
授课教师：{{姓名}}
```

替换数据：`{'姓名': '张三', '课程': '数据结构'}`

结果：
```
姓名：张三
课程：数据结构
授课教师：张三
```

### 首次替换（replace_all=False）

相同模板和数据，结果：
```
姓名：张三
课程：数据结构
授课教师：{{姓名}}  # 保持不变
```

## 配置选项

### ReplacerConfig

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| replace_all | bool | True | 是否全量替换 |
| strict_mode | bool | False | 严格模式：缺失键时报错 |
| log_level | str | "INFO" | 日志级别 |

### PlaceholderConfig

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| prefix | str | "{{" | 占位符前缀 |
| suffix | str | "}}" | 占位符后缀 |
| case_sensitive | bool | True | 是否区分大小写 |

### FormatConfig

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| preserve_run_format | bool | True | 保留run级别格式 |
| preserve_paragraph_format | bool | True | 保留段落级别格式 |

## 迁移指南

从 `teaching_plan/edit_doc.py` 迁移:

```python
# 旧代码
from teaching_plan.edit_doc import replace_doc_placeholders
doc = Document('template.docx')
replace_doc_placeholders(doc, data, 1)
doc.save('output.docx')

# 新代码
from common.docx_template import DocxTemplateReplacer
replacer = DocxTemplateReplacer('template.docx')
replacer.replace(data, 'output.docx')
```

## 改进点

相比现有实现:
1. 性能优化：一次遍历完成所有替换
2. 格式保留：支持15+格式属性 + 段落格式
3. 中文字体：使用qn方法正确处理
4. 错误处理：完善的异常和日志
5. 统计信息：提供替换统计
6. 双模式：支持全量和首次替换

## 版本

当前版本：1.0.0
