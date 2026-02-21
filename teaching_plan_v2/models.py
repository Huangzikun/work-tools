"""
数据模型定义
"""
from dataclasses import dataclass, field
from typing import Dict, Any, List


@dataclass
class LessonPlan:
    """单个教案数据模型"""
    课次: int
    授课内容: str
    授课学时: int = 5
    知识目标: str = ""
    能力目标: str = ""
    情感与价值观目标: str = ""
    重点: str = ""
    难点: str = ""
    教学工具或资源: str = ""
    课堂导入: str = ""
    课堂导入时间分配: str = "10"
    课堂内容: str = ""
    课堂内容时间分配: str = "180"
    教学方法与设计: str = ""
    课堂小结: str = ""
    课堂小结时间分配: str = "10"
    课后作业: str = ""
    课后作业时间分配: str = ""
    教学反思: str = ""

    def to_dict(self) -> Dict[str, Any]:
        """转换为字典，供模板替换使用"""
        return {
            '课次': self.课次,
            '授课内容': self.授课内容,
            '授课学时': self.授课学时,
            '知识目标': self.知识目标,
            '能力目标': self.能力目标,
            '情感与价值观目标': self.情感与价值观目标,
            '重点': self.重点,
            '难点': self.难点,
            '教学工具或资源': self.教学工具或资源,
            '课堂导入': self.课堂导入,
            '课堂导入时间分配': self.课堂导入时间分配,
            '课堂内容': self.课堂内容,
            '课堂内容时间分配': self.课堂内容时间分配,
            '教学方法与设计': self.教学方法与设计,
            '课堂小结': self.课堂小结,
            '课堂小结时间分配': self.课堂小结时间分配,
            '课后作业': self.课后作业,
            '课后作业时间分配': self.课后作业时间分配,
            '教学反思': self.教学反思,
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'LessonPlan':
        """从字典创建实例"""
        return cls(
            课次=data.get('课次', 0),
            授课内容=data.get('授课内容', ''),
            授课学时=data.get('授课学时', 5),
            知识目标=data.get('知识目标', ''),
            能力目标=data.get('能力目标', ''),
            情感与价值观目标=data.get('情感与价值观目标', ''),
            重点=data.get('重点', ''),
            难点=data.get('难点', ''),
            教学工具或资源=data.get('教学工具或资源', ''),
            课堂导入=data.get('课堂导入', ''),
            课堂导入时间分配=data.get('课堂导入时间分配', '10'),
            课堂内容=data.get('课堂内容', ''),
            课堂内容时间分配=data.get('课堂内容时间分配', '180'),
            教学方法与设计=data.get('教学方法与设计', ''),
            课堂小结=data.get('课堂小结', ''),
            课堂小结时间分配=data.get('课堂小结时间分配', '10'),
            课后作业=data.get('课后作业', ''),
            课后作业时间分配=data.get('课后作业时间分配', ''),
            教学反思=data.get('教学反思', ''),
        )

    def __str__(self) -> str:
        """字符串表示"""
        return f"第{self.课次}课: {self.授课内容}"
