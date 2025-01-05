# work-tools


## 前提条件
directory=指定目录， 目录内保存一个教务系统中下载的名单文件，并重命名为`名单.xlsx`

### OBE_MKDIR
创建符合要求的目录
need_mkdir_dir=课程考核、实验实训报告、期末试卷
```shell
python obe_mkdir/obe_mkdir.py --class_name=班级名称 --course_name=课程名称 --teacher_name=教师名称 --need_mkdir_str=need_mkdir_str --directory=directory
```

### OBE_CP
用于旧版本文件格式复制到新版本格式目录中
```shell
python obe_cp/obe_cp.py --old_dir=old_dir --new_dir=new_dir --directory=directory
```


### SIGN_WITH_SCORE
给出成绩单文件和原始实验报告所在目录，自动签名并输出到指定目录。默认使用WPS。         
请确保要签名的实验报告文件名中带有学生学号。   
Office Word签名未经过测试。   

```shell
# example
# python sign_with_score.py --base_path=C:\Users\Administrator\Desktop\临时 --save_path=C:\Users\Administrator\Desktop\result --sign_picture=E:\aa\张三的签名.png --sign=张三 --sign_date=2024年12月1日  --transcript_path=C:\Users\Administrator\Desktop\xxxx实验报告.xlsx   
```
- applicatioon: 使用WPS还是WORD，默认为WPS 非必输，WORD请输入Word.Application  
- base_path: 存放成绩单的目录 必输  
- save_path: 最终结果保存路径 必输  
- sign_picture: 签名图片的路径 必输  
- sign: 教师名称 必输  
- sign_date: 签名日期 必输  
- transcript_path: 成绩单路径，成绩单必须是xlsx或xls文件 必输  
- sid_col: 学号一列在表中的列号，从1开始 非必输，默认为1，和学习通导出一致    
- score_col: 分数列名一列在表中的列号，从1开始 非必输，默认为9，和学习通导出一致    
- data_row: 数据所在第一行的行号，从1开始 非必输，默认为3，和学习通导出一致    

