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


