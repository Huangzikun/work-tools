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

---

## Web 应用（脚手架）

项目内嵌一个 Vue3 + Flask 全栈 Web 脚手架，位于 `web/` 与 `backend/` 目录，目前仅包含登录鉴权与基础菜单，作为后续业务模块开发的起点。

### 目录结构
- `web/`：前端，基于 [soybean-admin](https://github.com/soybeanJS/soybean-admin)（Vue3 + Vite + Naive UI + Pinia + TS）
- `backend/`：后端，Flask 工厂模式 + SQLAlchemy + PyJWT，单表 `sys_user`

### 环境要求
- Node >= 20.19.0，pnpm >= 10.5.0
- conda env `teacherrecruitment`（Python 3.13）
- MariaDB（本机 localhost:3306, root/root123）

### 启动顺序

#### 1. 启动 MariaDB 并初始化数据库（仅首次）
```bash
conda activate teacherrecruitment
pip install -r backend/requirements.txt
python backend/init_db.py
# 输出 "DB ready, admin/123456" 即成功
```

#### 2. 启动后端
```bash
conda activate teacherrecruitment
python backend/app.py
# 监听 http://localhost:5001（macOS 5000 被 AirPlay 占用）
```

#### 3. 启动前端
```bash
cd web
pnpm install
pnpm dev
# 访问 http://localhost:9530
```

### 默认账号
- 用户名：`admin`
- 密码：`123456`

### 接口一览（统一前缀 `/api`）
- `POST /api/auth/login` - 登录，返回 `{token, refreshToken}`
- `GET /api/auth/getUserInfo` - 获取用户信息（需 Bearer Token）
- `POST /api/auth/refreshToken` - 刷新 token
- `GET /api/route/getUserRoutes` - 获取登录用户可见菜单
- `GET /api/route/getConstantRoutes` - 常量路由
- `GET /api/route/isRouteExist?routeName=` - 路由存在性检查
