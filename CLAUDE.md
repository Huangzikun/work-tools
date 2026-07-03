# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is an educational management system for handling student assignments, grading, and documentation. The system includes:
- Vue3 + Flask 全栈 Web 应用（`web/` + `backend/`，脚手架阶段，仅含登录鉴权）
- Command-line tools for OBE (Outcome-Based Education) workflows
- AI-assisted grading using VolcEngine Ark API
- Document processing for Word files (.doc/.docx)

## Architecture

### Core Components

#### 1. Web Application（`web/` + `backend/`）

##### Frontend (`web/`) - 基于 soybean-admin
- **技术栈**：Vue 3 + Vite + Naive UI + Pinia + TypeScript + UnoCSS
- **包管理器**：pnpm >= 10.5.0，Node >= 20.19.0
- **开发端口**：http://localhost:9530
- **来源**：克隆自 https://github.com/soybeanJS/soybean-admin.git（克隆后已删除其 `.git`）
- **路由系统**：基于 elegant-router 的文件路由
- **后端代理**：通过 `.env.dev` 中的 `VITE_SERVICE_BASE_URL` 指向 Flask

##### Backend (`backend/`) - Flask 工厂模式
- **入口**：`app.py` → `create_app()` 工厂函数
- **配置**：`config.py`（DB URI、JWT 密钥、CORS 白名单、默认 admin）
- **扩展**：`extensions.py`（`db = SQLAlchemy()`、`migrate`、`cors`）
- **Models**：`models/user.py` - `User`（单表 `sys_user`，roles/buttons 用 JSON 字段）
- **Routes**：
  - `routes/auth.py` - `/api/auth/login`、`/api/auth/getUserInfo`、`/api/auth/refreshToken`
  - `routes/route.py` - `/api/route/getUserRoutes`、`/api/route/getConstantRoutes`、`/api/route/isRouteExist`
  - `routes/obe.py` - `/api/obe/mkdir`（OBE 目录生成，multipart 上传名单 + ZIP 流式下载）
- **Services**：`services/auth_service.py`、`services/route_service.py`、`services/obe_mkdir_service.py`
- **Utils**：JWT 工具、Werkzeug 密码哈希、统一响应封装
- **Seed**：`seed/menus.py` - 固定菜单（home、user-center、obe > obe_mkdir）
- **Tests**：`tests/`（pytest 单元测试 + `tests/e2e/` Playwright 端到端测试）
- **端口**：http://localhost:5001（macOS 5000 被 AirPlay Receiver 占用）
- **环境**：conda env `teacherrecruitment`
- **数据库**：MariaDB `teacher_recruitment`（localhost:3306, root/root123）

#### 2. Command-Line Tools

##### OBE_MKDIR - Directory Structure Creation
- **File**: `obe_mkdir/obe_mkdir.py`
- **Purpose**: Creates standardized directory structures for courses
- **Key Features**:
  - Reads student roster from `名单.xlsx`
  - Creates directories for: 教学课件, 教学教案, and custom directories
  - Names directories with format: `{班级}《{课程}》{目录类型}{教师姓名}{人数}份`

##### OBE_CP - File Copy & Migration
- **File**: `obe_cp/obe_cp.py`
- **Purpose**: Copies student files from old to new directory structures
- **Key Features**:
  - SHA256-based duplicate detection
  - ZIP file extraction and processing
  - Intelligent file matching by student ID/name

##### OBE_CP_SIGN - AI-Assisted Grading
- **File**: `obe_cp_sign/obe_cp_sign.py`
- **Purpose**: Automated grading and signing of student reports
- **Key Features**:
  - Uses VolcEngine Ark API for AI grading
  - Processes .doc/.docx files
  - Adds teacher review sections and digital signatures
  - Generates grade summaries in Excel format

##### Teaching Plan Editor
- **File**: `teaching_plan/edit_doc.py`
- **Purpose**: AI-assisted teaching plan generation from templates
- **Key Features**:
  - Uses Ark API for content enhancement
  - Template-based document generation
  - Dynamic placeholder replacement

## Development Commands

### Setup & Installation

#### Web App - Backend (`backend/`)
```bash
conda activate teacherrecruitment
pip install -r backend/requirements.txt
python backend/init_db.py   # 建库 + 建表 + 写入默认 admin 用户
```

#### Web App - Frontend (`web/`)
```bash
cd web
pnpm install
pnpm dev   # http://localhost:9530
```

#### Command-Line Tools（共用 LLM 客户端）
```bash
pip install -r common/requirements.txt
pip install -r obe_mkdir/requirements.txt
pip install -r obe_cp/requirements.txt
pip install -r sign/requirements.txt
pip install -r sign_by_score/requirements.txt
```

#### Environment Variables
```bash
export ARK_API_KEY=your_volcengine_ark_api_key
```

### Running Applications

#### Web App 启动顺序
1. **MariaDB**：`mysql.server start`（或确认 localhost:3306 可连）
2. **Backend**：`conda activate teacherrecruitment && python backend/app.py`（监听 5001）
3. **Frontend**：`cd web && pnpm dev`（监听 9530）
4. **浏览器**：访问 http://localhost:9530，用 `admin / 123456` 登录

#### Command-Line Tools

##### Create Directory Structure
```bash
python obe_mkdir/obe_mkdir.py \
  --class_name="2022数据科学与大数据技术2班" \
  --course_name="数据结构" \
  --teacher_name="张三" \
  --need_mkdir_str="课程考核、实验实训报告、期末试卷" \
  --directory=/path/to/work/directory
```

##### Copy Student Files
```bash
python obe_cp/obe_cp.py \
  --old_dir=/path/to/old/student/files \
  --new_dir=/path/to/new/structure \
  --directory=/path/to/roster/directory
```

##### AI Grading & Signing
```bash
python obe_cp_sign/obe_cp_sign.py \
  --old_dir=/path/to/student/reports \
  --new_dir=/path/to/graded/reports \
  --directory=/path/to/roster/directory \
  --sign_picture=/path/to/teacher/signature.png \
  --sign="教师姓名" \
  --sign_date="2024-01-15"
```

##### Generate Teaching Plans
```bash
python teaching_plan/edit_doc.py
# Requires temp_plan.txt with JSON lesson plan data
```

### Testing

#### Run Tests
```bash
python test.py
python teaching_plan/test.py
python check_and_sign/test.py
```

### Production Release

项目根的 `release.sh` 是一键发布脚本：本地构建前后端 → 检测 Flask-Migrate 迁移 → 上传到生产服务器 → systemd 重启 → 健康检查。整体骨架（颜色日志、`-b`/`-t`/`-k` 参数、`trap restore_branch EXIT`、双重确认迁移）参考了 Plutus 项目的 `release.sh`，但每个 Step 针对本项目技术栈改写：

- **后端**：Flask 源码（无构建） + gunicorn（systemd 管理 `work-tools.service`），监听 `127.0.0.1:5001`
- **前端**：`pnpm build` 后部署到 1Panel 站点目录，nginx 反代 `/api` → Flask
- **同址部署**：前后端共享同一个站点根 `/opt/1panel/www/sites/work-tools/index`——前端静态文件直接放根目录，后端代码/venv/logs/.env.prod 放在子目录里
- **迁移**：Flask-Migrate/Alembic（查 `alembic_version.version_num`），远程通过 `venv/bin/flask db upgrade` 应用
- **进程管理**：`sudo systemctl restart work-tools`，日志走 `journalctl -u work-tools`

#### 部署目标
- **服务器**：`129.204.203.17`（与 Plutus 同台），SSH 用户 `ubuntu`，密钥默认 `~/Plutus.pem`（可用 `-k` 或环境变量 `SSH_KEY` 覆盖）
- **后端**：gunicorn `127.0.0.1:5001`，由 systemd 单元 `work-tools` 管理
- **前端**：1Panel 站点目录 `/opt/1panel/www/sites/work-tools/index`，对外端口 `9530`
- **数据库**：复用同机 MySQL 容器，独立库 `teacher_recruitment`、独立用户 `worktools`
- **远程目录**（全部在站点根下）：
  ```
  /opt/1panel/www/sites/work-tools/index/      ← 1Panel 站点根（nginx root）
    ├── index.html, assets/, ...                ← 前端静态文件（1Panel 直接服务）
    ├── backend/                                ← 后端 Flask 源码
    ├── venv/                                   ← Python 虚拟环境
    ├── logs/                                   ← gunicorn access/error 日志
    └── .env.prod                               ← 生产环境变量（含数据库密码/JWT 密钥）
  ```

#### 前置条件（首次部署必须人工完成）
1. **Flask-Migrate 已初始化**：本地执行 `cd backend && flask db init && flask db migrate -m "initial schema"`，提交 `backend/migrations/` 到仓库
2. **1Panel 站点已创建**：站点根目录设为 `/opt/1panel/www/sites/work-tools/index`，并配置反向代理 `/api/*` → `http://127.0.0.1:5001/api/*`
3. **⚠️ nginx 安全配置**（同址部署的关键要求）：站点根下有 `backend/`、`venv/`、`logs/`、`.env.prod`，**nginx 默认会暴露整个站点根，必须显式拒绝访问这些路径**，否则后端源码、`.env.prod`（含数据库密码）会被任意下载。在 1Panel 站点的"伪静态规则"/"自定义 nginx 配置"里加：
   ```nginx
   location ~ ^/(backend|venv|logs)(/.*)?$ { deny all; return 403; }
   location ~ /\\.env { deny all; return 403; }
   location ~ /_bak_frontend(/.*)?$ { deny all; return 403; }
   ```
   同时追加 **上传大小限制**（OBE 目录生成的名单上传需要 ≥10MB，nginx 默认 `client_max_body_size 1m` 会拦截）：
   ```nginx
   client_max_body_size 12m;
   ```
4. **生产库已建并授权**（在 MySQL 容器内执行）：
   ```sql
   CREATE DATABASE teacher_recruitment CHARACTER SET utf8mb4;
   CREATE USER 'worktools'@'%' IDENTIFIED BY '<your_strong_password>';
   GRANT ALL ON teacher_recruitment.* TO 'worktools'@'%';
   FLUSH PRIVILEGES;
   ```
5. **`.env.prod` 占位符已替换**：脚本首次部署会写入模板（含 `<CHANGE_ME>`）到 `/opt/1panel/www/sites/work-tools/index/.env.prod` 后**中止本次部署**。手动 `sudo vi` 替换全部 `<CHANGE_ME>` 为真实值（与上一步密码一致），再次运行脚本即可继续

#### 主流程（8 步）
1. **Step 0**：前置检查（`git/ssh/scp/curl/tar/pnpm` 可用、目录存在、`migrations/versions/` 有迁移文件、服务器连通）
2. **Step 0.5**（可选）：Git 切换（`-b` 分支或 `-t` 标签），`trap EXIT` 兜底恢复
3. **Step 1**：构建前端——临时改 `web/.env.prod` 的 `VITE_SERVICE_BASE_URL=/api` → `pnpm build` → 自动还原原文件
4. **Step 2**：打包后端代码（排除 `__pycache__`/`venv`/`.env*`）
5. **Step 3**：迁移检测——本地遍历 `migrations/versions/*.py` 求 heads，对比远程 `alembic_version.version_num`，列出待执行迁移文件内容，要求 `yes` + `BACKUP` 双确认
6. **Step 4**：上传前端——**只备份/替换站点根下的静态文件**（排除 `backend`/`venv`/`logs`/`.env.prod`），旧版本备份到 `_bak_frontend/`，依赖 `cp -r /tmp/dist/. <site_root>/` 覆盖。前端文件 mode 设为 644/755，不动 owner（nginx 任何用户可读，systemd 用 ubuntu 仍能写 backend/venv/logs）
7. **Step 5**：上传后端代码 + 同步依赖——`sudo mkdir backend/venv/logs` 子目录 → `sudo chown ubuntu:ubuntu` → 创建 venv（首次） → 备份旧 backend → 解压新代码 → `pip install -r requirements.txt` → 安装 systemd 单元 → 校验 `.env.prod`
8. **Step 6**：远程 `flask db upgrade` → `sudo systemctl restart work-tools` → `systemctl status` 摘要
9. **Step 7**：健康检查——`POST http://<host>:9530/api/auth/login` 重试 30 × 5s，HTTP 200 视为成功
10. **Step 8**：清理本地 tar.gz + 恢复原始 Git 引用

#### 常用命令
```bash
./release.sh                          # 部署当前代码
./release.sh -b main                  # 部署 main 分支最新代码
./release.sh -t deploy_20260624       # 部署指定标签（detached HEAD）
./release.sh -b flask -k ~/key.pem    # 指定分支和密钥
./release.sh -h                       # 显示帮助
```

**选项：**
- `-k <密钥路径>`：指定 SSH 密钥（默认 `~/Plutus.pem`，可由 `SSH_KEY` 环境变量覆盖）
- `-b <分支名>`：部署指定分支（先 `git fetch + pull`，再构建部署）
- `-t <标签名>`：部署指定标签（detached HEAD），与 `-b` 互斥
- `-h`：显示帮助

> 脚本须在项目根目录执行（依赖 `./web`、`./backend`、`./migrations` 等相对路径）。脚本退出时会自动 `git checkout` 回原始分支/标签（`trap restore_branch EXIT`），故 `-b`/`-t` 切换的 HEAD 不会残留。

#### 失败排查
```bash
# 服务状态
ssh ubuntu@129.204.203.17 'sudo systemctl status work-tools --no-pager'

# systemd 日志（最近 100 行）
ssh ubuntu@129.204.203.17 'sudo journalctl -u work-tools -n 100 --no-pager'

# gunicorn 应用日志
ssh ubuntu@129.204.203.17 'tail -100 /opt/1panel/www/sites/work-tools/index/logs/error.log'

# 手动重启
ssh ubuntu@129.204.203.17 'sudo systemctl restart work-tools'

# 确认 .env.prod 无法通过 HTTP 访问（应返回 403）
curl -I http://129.204.203.17:9530/.env.prod
curl -I http://129.204.203.17:9530/backend/app.py
```

#### 关键改造点（脚本依赖的代码改动）
- `backend/requirements.txt` 追加 `gunicorn>=22.0.0`（生产 WSGI）
- `backend/config.py` 的 `CORS_ORIGINS` 改为读环境变量（逗号分隔），生产值由 `.env.prod` 注入
- `backend/deploy/work-tools.service` 为 systemd 单元模板，脚本会同步到 `/etc/systemd/system/`；路径硬编码为 `/opt/1panel/www/sites/work-tools/index/{backend,venv,logs,.env.prod}`
- `web/.env.prod` **不改源文件**——脚本构建时临时 patch `VITE_SERVICE_BASE_URL=/api`，构建后还原

#### 与 Plutus `release.sh` 的差异
| 维度 | Plutus | work-tools |
|---|---|---|
| 后端构建 | `mvnw clean package -DskipTests` | 无构建，tar 打包源码 |
| 后端运行 | Spring Boot JAR (`nohup java -jar`) | gunicorn + systemd |
| 迁移工具 | Flyway（查 `flyway_schema_history`） | Flask-Migrate/Alembic（查 `alembic_version`） |
| 前后端位置 | 前端 openresty 独立站点、后端 `/home/ubuntu/app` | **前后端同址** `/opt/1panel/www/sites/work-tools/index`（nginx 反代 `/api` + 显式 deny 敏感路径） |
| 前端备份策略 | 整目录 mv → `_bak`（站点根纯静态） | **逐项 mv** 站点根下静态文件到 `_bak_frontend/`，保留 `backend`/`venv`/`logs`/`.env.prod` |
| 配置注入 | `application-prod.properties` | `.env.prod` + `os.getenv()` |
| 健康检查 | `POST /auth/login` 直接打后端端口 | `POST /api/auth/login` 经 1Panel 反代 |

## File Structure Patterns

### Input Requirements
- **Student Roster**: `名单.xlsx` with columns: student_id, student_name
- **Templates**: Word documents (.docx) with {{placeholder}} syntax
- **Signatures**: PNG images for digital signatures

### Output Patterns
- **Graded Reports**: Same structure as input with added teacher review sections
- **Grade Summaries**: `实验报告3.xlsx` with student scores
- **Generated Documents**: `new_v1_{count}.docx` for teaching plans

## Key Dependencies

### Backend (`backend/`)
- **Flask**: Web framework
- **Flask-SQLAlchemy**: ORM
- **Flask-Cors**: CORS
- **Flask-Migrate**: DB migration
- **PyMySQL**: MySQL driver
- **PyJWT**: JWT 鉴权
- **Werkzeug**: 密码哈希

### Frontend (`web/`)
- **Vue 3 + Vite + Naive UI + Pinia + TypeScript + UnoCSS**
- 详细依赖见 `web/package.json`

### Command-Line Tools
- **Pandas**: Excel file processing
- **python-docx**: Word document manipulation
- **openai**: OpenAI SDK（用于调用火山引擎 Ark Responses API）

### External Tools
- **LibreOffice**: Required for .doc to .docx conversion via CLI
- **VolcEngine Ark API**: AI grading and content generation
- **MariaDB 12.2.2**: 本机数据库

## Development Notes

### 数据库约束（重要）
- **不要使用关联查询，尽可能将 SQL 拆分成多条**。如果业务场景一定需要使用关联查询，必须得到我的同意。
- 当前 `User` 模型把 `roles` / `buttons` 用 JSON 字段放在单表，避免引入角色/权限关联表。

### API Endpoints (Flask Backend)
所有接口统一前缀 `/api`，响应统一为 `{ code, msg, data }`（与 soybean-admin 前端拦截器约定）。
- `POST /api/auth/login` - 登录，入参 `{userName, password}`，出参 `{token, refreshToken}`
- `GET /api/auth/getUserInfo` - 获取当前用户信息（需 Bearer Token）
- `POST /api/auth/refreshToken` - 刷新 token
- `GET /api/route/getUserRoutes` - 获取登录用户可见菜单（需 Bearer Token）
- `GET /api/route/getConstantRoutes` - 获取常量路由
- `GET /api/route/isRouteExist?routeName=` - 检查路由是否存在
- `POST /api/obe/mkdir` - OBE 目录生成（multipart/form-data，需 Bearer Token）
  - 表单字段：`className / courseName / teacherName`（必填）、`fixedDirTypes`（JSON 数组，可空，如 `["教学课件","教学教案"]`，仅创建空目录）、`studentDirTypes`（JSON 数组，至少 1 项或 fixedDirTypes 非空，如 `["课程考核","实验实训报告"]`，每个类型创建 `{班级}《{课程}》{类型}{教师}{人数}份/` 目录并在其下为每个学生建 `{学号}{专业名}{姓名}/` 子目录）
  - 文件字段：`roster`（桂林学院上课点名册 .xls，HTML 格式，列含 序号/行政班级/学号/姓名；或标准 .xlsx）
  - 成功返回 `application/zip`（Content-Disposition 用 RFC 5987 编码中文文件名）；失败返回统一 JSON `{code, msg, data}`

### AI Integration
- 通过 `common/llm_client.py` 统一封装 `LLMClient`，基于 OpenAI SDK 的 Responses API 调用火山引擎 Ark
- 默认模型：`doubao-seed-2-0-mini-260428`
- Base URL：`https://ark.cn-beijing.volces.com/api/v3`
- Requires `ARK_API_KEY` environment variable
- Supports JSON output (`json_object`)、JSON Schema 严格模式、多模态输入（`input_text` / `input_image` / `input_file`）

### Document Processing
- Handles both .doc and .docx formats
- Automatic conversion from .doc to .docx using LibreOffice
- Template-based placeholder replacement with {{key}} syntax
