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
- **开发端口**：http://localhost:9527
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
- **Services**：`services/auth_service.py`、`services/route_service.py`
- **Utils**：JWT 工具、Werkzeug 密码哈希、统一响应封装
- **Seed**：`seed/menus.py` - 固定菜单（home、user-center）
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
pnpm dev   # http://localhost:9527
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
3. **Frontend**：`cd web && pnpm dev`（监听 9527）
4. **浏览器**：访问 http://localhost:9527，用 `admin / 123456` 登录

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
