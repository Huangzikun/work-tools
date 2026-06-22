# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is an educational management system for handling student assignments, grading, and documentation. The system includes:
- Flask-based web application for template management
- Command-line tools for OBE (Outcome-Based Education) workflows
- AI-assisted grading using VolcEngine Ark API
- Document processing for Word files (.doc/.docx)

## Architecture

### Core Components

#### 1. Flask Web Application (`flask_app/`)  
- **Entry Point**: `app.py` - Simple Flask app runner
- **Main App**: `flask_app/__init__.py` - Initializes Flask with SQLAlchemy, CORS, and logging
- **Models**: `flask_app/models/` - TemplateConfig and TemplateInfo for managing document templates
- **Routes**: 
  - `flask_app/routes/routes.py` - Template management API endpoints
  - `flask_app/obe/routes.py` - OBE-related endpoints (if implemented)
- **Services**: `flask_app/services/TemplateService.py` - Business logic for template operations
- **Utils**: `flask_app/utils/WordReplacer.py` - Word document processing utilities

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

#### Install Dependencies
```bash
# Flask app dependencies
pip install flask flask-sqlalchemy flask-cors pandas python-docx openai

# Common shared dependencies (统一 LLM 客户端)
pip install -r common/requirements.txt

# Individual tool dependencies
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

#### Flask Web Application
```bash
# Development server
python app.py
# Runs on http://localhost:5000 with debug mode enabled
```

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
# Test individual components
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

### Core Libraries
- **Flask**: Web framework
- **Pandas**: Excel file processing
- **python-docx**: Word document manipulation
- **openai**: OpenAI SDK（用于调用火山引擎 Ark Responses API）
- **SQLAlchemy**: Database ORM

### External Tools
- **LibreOffice**: Required for .doc to .docx conversion via CLI
- **VolcEngine Ark API**: AI grading and content generation

## Development Notes

### API Endpoints (Flask)
- `GET|POST /template/config/list` - List template configurations
- `GET|POST /template/info/list` - List template information
- `POST /template/info/get` - Get specific template info
- `POST /template/info/save` - Save template information
- `POST|GET /template/info/output` - Export template as document

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