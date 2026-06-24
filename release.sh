#!/bin/bash

###############################################################################
# work-tools 项目自动化发布脚本
# 功能：本地构建前端 → 打包后端 → 检测 Flask-Migrate 迁移 → 上传部署 → 健康检查
# 后端: gunicorn (127.0.0.1:5001) 由 systemd 管理，前端由 1Panel 站点托管 + /api 反代
###############################################################################

set -euo pipefail

# ==================== 颜色输出 ====================
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m'

log_info()    { echo -e "${BLUE}[INFO]${NC} $(date '+%Y-%m-%d %H:%M:%S') - $1"; }
log_success() { echo -e "${GREEN}[SUCCESS]${NC} $(date '+%Y-%m-%d %H:%M:%S') - $1"; }
log_warning() { echo -e "${YELLOW}[WARNING]${NC} $(date '+%Y-%m-%d %H:%M:%S') - $1"; }
log_error()   { echo -e "${RED}[ERROR]${NC} $(date '+%Y-%m-%d %H:%M:%S') - $1"; }

# ==================== 参数解析 ====================
SSH_KEY="${SSH_KEY:-${HOME}/Plutus.pem}"
DEPLOY_BRANCH=""
DEPLOY_TAG=""
ORIGINAL_REF=""

restore_branch() {
    if [ -n "${ORIGINAL_REF}" ]; then
        echo -e "${YELLOW}[WARNING]${NC} 脚本退出，恢复原始引用: ${ORIGINAL_REF}"
        git checkout "${ORIGINAL_REF}" 2>/dev/null || true
    fi
}
trap restore_branch EXIT

while getopts "k:b:t:h" opt; do
    case $opt in
        k) SSH_KEY="$OPTARG" ;;
        b) DEPLOY_BRANCH="$OPTARG" ;;
        t) DEPLOY_TAG="$OPTARG" ;;
        h) echo "用法: $0 [-k 密钥路径] [-b 分支名] [-t 标签名]"
           echo ""
           echo "选项:"
           echo "  -k  指定 SSH 密钥 (默认: ~/Plutus.pem，可通过 SSH_KEY 环境变量覆盖)"
           echo "  -b  部署指定分支 (自动 fetch + pull)"
           echo "  -t  部署指定标签 (detached HEAD)"
           echo "  -h  显示帮助"
           echo ""
           echo "示例:"
           echo "  $0                        # 部署当前代码"
           echo "  $0 -b main                # 部署 main 分支最新代码"
           echo "  $0 -t deploy_20260624     # 部署指定标签"
           echo "  $0 -b flask -k ~/key.pem  # 指定分支和密钥"
           exit 0 ;;
        *) exit 1 ;;
    esac
done

if [ -n "${DEPLOY_BRANCH}" ] && [ -n "${DEPLOY_TAG}" ]; then
    log_error "不能同时指定分支 (-b) 和标签 (-t)"
    exit 1
fi

# ==================== 生产环境配置 ====================
SERVER_HOST="129.204.203.17"
SERVER_USER="ubuntu"
BACKEND_PORT=5001
FRONTEND_PORT=9530
SERVICE_NAME="work-tools"

# 远程路径（前后端同址部署在 1Panel 站点根目录下）
REMOTE_APP_ROOT="/opt/1panel/www/sites/work-tools/index"
REMOTE_BACKEND_DIR="${REMOTE_APP_ROOT}/backend"
REMOTE_VENV_DIR="${REMOTE_APP_ROOT}/venv"
REMOTE_LOG_DIR="${REMOTE_APP_ROOT}/logs"
REMOTE_ENV_FILE="${REMOTE_APP_ROOT}/.env.prod"
# REMOTE_FRONTEND_DIR 与 REMOTE_APP_ROOT 同址：前端静态文件直接放站点根，
# 后端代码/venv/logs/.env.prod 放在站点根的子目录里
REMOTE_FRONTEND_DIR="${REMOTE_APP_ROOT}"
REMOTE_FRONTEND_BAK="${REMOTE_APP_ROOT}_bak_frontend"
# 前端部署时必须保留的子目录/文件（站点根同址部署后端）
PRESERVE_NAMES="backend venv logs .env.prod"

# 本地路径
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
LOCAL_BACKEND_DIR="${SCRIPT_DIR}/backend"
LOCAL_FRONTEND_DIR="${SCRIPT_DIR}/web"
LOCAL_MIGRATION_DIR="${LOCAL_BACKEND_DIR}/migrations"
LOCAL_SERVICE_FILE="${LOCAL_BACKEND_DIR}/deploy/${SERVICE_NAME}.service"

# SSH 密钥认证（优先级：命令行 -k > 环境变量 SSH_KEY > 默认 ~/Plutus.pem）
SSH_OPTS="-o StrictHostKeyChecking=no -o ConnectTimeout=10 -i ${SSH_KEY}"

ssh_exec() {
    ssh ${SSH_OPTS} "${SERVER_USER}@${SERVER_HOST}" "$@"
}

scp_to_tmp() {
    local local_file="$1"
    local remote_name="$(basename "${local_file}")"
    scp ${SSH_OPTS} "${local_file}" "${SERVER_USER}@${SERVER_HOST}:/tmp/${remote_name}"
}

# .env.prod 内容（首次部署时写入，含 <CHANGE_ME> 占位符）
PROD_ENV_CONTENT="FLASK_SECRET_KEY=<CHANGE_ME>
JWT_SECRET_KEY=<CHANGE_ME>
DATABASE_URI=mysql+pymysql://worktools:<CHANGE_ME>@localhost:3306/teacher_recruitment?charset=utf8mb4
CORS_ORIGINS=http://${SERVER_HOST}:${FRONTEND_PORT}
DB_HOST=localhost
DB_PORT=3306
DB_USER=worktools
DB_PASSWORD=<CHANGE_ME>
DB_NAME=teacher_recruitment"

# ==================== 工具函数 ====================

check_command() {
    if ! command -v "$1" >/dev/null 2>&1; then
        log_error "$1 未安装，请先安装"
        exit 1
    fi
}

abort_release() {
    log_error "$1"
    log_info "部署已中止，生产环境未受影响"
    exit 1
}

# ==================== Step 0: 前置检查 ====================

log_info "Step 0: 前置检查..."

check_command "git"
check_command "ssh"
check_command "scp"
check_command "curl"
check_command "tar"

if [ ! -d "${LOCAL_BACKEND_DIR}" ]; then
    log_error "后端目录不存在: ${LOCAL_BACKEND_DIR}"; exit 1
fi
if [ ! -d "${LOCAL_FRONTEND_DIR}" ]; then
    log_error "前端目录不存在: ${LOCAL_FRONTEND_DIR}"; exit 1
fi
if ! command -v pnpm >/dev/null 2>&1; then
    log_error "pnpm 未安装"; exit 1
fi
if [ ! -d "${LOCAL_MIGRATION_DIR}" ]; then
    log_error "Flask-Migrate 未初始化: ${LOCAL_MIGRATION_DIR} 不存在"
    log_error "请先在本地执行: cd backend && flask db init && flask db migrate -m 'initial schema'"
    exit 1
fi
if [ ! -f "${LOCAL_SERVICE_FILE}" ]; then
    log_error "systemd 单元模板不存在: ${LOCAL_SERVICE_FILE}"; exit 1
fi

# 至少有一个迁移版本
MIGRATION_FILE_COUNT=0
for f in "${LOCAL_MIGRATION_DIR}"/versions/*.py; do
    [ -f "$f" ] || continue
    MIGRATION_FILE_COUNT=$((MIGRATION_FILE_COUNT + 1))
done
if [ "${MIGRATION_FILE_COUNT}" -eq 0 ]; then
    log_error "${LOCAL_MIGRATION_DIR}/versions 下无迁移文件，请先 flask db migrate"
    exit 1
fi

log_info "检查服务器连通性: ${SERVER_USER}@${SERVER_HOST}"
if ! ssh_exec "echo ok" >/dev/null 2>&1; then
    log_error "无法连接到服务器 ${SERVER_USER}@${SERVER_HOST}"; exit 1
fi
log_success "服务器连通性正常"

# ==================== Step 0.5: Git 切换 (可选) ====================

if [ -n "${DEPLOY_BRANCH}" ] || [ -n "${DEPLOY_TAG}" ]; then
    log_info "Step 0.5: Git 切换..."

    ORIGINAL_REF=$(git rev-parse --abbrev-ref HEAD 2>/dev/null)
    if [ "${ORIGINAL_REF}" = "HEAD" ]; then
        ORIGINAL_REF=$(git rev-parse HEAD)
    fi
    log_info "保存当前引用: ${ORIGINAL_REF}"

    if [ -n "${DEPLOY_BRANCH}" ]; then
        log_info "拉取远程分支: ${DEPLOY_BRANCH}"
        if ! git fetch origin "${DEPLOY_BRANCH}" 2>/dev/null; then
            abort_release "远程分支 ${DEPLOY_BRANCH} 不存在"
        fi
        if ! git checkout "${DEPLOY_BRANCH}" 2>/dev/null; then
            abort_release "无法切换到分支 ${DEPLOY_BRANCH}"
        fi
        if ! git pull origin "${DEPLOY_BRANCH}" 2>/dev/null; then
            log_warning "拉取远程更新失败，使用本地版本"
        fi
        log_success "已切换到分支: ${DEPLOY_BRANCH}"
    fi

    if [ -n "${DEPLOY_TAG}" ]; then
        log_info "拉取远程标签: ${DEPLOY_TAG}"
        git fetch --tags 2>/dev/null || true
        if ! git rev-parse "${DEPLOY_TAG}" >/dev/null 2>&1; then
            abort_release "标签 ${DEPLOY_TAG} 不存在"
        fi
        if ! git checkout "${DEPLOY_TAG}" 2>/dev/null; then
            abort_release "无法切换到标签 ${DEPLOY_TAG}"
        fi
        log_success "已切换到标签: ${DEPLOY_TAG} (detached HEAD)"
    fi

    DEPLOY_COMMIT=$(git rev-parse --short HEAD)
    DEPLOY_COMMIT_MSG=$(git log -1 --pretty=%s)
    log_info "部署版本: ${DEPLOY_COMMIT} - ${DEPLOY_COMMIT_MSG}"
fi

TIMESTAMP="$(date '+%Y%m%d%H%M%S')"

# ==================== Step 1: 构建前端 ====================

log_info "Step 1: 构建前端..."
cd "${LOCAL_FRONTEND_DIR}"

log_info "安装前端依赖..."
pnpm install --frozen-lockfile 2>/dev/null || pnpm install

log_info "临时修改 .env.prod 的 VITE_SERVICE_BASE_URL=/api (依赖 1Panel 反代)..."
ENV_PROD_FILE="${LOCAL_FRONTEND_DIR}/.env.prod"
if [ -f "${ENV_PROD_FILE}" ]; then
    cp "${ENV_PROD_FILE}" "${ENV_PROD_FILE}.release_bak"
    sed -i '' 's|^VITE_SERVICE_BASE_URL=.*|VITE_SERVICE_BASE_URL=/api|' "${ENV_PROD_FILE}"
    log_info "已临时覆盖 VITE_SERVICE_BASE_URL，构建后将自动还原"
else
    abort_release ".env.prod 不存在: ${ENV_PROD_FILE}"
fi

log_info "执行前端生产构建 (pnpm build)..."
set +e
pnpm build
BUILD_RC=$?
set -e
mv "${ENV_PROD_FILE}.release_bak" "${ENV_PROD_FILE}"
log_info ".env.prod 已还原"
if [ ${BUILD_RC} -ne 0 ]; then
    abort_release "前端构建失败"
fi

if [ ! -d "${LOCAL_FRONTEND_DIR}/dist" ]; then
    abort_release "前端构建产物不存在: ${LOCAL_FRONTEND_DIR}/dist"
fi

FRONTEND_ARCHIVE="work_tools_frontend_${TIMESTAMP}.tar.gz"
log_info "打包前端产物: ${FRONTEND_ARCHIVE}"
COPYFILE_DISABLE=1 tar -czf "${SCRIPT_DIR}/${FRONTEND_ARCHIVE}" -C "${LOCAL_FRONTEND_DIR}" dist
log_success "前端打包完成: ${FRONTEND_ARCHIVE}"

cd "${SCRIPT_DIR}"

# ==================== Step 2: 准备后端 ====================

log_info "Step 2: 准备后端..."

log_info "打包后端代码 (排除 __pycache__/venv)..."
BACKEND_ARCHIVE="work_tools_backend_${TIMESTAMP}.tar.gz"
COPYFILE_DISABLE=1 tar -czf "${SCRIPT_DIR}/${BACKEND_ARCHIVE}" \
    --exclude='__pycache__' \
    --exclude='*.pyc' \
    --exclude='.pytest_cache' \
    --exclude='venv' \
    --exclude='.env' \
    --exclude='.env.prod' \
    -C "${SCRIPT_DIR}" backend
log_success "后端打包完成: ${BACKEND_ARCHIVE}"

# ==================== Step 3: 数据库迁移检测 ====================

log_info "Step 3: 数据库迁移检测..."

# 解析本地迁移文件，存为 "rev|down|filename" 元组数组（兼容 bash 3.2）
MIGRATION_ENTRIES=()
for f in "${LOCAL_MIGRATION_DIR}"/versions/*.py; do
    [ -f "$f" ] || continue
    rev=$(grep -E "^revision = " "$f" | head -1 | grep -oE "'[^']+'|\"[^\"]+\"" | head -1 | tr -d "'\"")
    down=$(grep -E "^down_revision = " "$f" | head -1 | grep -oE "'[^']+'|\"[^\"]+\"|None" | head -1 | tr -d "'\"")
    [ -z "${rev}" ] && continue
    [ -z "${down}" ] && down="None"
    MIGRATION_ENTRIES+=("${rev}|${down}|${f}")
done

# 辅助：通过 revision 查 down_revision
get_down_by_rev() {
    local target="$1"
    for entry in "${MIGRATION_ENTRIES[@]}"; do
        if [ "$(echo "${entry}" | cut -d'|' -f1)" = "${target}" ]; then
            echo "${entry}" | cut -d'|' -f2
            return 0
        fi
    done
    echo ""
}

# 辅助：通过 revision 查文件路径
get_file_by_rev() {
    local target="$1"
    for entry in "${MIGRATION_ENTRIES[@]}"; do
        if [ "$(echo "${entry}" | cut -d'|' -f1)" = "${target}" ]; then
            echo "${entry}" | cut -d'|' -f3
            return 0
        fi
    done
    echo ""
}

log_info "本地迁移版本数量: ${#MIGRATION_ENTRIES[@]}"

# 获取远程已应用的迁移版本（通过 Docker MySQL 容器查询 alembic_version 表）
MYSQL_CONTAINER=$(ssh_exec "sudo docker ps --format '{{.Names}}' 2>/dev/null | grep -i mysql | head -1" 2>/dev/null) || MYSQL_CONTAINER=""

REMOTE_VERSIONS_RAW=""
if [ -n "${MYSQL_CONTAINER}" ]; then
    log_info "找到 MySQL 容器: ${MYSQL_CONTAINER}"

    # 从远程 .env.prod 读取数据库账号（若不存在说明首次部署）
    REMOTE_DB_USER=$(ssh_exec "sudo grep '^DB_USER=' ${REMOTE_ENV_FILE} 2>/dev/null | cut -d= -f2" 2>/dev/null | tr -d '[:space:]') || REMOTE_DB_USER=""
    REMOTE_DB_PWD=$(ssh_exec "sudo grep '^DB_PASSWORD=' ${REMOTE_ENV_FILE} 2>/dev/null | cut -d= -f2" 2>/dev/null | tr -d '[:space:]') || REMOTE_DB_PWD=""
    REMOTE_DB_NAME=$(ssh_exec "sudo grep '^DB_NAME=' ${REMOTE_ENV_FILE} 2>/dev/null | cut -d= -f2" 2>/dev/null | tr -d '[:space:]') || REMOTE_DB_NAME=""

    if [ -n "${REMOTE_DB_USER}" ] && [ -n "${REMOTE_DB_PWD}" ] && [ -n "${REMOTE_DB_NAME}" ]; then
        REMOTE_VERSIONS_RAW=$(ssh_exec \
            "sudo docker exec ${MYSQL_CONTAINER} mysql -u${REMOTE_DB_USER} -p${REMOTE_DB_PWD} ${REMOTE_DB_NAME} -N \
             -e 'SELECT version_num FROM alembic_version' 2>/dev/null" 2>/dev/null) || REMOTE_VERSIONS_RAW=""
    else
        log_warning "远程 .env.prod 不存在或缺失 DB 配置（可能是首次部署）"
    fi
else
    log_warning "未找到 MySQL 容器（可能是首次部署或 MySQL 未容器化）"
fi

# 解析远程已执行迁移链（从当前 head 沿 down_revision 反向遍历）
PENDING_REVISIONS=()
if [ -n "${REMOTE_VERSIONS_RAW}" ]; then
    CURRENT_HEAD=$(echo "${REMOTE_VERSIONS_RAW}" | grep -v 'Warning' | grep -v '^$' | head -1)
    log_info "远程当前迁移 head: ${CURRENT_HEAD}"

    # 沿 down_revision 反向收集已应用
    EXECUTED_SET=":${CURRENT_HEAD}:"
    CUR="${CURRENT_HEAD}"
    while [ -n "${CUR}" ] && [ "${CUR}" != "None" ] && [ "${CUR}" != "null" ]; do
        NEXT=$(get_down_by_rev "${CUR}")
        [ -z "${NEXT}" ] && break
        [ "${NEXT}" = "None" ] && break
        EXECUTED_SET="${EXECUTED_SET}${NEXT}:"
        CUR="${NEXT}"
    done
    log_info "已应用迁移集合: ${EXECUTED_SET}"

    # 求待执行
    for entry in "${MIGRATION_ENTRIES[@]}"; do
        rev=$(echo "${entry}" | cut -d'|' -f1)
        case "${EXECUTED_SET}" in
            *":${rev}:"*) ;;
            *) PENDING_REVISIONS+=("${rev}") ;;
        esac
    done
else
    log_warning "无法获取远程迁移状态（首次部署或库未就绪），全部本地迁移都将被视为待执行"
    for entry in "${MIGRATION_ENTRIES[@]}"; do
        PENDING_REVISIONS+=("$(echo "${entry}" | cut -d'|' -f1)")
    done
fi

if [ ${#PENDING_REVISIONS[@]} -eq 0 ]; then
    log_success "无待执行的数据库迁移"
else
    log_warning "========================================="
    log_warning "检测到 ${#PENDING_REVISIONS[@]} 个待执行迁移!"
    log_warning "========================================="

    for rev in "${PENDING_REVISIONS[@]}"; do
        f=$(get_file_by_rev "${rev}")
        filename=$(basename "${f}")
        echo ""
        log_warning "--- 迁移: ${filename} (revision=${rev}) ---"
        cat "${f}"
        echo ""
        log_warning "--- 结束 ---"
    done

    echo ""
    log_warning "以上迁移将在远程执行: venv/bin/flask db upgrade"
    echo -n "确认执行这些迁移？(yes/no): "
    read -r confirm
    if [ "$confirm" != "yes" ] && [ "$confirm" != "y" ]; then
        abort_release "用户取消部署"
    fi

    echo ""
    log_warning "========================================="
    log_warning "重要：执行迁移前必须备份数据库！"
    log_warning "========================================="
    echo ""
    echo "请在远程服务器手动执行数据库备份。"
    echo "登录服务器: ssh ${SERVER_USER}@${SERVER_HOST}"
    echo ""
    echo -n "已完成数据库备份？输入 BACKUP 确认: "
    read -r backup_confirm
    if [ "$backup_confirm" != "BACKUP" ]; then
        abort_release "未确认备份，部署中止"
    fi
    log_success "备份已确认"
fi

# ==================== Step 4: 上传前端 ====================

log_info "Step 4: 上传前端..."

log_info "上传前端压缩包..."
scp_to_tmp "${SCRIPT_DIR}/${FRONTEND_ARCHIVE}"

log_info "远程更新前端（站点根同址部署，保留 backend/venv/logs/.env.prod）..."
ssh_exec bash -s <<REMOTE_SCRIPT
set -e

# 站点根目录必须存在（1Panel 已创建）
if [ ! -d "${REMOTE_FRONTEND_DIR}" ]; then
    echo "[ERROR] 站点根目录不存在: ${REMOTE_FRONTEND_DIR}"
    echo "        请先在 1Panel 后台创建 work-tools 站点"
    exit 1
fi

# 备份旧前端的"静态文件部分"（排除 backend/venv/logs/.env.prod/_bak_*）
echo "[INFO] 备份旧前端静态文件到 ${REMOTE_FRONTEND_BAK}..."
sudo rm -rf "${REMOTE_FRONTEND_BAK}"
sudo mkdir -p "${REMOTE_FRONTEND_BAK}"
cd "${REMOTE_FRONTEND_DIR}"

# 普通文件（非隐藏）
for item in *; do
    [ -e "\$item" ] || continue
    case "\$item" in
        backend|venv|logs) continue ;;
        *) sudo mv -- "\$item" "${REMOTE_FRONTEND_BAK}/" 2>/dev/null || true ;;
    esac
done

# 隐藏文件（排除 .env.prod 和 . ..）
for item in .*; do
    [ -e "\$item" ] || continue
    case "\$item" in
        .|..|.env.prod) continue ;;
        *) sudo mv -- "\$item" "${REMOTE_FRONTEND_BAK}/" 2>/dev/null || true ;;
    esac
done

# 解压新前端到临时目录
cd /tmp
rm -rf /tmp/dist
tar -xzf "${FRONTEND_ARCHIVE}"

# 把 dist 内容（含隐藏文件）覆盖到站点根
sudo cp -r /tmp/dist/. "${REMOTE_FRONTEND_DIR}/"
sudo rm -rf /tmp/dist /tmp/${FRONTEND_ARCHIVE}

# 确保前端静态文件可读（文件 644、目录 755），nginx 以任意用户都能读取
# 不动 backend/venv/logs/.env.prod 的属主（保持 ubuntu 拥有，供 systemd 使用）
sudo find "${REMOTE_FRONTEND_DIR}" -mindepth 1 -maxdepth 1 \
    ! -name backend ! -name venv ! -name logs ! -name .env.prod ! -name _bak_frontend \
    -type d -exec sudo chmod 755 {} +
sudo find "${REMOTE_FRONTEND_DIR}" -mindepth 1 -maxdepth 1 \
    ! -name backend ! -name venv ! -name logs ! -name .env.prod ! -name _bak_frontend \
    -type f -exec sudo chmod 644 {} +

echo "[SUCCESS] 前端部署完成（后端目录已保留）"
REMOTE_SCRIPT

log_success "前端更新完成"

# ==================== Step 5: 上传后端 + 同步依赖 ====================

log_info "Step 5: 上传后端 + 同步依赖..."

log_info "上传后端代码包到 /tmp..."
scp_to_tmp "${SCRIPT_DIR}/${BACKEND_ARCHIVE}"

log_info "上传 systemd 单元文件..."
scp_to_tmp "${LOCAL_SERVICE_FILE}"

log_info "远程解压代码、同步依赖..."
ssh_exec bash -s <<REMOTE_SCRIPT
set -e

# 站点根必须存在（1Panel 创建）
if [ ! -d "${REMOTE_APP_ROOT}" ]; then
    echo "[ERROR] 站点根目录不存在: ${REMOTE_APP_ROOT}"
    echo "        请先在 1Panel 后台创建 work-tools 站点"
    exit 1
fi

# 创建 backend/venv/logs 子目录并让 ${SERVER_USER} 拥有（站点根本身归 root/1Panel）
sudo mkdir -p "${REMOTE_BACKEND_DIR}" "${REMOTE_VENV_DIR}" "${REMOTE_LOG_DIR}"
sudo chown -R ${SERVER_USER}:${SERVER_USER} "${REMOTE_BACKEND_DIR}" "${REMOTE_VENV_DIR}" "${REMOTE_LOG_DIR}"

# 首次部署：创建 venv
if [ ! -x "${REMOTE_VENV_DIR}/bin/python" ]; then
    echo "[INFO] 首次部署，创建 venv: ${REMOTE_VENV_DIR}"
    python3 -m venv "${REMOTE_VENV_DIR}"
fi

# 备份旧 backend（含 app.py 视为已部署过）
if [ -e "${REMOTE_BACKEND_DIR}/app.py" ]; then
    echo "[INFO] 备份旧后端代码到 ${REMOTE_BACKEND_DIR}_bak"
    sudo rm -rf "${REMOTE_BACKEND_DIR}_bak"
    sudo mv "${REMOTE_BACKEND_DIR}" "${REMOTE_BACKEND_DIR}_bak"
    sudo mkdir -p "${REMOTE_BACKEND_DIR}"
    sudo chown ${SERVER_USER}:${SERVER_USER} "${REMOTE_BACKEND_DIR}"
fi

# 解压新代码并复制到 backend 目录
cd /tmp
tar -xzf "${BACKEND_ARCHIVE}"
cp -r /tmp/backend/. "${REMOTE_BACKEND_DIR}/"
rm -rf /tmp/backend "/tmp/${BACKEND_ARCHIVE}"

# 安装/升级依赖
echo "[INFO] 同步 Python 依赖..."
"${REMOTE_VENV_DIR}/bin/pip" install --upgrade pip >/dev/null
"${REMOTE_VENV_DIR}/bin/pip" install -r "${REMOTE_BACKEND_DIR}/requirements.txt"

# 安装 systemd 单元
echo "[INFO] 安装 systemd 单元..."
sudo cp "/tmp/${SERVICE_NAME}.service" "${REMOTE_SERVICE_PATH}"
sudo chown root:root "${REMOTE_SERVICE_PATH}"
sudo chmod 644 "${REMOTE_SERVICE_PATH}"
sudo systemctl daemon-reload
sudo systemctl enable "${SERVICE_NAME}" >/dev/null 2>&1 || true
rm -f "/tmp/${SERVICE_NAME}.service"

# 确保 .env.prod 存在
if [ ! -f "${REMOTE_ENV_FILE}" ]; then
    echo "[WARN] ${REMOTE_ENV_FILE} 不存在（首次部署）"
    sudo tee "${REMOTE_ENV_FILE}" > /dev/null <<'ENVEOF'
${PROD_ENV_CONTENT}
ENVEOF
    sudo chown ${SERVER_USER}:${SERVER_USER} "${REMOTE_ENV_FILE}"
    sudo chmod 600 "${REMOTE_ENV_FILE}"
    echo "[ERROR] 已写入 .env.prod 模板（含 <CHANGE_ME> 占位符），请手动："
    echo "        1. 登录服务器: ssh ${SERVER_USER}@${SERVER_HOST}"
    echo "        2. 编辑: sudo vi ${REMOTE_ENV_FILE}"
    echo "        3. 替换所有 <CHANGE_ME> 为真实值"
    echo "        4. 建库授权（在 MySQL 容器内）:"
    echo "           CREATE DATABASE teacher_recruitment CHARACTER SET utf8mb4;"
    echo "           CREATE USER 'worktools'@'%' IDENTIFIED BY '<your_password>';"
    echo "           GRANT ALL ON teacher_recruitment.* TO 'worktools'@'%';"
    echo "           FLUSH PRIVILEGES;"
    echo "        5. 再次运行本脚本完成部署"
    exit 2
fi

# 校验 .env.prod 不含占位符
if grep -q '<CHANGE_ME>' "${REMOTE_ENV_FILE}"; then
    echo "[ERROR] ${REMOTE_ENV_FILE} 仍含 <CHANGE_ME> 占位符，请先替换"
    echo "        sudo vi ${REMOTE_ENV_FILE}"
    exit 2
fi

echo "[SUCCESS] 后端代码与依赖已同步"
REMOTE_SCRIPT

log_success "后端上传完成"

# ==================== Step 6: 远程应用迁移 + 重启服务 ====================

log_info "Step 6: 远程应用迁移 + 重启服务..."

ssh_exec bash -s <<REMOTE_SCRIPT
set -e
cd "${REMOTE_BACKEND_DIR}"
export FLASK_APP=app:create_app
echo "[INFO] 执行 flask db upgrade..."
"${REMOTE_VENV_DIR}/bin/flask" db upgrade

echo "[INFO] 重启 systemd 服务 ${SERVICE_NAME}..."
sudo systemctl restart "${SERVICE_NAME}"

sleep 2
sudo systemctl status "${SERVICE_NAME}" --no-pager | head -15 || true
REMOTE_SCRIPT

log_success "重启命令已执行"

# ==================== Step 7: 健康检查 ====================

log_info "Step 7: 健康检查..."

# 通过 1Panel 反代的 /api 路径访问 Flask 后端
LOGIN_URL="http://${SERVER_HOST}:${FRONTEND_PORT}/api/auth/login"
MAX_RETRY=30
RETRY_INTERVAL=5
HEALTHY=false

log_info "等待后端服务启动..."
for i in $(seq 1 $MAX_RETRY); do
    RESPONSE=$(curl -s -o /dev/null -w "%{http_code}" -X POST -H "Content-Type: application/json" \
        -d '{"userName":"admin","password":"123456"}' "${LOGIN_URL}" 2>/dev/null || echo "000")
    if [ "${RESPONSE}" = "200" ]; then
        HEALTHY=true
        break
    fi
    log_info "等待后端启动... (${i}/${MAX_RETRY}, HTTP=${RESPONSE})"
    sleep ${RETRY_INTERVAL}
done

if [ "${HEALTHY}" = true ]; then
    log_success "后端健康检查通过: ${LOGIN_URL}"
else
    log_error "后端启动超时！请检查："
    log_error "  ssh ${SERVER_USER}@${SERVER_HOST} 'sudo journalctl -u ${SERVICE_NAME} -n 100 --no-pager'"
    log_error "  ssh ${SERVER_USER}@${SERVER_HOST} 'sudo systemctl status ${SERVICE_NAME} --no-pager'"
    exit 1
fi

FRONTEND_URL="http://${SERVER_HOST}:${FRONTEND_PORT}"
FRONTEND_HTTP=$(curl -s -o /dev/null -w "%{http_code}" "${FRONTEND_URL}" 2>/dev/null || echo "000")
if [ "${FRONTEND_HTTP}" = "200" ]; then
    log_success "前端可访问: ${FRONTEND_URL}"
else
    log_warning "前端返回 HTTP ${FRONTEND_HTTP}: ${FRONTEND_URL}"
fi

# ==================== Step 8: 清理 ====================

log_info "Step 8: 清理..."
rm -f "${SCRIPT_DIR}/${FRONTEND_ARCHIVE}"
rm -f "${SCRIPT_DIR}/${BACKEND_ARCHIVE}"

# 恢复原始分支
if [ -n "${ORIGINAL_REF}" ]; then
    log_info "恢复原始引用: ${ORIGINAL_REF}"
    if git checkout "${ORIGINAL_REF}" 2>/dev/null; then
        log_success "已恢复到: ${ORIGINAL_REF}"
    else
        log_warning "无法恢复到 ${ORIGINAL_REF}，请手动切换"
    fi
    ORIGINAL_REF=""
fi

log_success "清理完成"

# ==================== 部署完成 ====================

echo ""
log_success "========================================="
log_success "  work-tools 部署完成！"
log_success "========================================="
if [ -n "${DEPLOY_BRANCH}" ]; then
    log_success "  来源: 分支 ${DEPLOY_BRANCH} (${DEPLOY_COMMIT:-unknown})"
elif [ -n "${DEPLOY_TAG}" ]; then
    log_success "  来源: 标签 ${DEPLOY_TAG} (${DEPLOY_COMMIT:-unknown})"
else
    log_success "  来源: 当前代码"
fi
log_success "  前端: ${FRONTEND_URL}"
log_success "  后端: ${LOGIN_URL}"
log_success "  时间: $(date '+%Y-%m-%d %H:%M:%S')"
log_success "========================================="
