#!/bin/bash
###############################################################################
# work-tools 一键启动脚本（本地）
#
# 用法:
#   ./start.sh dev  [fe|be]    debug 模式（默认）：前端 vite dev(9530,HMR) + 后端 Flask debug(5001)
#   ./start.sh prod [fe|be]    release 模式：前端 build + vite preview(9725) + 后端 gunicorn(5001)
#   ./start.sh stop            停止所有前后端进程
#   ./start.sh restart [dev|prod] [fe|be]
#   ./start.sh status          查看运行状态
#   ./start.sh logs fe|be      跟踪日志（Ctrl+C 退出）
#   ./start.sh -h              帮助
#
# 说明:
#   - debug  (dev):  代码改动前后端均自动热重载，日常开发用
#   - release(prod): 前端跑构建产物、后端 gunicorn 单 worker（关 debug），上线前本地验证用
#   - 前端通过 .env.* 里的 VITE_SERVICE_BASE_URL 直连 http://localhost:5001/api，无需反代
#   - PID 存 .run/，日志存 logs/，按端口停止（不依赖 PID 文件准确性）
###############################################################################

set -uo pipefail

# ==================== 颜色输出（与 release.sh 一致）====================
RED='\033[0;31m'; GREEN='\033[0;32m'; YELLOW='\033[1;33m'; BLUE='\033[0;34m'; NC='\033[0m'
log_info()    { echo -e "${BLUE}[INFO]${NC} $(date '+%H:%M:%S') - $1"; }
log_success() { echo -e "${GREEN}[OK]${NC} $(date '+%H:%M:%S') - $1"; }
log_warning() { echo -e "${YELLOW}[WARN]${NC} $(date '+%H:%M:%S') - $1"; }
log_error()   { echo -e "${RED}[ERROR]${NC} $(date '+%H:%M:%S') - $1"; }

# ==================== 配置 ====================
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
WEB_DIR="$SCRIPT_DIR/web"
BACKEND_DIR="$SCRIPT_DIR/backend"
RUN_DIR="$SCRIPT_DIR/.run"
LOG_DIR="$SCRIPT_DIR/logs"
CONDA_ENV="teacherrecruitment"

BE_PORT=5001
FE_DEV_PORT=9530
FE_PROD_PORT=9725   # vite preview（vite.config.ts 已配置）

FE_PID_FILE="$RUN_DIR/fe.pid"
BE_PID_FILE="$RUN_DIR/be.pid"
FE_LOG="$LOG_DIR/frontend.log"
BE_LOG="$LOG_DIR/backend.log"

mkdir -p "$RUN_DIR" "$LOG_DIR"

# ==================== conda 激活 ====================
activate_conda() {
    local conda_sh=""
    for p in "$HOME/miniconda3/etc/profile.d/conda.sh" "$HOME/anaconda3/etc/profile.d/conda.sh" \
             "/opt/miniconda3/etc/profile.d/conda.sh" "/opt/anaconda3/etc/profile.d/conda.sh"; do
        [ -f "$p" ] && conda_sh="$p" && break
    done
    if [ -z "$conda_sh" ]; then
        log_error "找不到 conda.sh（尝试过 ~/miniconda3、~/anaconda3、/opt/...）"
        exit 1
    fi
    # shellcheck disable=SC1090
    source "$conda_sh"
    if ! conda activate "$CONDA_ENV" 2>/dev/null; then
        log_error "conda 环境 '$CONDA_ENV' 激活失败，请先创建：conda create -n $CONDA_ENV python=3.11"
        exit 1
    fi
}

# ==================== 工具函数 ====================
port_pids()    { lsof -ti:"$1" 2>/dev/null; }
port_in_use()  { [ -n "$(port_pids "$1")" ]; }

wait_port() {  # <port> <timeout_sec>
    local port=$1 timeout=$2 i=0
    while [ "$i" -lt "$timeout" ]; do
        port_in_use "$port" && return 0
        sleep 1; i=$((i + 1))
    done
    return 1
}

stop_port() {  # <port> <name>
    local port=$1 name=$2 pids
    pids="$(port_pids "$port")"
    [ -z "$pids" ] && return 0
    log_info "停止 $name（port $port, pid $(echo $pids | tr '\n' ' ')）"
    kill $pids 2>/dev/null || true
    local i=0
    while [ "$i" -lt 8 ]; do
        [ -z "$(port_pids "$port")" ] && return 0
        sleep 1; i=$((i + 1))
    done
    pids="$(port_pids "$port")"
    if [ -n "$pids" ]; then
        log_warning "$name 未响应 SIGTERM，强制 kill -9"
        kill -9 $pids 2>/dev/null || true
    fi
}

# ==================== 启动后端 ====================
start_backend() {
    local mode=$1
    if port_in_use "$BE_PORT"; then
        log_warning "后端端口 $BE_PORT 已被占用，跳过（如需重启先 ./start.sh stop）"
        return 0
    fi
    log_info "启动后端 [$mode] → http://127.0.0.1:$BE_PORT"
    : > "$BE_LOG"
    cd "$BACKEND_DIR"
    if [ "$mode" = "prod" ]; then
        # gunicorn 不走 app.py 的 __main__，自然关闭 debug/reloader；单 worker 避免后台清理线程重复
        nohup gunicorn --chdir "$BACKEND_DIR" -w 1 \
            -b "127.0.0.1:$BE_PORT" \
            --access-logfile - --error-logfile - \
            "app:create_app()" >> "$BE_LOG" 2>&1 &
    else
        nohup python -u app.py >> "$BE_LOG" 2>&1 &
    fi
    echo $! > "$BE_PID_FILE"
    cd "$SCRIPT_DIR"
    if wait_port "$BE_PORT" 30; then
        log_success "后端已就绪（pid $(cat "$BE_PID_FILE")）"
    else
        log_error "后端 30s 内未监听 $BE_PORT，日志片段："
        tail -n 15 "$BE_LOG" 2>/dev/null
    fi
}

# ==================== 启动前端 ====================
start_frontend() {
    local mode=$1 port cmd
    if [ "$mode" = "prod" ]; then
        port=$FE_PROD_PORT; cmd="preview"
        log_info "构建前端（pnpm build）…"
        : > "$FE_LOG"
        (cd "$WEB_DIR" && pnpm build) >> "$FE_LOG" 2>&1 || {
            log_error "前端构建失败，日志片段："; tail -n 20 "$FE_LOG"; return 1; }
        log_info "启动前端 [prod] → http://localhost:$port（vite preview）"
    else
        port=$FE_DEV_PORT; cmd="dev"
        log_info "启动前端 [dev] → http://localhost:$port（vite dev, HMR）"
        : > "$FE_LOG"
    fi
    if port_in_use "$port"; then
        log_warning "前端端口 $port 已被占用，跳过"
        return 0
    fi
    cd "$WEB_DIR"
    if [ "$cmd" = "preview" ]; then
        nohup pnpm preview --port "$port" --host >> "$FE_LOG" 2>&1 &
    else
        nohup pnpm dev >> "$FE_LOG" 2>&1 &
    fi
    echo $! > "$FE_PID_FILE"
    cd "$SCRIPT_DIR"
    # vite dev 首次启动需做依赖预构建，给足时间
    if wait_port "$port" 60; then
        log_success "前端已就绪 → http://localhost:$port"
    else
        log_error "前端 60s 内未监听 $port，日志片段："
        tail -n 15 "$FE_LOG" 2>/dev/null
    fi
}

# ==================== 停止 / 状态 / 日志 ====================
stop_all() {
    stop_port "$BE_PORT"      "后端"
    stop_port "$FE_DEV_PORT"  "前端(dev)"
    stop_port "$FE_PROD_PORT" "前端(prod)"
    rm -f "$FE_PID_FILE" "$BE_PID_FILE" 2>/dev/null || true
    log_success "已停止全部服务"
}

show_status() {
    local be_pids fe_pids fe_port be_disp fe_disp
    be_pids="$(port_pids "$BE_PORT")"
    # dev / prod 前端只可能其一在跑
    if port_in_use "$FE_DEV_PORT"; then fe_port="$FE_DEV_PORT"; fe_pids="$(port_pids "$FE_DEV_PORT")"
    elif port_in_use "$FE_PROD_PORT"; then fe_port="$FE_PROD_PORT"; fe_pids="$(port_pids "$FE_PROD_PORT")"
    else fe_port="-"; fe_pids=""; fi

    # debug reloader / vite 会派生多进程，折叠成单行显示
    be_disp="$(echo $be_pids | tr '\n' ' ' | tr -s ' ')"; [ -z "$be_disp" ] && be_disp="-"
    fe_disp="$(echo $fe_pids | tr '\n' ' ' | tr -s ' ')"; [ -z "$fe_disp" ] && fe_disp="-"

    printf "\n%-12s %8s   %-6s %s\n" "服务" "端口" "状态" "PID"
    printf "%-12s %8s   %b%-4s%b   %s\n" "后端" "$BE_PORT" \
        "$([ -n "$be_pids" ] && echo -e "$GREEN" || echo -e "$RED")" \
        "$([ -n "$be_pids" ] && echo "运行" || echo "停止")" "$NC" "$be_disp"
    printf "%-12s %8s   %b%-4s%b   %s\n" "前端" "$fe_port" \
        "$([ -n "$fe_pids" ] && echo -e "$GREEN" || echo -e "$RED")" \
        "$([ -n "$fe_pids" ] && echo "运行" || echo "停止")" "$NC" "$fe_disp"
    echo
    if [ -n "$be_pids$fe_pids" ]; then
        log_info "前端访问：http://localhost:${fe_port}（后端 API：http://localhost:$BE_PORT/api）"
    else
        log_info "无服务运行，用 ./start.sh dev 启动"
    fi
}

tail_logs() {
    local which=$1 f
    case "$which" in
        fe|frontend) f="$FE_LOG" ;;
        be|backend)  f="$BE_LOG" ;;
        *) log_error "用法: ./start.sh logs fe|be"; return 1 ;;
    esac
    [ -f "$f" ] || { log_warning "日志不存在: $f"; return; }
    log_info "跟踪 $f（Ctrl+C 退出）"
    tail -n 50 -f "$f"
}

# ==================== 帮助 ====================
usage() {
    cat <<'EOF'
work-tools 一键启动脚本

用法:
  ./start.sh dev  [fe|be]      debug 模式：前端 vite dev(9530) + 后端 Flask debug(5001)
  ./start.sh prod [fe|be]      release 模式：前端 build + vite preview(9725) + 后端 gunicorn(5001)
  ./start.sh stop              停止所有前后端进程
  ./start.sh restart [dev|prod] [fe|be]
  ./start.sh status            查看运行状态
  ./start.sh logs fe|be        跟踪日志（Ctrl+C 退出）

可选目标（默认全部）:
  fe   只启动前端
  be   只启动后端

示例:
  ./start.sh dev               # 一键起前后端（开发模式）
  ./start.sh prod              # 构建前端 + gunicorn 跑后端（本地生产模拟）
  ./start.sh dev be            # 只起后端
  ./start.sh stop && ./start.sh prod
EOF
}

# ==================== 参数解析 / 主逻辑 ====================
CMD="${1:-}"
[ -n "$CMD" ] && shift || true

case "$CMD" in
    dev|prod)
        activate_conda
        TARGET="${1:-all}"
        [ "$TARGET" = "proj" ] && TARGET="all"
        case "$TARGET" in
            all) start_backend "$CMD"; start_frontend "$CMD" ;;
            fe)  start_frontend "$CMD" ;;
            be)  start_backend "$CMD" ;;
            *)   log_error "未知目标: $TARGET（应为 all / fe / be）"; exit 1 ;;
        esac
        show_status
        ;;
    stop)
        stop_all
        ;;
    restart)
        MODE="${1:-dev}"
        TARGET="${2:-all}"
        log_info "重启（$MODE $TARGET）"
        stop_all
        sleep 1
        exec "$0" "$MODE" "$TARGET"
        ;;
    status)
        show_status
        ;;
    logs)
        WHICH="${1:-}"
        [ -z "$WHICH" ] && { log_error "用法: ./start.sh logs fe|be"; exit 1; }
        tail_logs "$WHICH"
        ;;
    -h|--help|"")
        usage
        ;;
    *)
        log_error "未知命令: $CMD"
        usage
        exit 1
        ;;
esac
