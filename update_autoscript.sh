#!/bin/bash

# Docker项目自动更新脚本 - 针对AutoScript项目
# 使用方法: ./update_autoscript.sh

# 设置颜色输出
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# 项目配置
PROJECT_NAME="AutoScript"
PROJECT_DIR="./AutoScript"
GIT_REPO="https://github.com/Meteor-Comet/AutoScript.git"
IMAGE_NAME="streamlit-app"
CONTAINER_NAME="streamlit-app-container"
PORT="8501"

# 日志函数
log_info() {
    echo -e "${BLUE}[INFO]${NC} $1"
}

log_success() {
    echo -e "${GREEN}[SUCCESS]${NC} $1"
}

log_warning() {
    echo -e "${YELLOW}[WARNING]${NC} $1"
}

log_error() {
    echo -e "${RED}[ERROR]${NC} $1"
}

# 显示配置信息
echo "========================================"
log_info "AutoScript 自动部署脚本"
echo "========================================"
log_info "项目名称: $PROJECT_NAME"
log_info "项目目录: $PROJECT_DIR"
log_info "Git仓库: $GIT_REPO"
log_info "镜像名称: $IMAGE_NAME"
log_info "容器名称: $CONTAINER_NAME"
log_info "端口映射: $PORT:8501"
echo "========================================"
echo ""

# 函数：检查命令执行结果
check_result() {
    if [ $? -eq 0 ]; then
        log_success "$1"
    else
        log_warning "$2"
    fi
}

# 1. 停止并删除现有容器
log_info "步骤1: 停止并删除现有容器..."
if docker ps -a | grep -q "$CONTAINER_NAME"; then
    docker stop "$CONTAINER_NAME"
    check_result "容器 $CONTAINER_NAME 已停止" "停止容器失败"

    docker rm "$CONTAINER_NAME"
    check_result "容器 $CONTAINER_NAME 已删除" "删除容器失败"
else
    log_warning "容器 $CONTAINER_NAME 不存在，跳过此步骤"
fi

# 2. 删除旧镜像
log_info "步骤2: 删除旧镜像..."
if docker images | grep -q "$IMAGE_NAME"; then
    docker rmi "$IMAGE_NAME"
    check_result "旧镜像 $IMAGE_NAME 已删除" "删除镜像失败"
else
    log_warning "镜像 $IMAGE_NAME 不存在，跳过此步骤"
fi

# 3. 删除旧项目目录
log_info "步骤3: 清理旧项目文件..."
if [ -d "$PROJECT_DIR" ]; then
    rm -rf "$PROJECT_DIR"
    check_result "旧目录 $PROJECT_DIR 已删除" "删除目录失败"
else
    log_warning "目录 $PROJECT_DIR 不存在，跳过此步骤"
fi

# 4. 拉取最新代码
log_info "步骤4: 从GitHub拉取最新代码..."
git clone "$GIT_REPO" "$PROJECT_DIR"
if [ $? -eq 0 ]; then
    log_success "代码拉取成功"
else
    log_error "代码拉取失败，请检查网络连接和仓库URL"
    exit 1
fi

# 5. 进入项目目录并构建镜像
log_info "步骤5: 构建Docker镜像..."
cd "$PROJECT_DIR"
if [ $? -eq 0 ]; then
    log_success "进入项目目录: $(pwd)"

    # 检查Dockerfile是否存在
    if [ -f "Dockerfile" ]; then
        log_info "找到Dockerfile，开始构建镜像..."
        docker build -t "$IMAGE_NAME" .
        if [ $? -eq 0 ]; then
            log_success "Docker镜像构建成功: $IMAGE_NAME"
        else
            log_error "Docker镜像构建失败"
            exit 1
        fi
    else
        log_error "在项目目录中未找到Dockerfile"
        exit 1
    fi
else
    log_error "无法进入项目目录"
    exit 1
fi

# 6. 运行容器
log_info "步骤6: 启动Docker容器..."
docker run -d -p "$PORT":8501 --name "$CONTAINER_NAME" "$IMAGE_NAME"
if [ $? -eq 0 ]; then
    log_success "容器启动成功"
else
    log_error "容器启动失败"
    exit 1
fi

# 7. 显示部署结果
log_info "步骤7: 检查部署状态..."
sleep 3
echo ""
echo "========================================"
log_success "部署完成!"
echo "========================================"
log_info "容器状态:"
docker ps -f "name=$CONTAINER_NAME"

echo ""
log_info "镜像信息:"
docker images | grep "$IMAGE_NAME"

echo ""
log_info "容器日志 (最近10行):"
docker logs --tail 10 "$CONTAINER_NAME"

echo ""
log_info "访问地址: http://localhost:$PORT"
log_info "或者: http://$(hostname -I | awk '{print $1}'):$PORT"

echo ""
log_info "常用管理命令:"
echo "  查看日志: docker logs -f $CONTAINER_NAME"
echo "  进入容器: docker exec -it $CONTAINER_NAME /bin/bash"
echo "  停止容器: docker stop $CONTAINER_NAME"
echo "  重启容器: docker restart $CONTAINER_NAME"
echo "  删除容器: docker rm $CONTAINER_NAME"
echo "  删除镜像: docker rmi $IMAGE_NAME"