# 数据处理系统

这是一个基于Streamlit的数据处理系统，支持多种数据处理功能。

## Docker构建和运行

### 构建Docker镜像

```bash
docker build -t streamlit-app .
```

### 后台运行容器

```bash
docker run -d -p 8501:8501 streamlit-app
```

### 查看运行状态

```bash
docker ps
```

### 停止容器

```bash
docker stop streamlit-app
```

### 启动已停止的容器

```bash
docker start streamlit-app
```

### 查看容器日志

```bash
docker logs streamlit-app
```

## 直接运行（非Docker）

### 安装依赖

```bash
pip install -r requirements.txt -i https://mirrors.aliyun.com/pypi/simple/
```

### 运行应用

```bash
streamlit run app.py
```