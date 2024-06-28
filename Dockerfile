# 使用 Python 3.11.7 作为基础镜像
FROM python:3.11.7-slim

# 设置工作目录
WORKDIR /app

# 复制项目的依赖文件到工作目录
COPY requirements.txt .
COPY templates .
COPY medication_ajust.py .

# 安装项目的依赖
RUN pip install -r requirements.txt

# 复制项目的源代码到工作目录
COPY . .

# 设置环境变量以指定 Flask 的运行主机和端口
ENV FLASK_APP=medication_ajust.py
ENV FLASK_RUN_HOST=0.0.0.0
ENV FLASK_RUN_PORT=9000

# 暴露 Flask 服务器的端口
EXPOSE 9000

# 运行 Flask 应用
CMD ["flask", "run"]
