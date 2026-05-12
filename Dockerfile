FROM python:3.11-slim

# 安装 Pandoc 和 LibreOffice（LibreOffice 用于 PDF 转换，可选）
RUN apt-get update && \
    apt-get install -y pandoc libreoffice-core libreoffice-writer default-jre && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY . .

RUN pip install --no-cache-dir -r requirements.txt

# 使用 gunicorn 启动，端口通过 $PORT 环境变量动态获取
CMD sh -c "gunicorn app:app --bind 0.0.0.0:${PORT:-10000}"