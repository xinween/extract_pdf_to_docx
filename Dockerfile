# 基于 CentOS 7（和你的 Alibaba Cloud Linux 完全同源）
FROM centos:7

# 安装 Python 3.10 + 依赖（用 yum，完全适配你的系统）
RUN yum install -y https://dl.fedoraproject.org/pub/epel/epel-release-latest-7.noarch.rpm \
    && yum install -y python310 python310-pip python310-devel \
    && yum install -y poppler-utils libglvnd-glx libglib2.0 \
    && yum clean all

# 使用 python3.10
RUN ln -s /usr/bin/python3.10 /usr/bin/python3 && ln -s /usr/bin/pip3.10 /usr/bin/pip

WORKDIR /app

# 阿里云 PyPI 加速
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt -i https://mirrors.aliyun.com/pypi/simple/

# 复制代码
COPY . .

# 目录
RUN mkdir -p uploads output

EXPOSE 5000

CMD ["python3", "app.py"]