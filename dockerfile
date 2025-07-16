FROM node:18-bullseye-slim

# システムパッケージの更新とインストール（fileとdos2unixを追加）
RUN apt-get update && apt-get install -y \
    git \
    curl \
    bash \
    vim \
    nano \
    python3 \
    python3-pip \
    python3-dev \
    python3-venv \
    make \
    g++ \
    ca-certificates \
    build-essential \
    pkg-config \
    file \
    dos2unix \
    unzip \
    && rm -rf /var/lib/apt/lists/*

# pipのアップグレードとシンボリックリンク作成
RUN python3 -m pip install --upgrade pip && \
    if [ ! -f /usr/bin/python ]; then ln -s /usr/bin/python3 /usr/bin/python; fi && \
    if [ ! -f /usr/bin/pip ]; then ln -s /usr/bin/pip3 /usr/bin/pip; fi

# Node.js 20 LTSをインストール
RUN curl -fsSL https://deb.nodesource.com/setup_20.x | bash - && \
    apt-get install -y nodejs

#bunのインストール
RUN curl -fsSL https://bun.sh/install | bash
ENV BUN_INSTALL="/root/.bun"
ENV PATH="$BUN_INSTALL/bin:$PATH"


# Pythonパッケージのインストール
RUN pip install --no-cache-dir \
    google-generativeai \
    google-auth \
    google-auth-oauthlib \
    google-auth-httplib2 \
    google-cloud-aiplatform \
    requests \
    aiohttp \
    pandas \
    openpyxl \
    python-dotenv

# Node.jsパッケージのインストール
ENV CLAUDE_AUTO_UPDATE=false

RUN npm install -g \
    @anthropic-ai/claude-code@latest \
    azure-functions-core-tools@4 \
    @google/gemini-cli

# 作業ディレクトリの設定
WORKDIR /src


CMD ["/bin/bash"]