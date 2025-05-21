# 1. 选择一个官方 Node.js 运行时作为基础镜像
# node:18-alpine 或 node:20-alpine 是一个不错的选择，因为它体积较小
FROM node:18-alpine

# 2. 在容器内设置工作目录
WORKDIR /usr/src/app

# 3. 复制 package.json 和 package-lock.json (如果存在)
# 将这些文件单独复制并先安装依赖，可以更好地利用 Docker 的构建缓存
COPY package*.json ./

# 4. 安装生产环境依赖
# 如果有 package-lock.json，npm ci 是首选，它可以提供更可靠和可复现的构建
# --only=production 确保只安装生产依赖
RUN npm ci --only=production || npm install --only=production

# 5. 复制应用程序的其余代码到工作目录
# 这应该在 npm install 之后，这样代码的改动不会让 npm install 缓存失效 (除非 package*.json 变动)
COPY . .
# 您的 server.js 代码应该包含创建 'uploads' 和 'filled_documents' 目录的逻辑
# (例如使用 fs.mkdirSync(..., { recursive: true }))

# 6. 创建一个非 root 用户和组，并更改工作目录的所有权
# 在复制完所有文件后执行此操作，然后切换用户
RUN addgroup -S appgroup && adduser -S appuser -G appgroup
RUN chown -R appuser:appgroup /usr/src/app

# 7. 暴露应用程序运行的端口 (与 server.js 中监听的端口一致)
EXPOSE 3000

# 8. 切换到非 root 用户运行应用，增强安全性
USER appuser

# 9. 定义容器启动时运行的命令
CMD ["node", "server.js"]