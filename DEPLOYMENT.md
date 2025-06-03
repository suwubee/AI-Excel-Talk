# 🚀 AI Excel 智能分析工具 - 部署指南

## 📋 部署方式概览

本应用支持多种部署方式，您可以根据需要选择最适合的方案：

| 部署方式 | 难度 | 成本 | 推荐场景 |
|---------|------|------|---------|
| Streamlit Cloud | ⭐ | 免费 | 个人使用、快速演示 |
| Docker本地 | ⭐⭐ | 免费 | 本地开发、测试 |
| Railway | ⭐⭐ | 付费 | 小团队、快速部署 |
| Render | ⭐⭐ | 付费 | 生产环境 |
| Heroku | ⭐⭐⭐ | 付费 | 企业级应用 |
| 自建服务器 | ⭐⭐⭐⭐ | 可控 | 企业内部部署 |

## ☁️ Streamlit Cloud部署（推荐）

### 优势
- 🆓 完全免费
- 🚀 部署简单
- 🔄 自动更新
- 📊 内置监控

### 部署步骤

1. **Fork仓库**
   ```bash
   # 在GitHub上Fork这个仓库到您的账户
   https://github.com/your-username/AI-Excel-Talk
   ```

2. **访问Streamlit Cloud**
   - 打开 [share.streamlit.io](https://share.streamlit.io)
   - 使用GitHub账户登录

3. **创建新应用**
   - 点击 "New app"
   - 选择您fork的仓库
   - 分支：`main`
   - 主文件路径：`app_enhanced_multiuser.py`
   - 应用URL：自定义或使用默认

4. **环境变量配置（可选）**
   在 App 设置 → Secrets 中添加：
   ```toml
   OPENAI_API_KEY = "sk-your-openai-api-key"
   DEFAULT_MODEL = "deepseek-v3"
   ```

5. **部署！**
   - 点击 "Deploy!" 按钮
   - 等待应用启动（约2-3分钟）

### 访问应用
部署成功后，您将获得一个类似的URL：
```
https://ai-excel-tool-your-app-name.streamlit.app
```

## 🐳 Docker本地部署

### 方式一：使用Docker Compose（推荐）

1. **克隆仓库**
   ```bash
   git clone https://github.com/your-username/AI-Excel-Talk.git
   cd AI-Excel-Talk
   ```

2. **启动应用**
   ```bash
   # 构建并启动
   docker-compose up -d
   
   # 查看日志
   docker-compose logs -f
   ```

3. **访问应用**
   ```
   http://localhost:8501
   ```

4. **停止应用**
   ```bash
   docker-compose down
   ```

### 方式二：直接使用Docker

1. **构建镜像**
   ```bash
   docker build -t ai-excel-tool .
   ```

2. **运行容器**
   ```bash
   docker run -d \
     --name ai-excel-tool \
     -p 8501:8501 \
     -v $(pwd)/user_uploads:/app/user_uploads \
     ai-excel-tool
   ```

3. **查看日志**
   ```bash
   docker logs -f ai-excel-tool
   ```

## 🚂 Railway部署

### 特点
- 💳 按使用量付费
- 🚀 Git集成自动部署
- 📊 内置监控和日志

### 部署步骤

1. **访问Railway**
   - 打开 [railway.app](https://railway.app)
   - 使用GitHub登录

2. **新建项目**
   - 点击 "New Project"
   - 选择 "Deploy from GitHub repo"
   - 选择您的仓库

3. **配置环境变量**
   ```
   OPENAI_API_KEY=sk-your-openai-api-key
   PORT=8501
   ```

4. **自动部署**
   - Railway会自动检测Dockerfile
   - 自动构建和部署

## 🎨 Render部署

### 特点
- 🆓 有免费层
- 🔄 自动SSL证书
- 🌍 全球CDN

### 部署步骤

1. **访问Render**
   - 打开 [render.com](https://render.com)
   - 连接GitHub账户

2. **创建Web Service**
   - 点击 "New" → "Web Service"
   - 连接您的GitHub仓库

3. **配置设置**
   ```
   Name: ai-excel-tool
   Environment: Docker
   Region: 选择最近的区域
   Branch: main
   ```

4. **环境变量**
   ```
   OPENAI_API_KEY=sk-your-openai-api-key
   PORT=8501
   ```

## 🔴 Heroku部署

### 准备文件
确保项目包含以下文件：
- `Procfile`
- `runtime.txt` 
- `requirements.txt`

### 部署步骤

1. **安装Heroku CLI**
   ```bash
   # macOS
   brew tap heroku/brew && brew install heroku
   
   # Windows
   # 下载并安装 Heroku CLI
   ```

2. **登录Heroku**
   ```bash
   heroku login
   ```

3. **创建应用**
   ```bash
   heroku create ai-excel-tool-yourname
   ```

4. **设置环境变量**
   ```bash
   heroku config:set OPENAI_API_KEY=sk-your-openai-api-key
   ```

5. **部署**
   ```bash
   git push heroku main
   ```

6. **查看应用**
   ```bash
   heroku open
   ```

## 🖥️ 自建服务器部署

### 系统要求
- Ubuntu 20.04+ / CentOS 8+
- Python 3.8+
- 2GB+ RAM
- 10GB+ 存储空间

### 部署步骤

1. **服务器准备**
   ```bash
   # 更新系统
   sudo apt update && sudo apt upgrade -y
   
   # 安装依赖
   sudo apt install -y python3 python3-pip git nginx
   ```

2. **克隆项目**
   ```bash
   git clone https://github.com/your-username/AI-Excel-Talk.git
   cd AI-Excel-Talk
   ```

3. **安装Python依赖**
   ```bash
   pip3 install -r requirements.txt
   ```

4. **创建系统服务**
   ```bash
   sudo nano /etc/systemd/system/ai-excel-tool.service
   ```
   
   添加以下内容：
   ```ini
   [Unit]
   Description=AI Excel Tool
   After=network.target
   
   [Service]
   Type=simple
   User=www-data
   WorkingDirectory=/path/to/AI-Excel-Talk
   ExecStart=/usr/bin/python3 run_multiuser.py
   Restart=always
   
   [Install]
   WantedBy=multi-user.target
   ```

5. **启动服务**
   ```bash
   sudo systemctl daemon-reload
   sudo systemctl enable ai-excel-tool
   sudo systemctl start ai-excel-tool
   ```

6. **配置Nginx反向代理**
   ```bash
   sudo nano /etc/nginx/sites-available/ai-excel-tool
   ```
   
   添加配置：
   ```nginx
   server {
       listen 80;
       server_name your-domain.com;
       
       location / {
           proxy_pass http://localhost:8501;
           proxy_http_version 1.1;
           proxy_set_header Upgrade $http_upgrade;
           proxy_set_header Connection 'upgrade';
           proxy_set_header Host $host;
           proxy_set_header X-Real-IP $remote_addr;
           proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
           proxy_set_header X-Forwarded-Proto $scheme;
           proxy_cache_bypass $http_upgrade;
       }
   }
   ```

7. **启用配置**
   ```bash
   sudo ln -s /etc/nginx/sites-available/ai-excel-tool /etc/nginx/sites-enabled/
   sudo nginx -t
   sudo systemctl restart nginx
   ```

## 🔧 配置优化

### 性能优化

1. **Streamlit配置**
   编辑 `.streamlit/config.toml`：
   ```toml
   [server]
   maxUploadSize = 200
   enableXsrfProtection = false
   
   [browser]
   gatherUsageStats = false
   
   [runner]
   magicEnabled = true
   ```

2. **内存优化**
   ```python
   # 在 config_multiuser.py 中调整
   MAX_CONCURRENT_USERS = 50  # 根据服务器配置调整
   SESSION_TIMEOUT = 12       # 缩短会话超时时间
   ```

### 安全配置

1. **环境变量**
   ```bash
   # 不要在代码中硬编码API Key
   export OPENAI_API_KEY="sk-your-api-key"
   ```

2. **防火墙设置**
   ```bash
   # 只开放必要端口
   sudo ufw allow 22   # SSH
   sudo ufw allow 80   # HTTP
   sudo ufw allow 443  # HTTPS
   sudo ufw enable
   ```

## 📊 监控和维护

### 日志查看

1. **Streamlit Cloud**
   - 在应用管理页面查看实时日志

2. **Docker**
   ```bash
   docker logs -f ai-excel-tool
   ```

3. **系统服务**
   ```bash
   sudo journalctl -u ai-excel-tool -f
   ```

### 健康检查

访问健康检查端点：
```
http://your-domain.com/_stcore/health
```

### 备份策略

1. **用户数据备份**
   ```bash
   # 定期备份用户上传的文件
   tar -czf backup-$(date +%Y%m%d).tar.gz user_uploads/
   ```

2. **配置备份**
   ```bash
   # 备份应用配置
   cp -r .streamlit/ backup/
   ```

## 🚨 故障排除

### 常见问题

1. **应用启动失败**
   ```bash
   # 检查依赖是否完整安装
   pip install -r requirements.txt
   
   # 检查Python版本
   python --version  # 需要 >= 3.8
   ```

2. **端口冲突**
   ```bash
   # 查看端口占用
   lsof -i :8501
   
   # 更改端口
   streamlit run app_enhanced_multiuser.py --server.port 8502
   ```

3. **内存不足**
   ```bash
   # 检查内存使用
   free -h
   
   # 清理缓存
   sudo sync && sudo sysctl vm.drop_caches=3
   ```

4. **文件权限问题**
   ```bash
   # 修复权限
   chmod -R 755 user_uploads/
   chown -R www-data:www-data user_uploads/
   ```

### 日志分析

查看详细错误信息：
```bash
# 实时查看应用日志
tail -f app.log

# 搜索错误信息
grep -i "error" app.log

# 分析访问模式
grep "User activity" app.log | tail -20
```

## 📞 技术支持

如果您在部署过程中遇到问题：

1. **查看文档**
   - [用户指南](USER_GUIDE.md)
   - [开发者指南](DEVELOPER_GUIDE.md)

2. **GitHub Issues**
   - 提交问题到项目Issues页面

3. **社区支持**
   - 加入项目Discussion讨论

---

**🎉 恭喜！** 您已成功部署AI Excel智能分析工具。立即开始体验智能化的Excel数据分析吧！ 