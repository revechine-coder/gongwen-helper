# VPS / openclaw 上传 GitHub 命令

以下命令假设你已经把压缩包上传到 VPS，并解压到当前目录。

## 1. 解压项目

```bash
unzip gongwen-format-helper-github-upload.zip -d gongwen-format-helper
cd gongwen-format-helper
```

如果 VPS 上没有 `unzip`：

```bash
sudo apt update
sudo apt install -y unzip
```

## 2. 初始化 Git 仓库

```bash
git init
git add .
git commit -m "Initial public version"
git branch -M main
```

如首次使用 Git，需要先配置提交身份：

```bash
git config --global user.name "你的名字"
git config --global user.email "你的邮箱"
```

## 3. 绑定 GitHub 仓库并推送

把下面的 `<你的 GitHub 仓库地址>` 替换成真实地址，例如：

```text
https://github.com/用户名/仓库名.git
```

执行：

```bash
git remote add origin <你的 GitHub 仓库地址>
git push -u origin main
```

## 4. 触发 Windows 云端打包

推送到 GitHub 后，进入仓库页面：

```text
Actions -> Build Windows
```

如果没有自动运行，点击：

```text
Run workflow
```

运行完成后下载 artifact：

```text
公文格式助手-win
```

下载后请保持目录结构：

```text
公文格式助手.exe
public/
```

`public/` 必须和 `.exe` 在同一目录，否则 Windows 打开页面可能出现 `404`。

## 5. VPS 上本地测试网页

如果只想在 VPS 上测试 Web 服务：

```bash
python3 server.py
```

指定端口：

```bash
PORT=8778 python3 server.py
```

如果服务器需要外部访问，需要额外配置反向代理或开放端口。本项目默认只监听本机 `127.0.0.1`，适合本地桌面工具使用。
