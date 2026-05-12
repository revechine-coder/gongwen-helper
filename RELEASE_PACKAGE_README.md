# 公文格式助手发布包说明

这个发布包用于交付当前版本。

## 包内内容

```text
macos/公文格式助手.app
source/gongwen-format-helper-github-upload.zip
README.md
```

## macOS 使用

进入 `macos` 文件夹，双击：

```text
公文格式助手.app
```

软件会启动本地服务并打开浏览器界面。

## Windows 使用

当前发布包不直接包含 Windows `.exe`。

Windows 版本请使用 `source/gongwen-format-helper-github-upload.zip` 上传到 GitHub，触发 GitHub Actions 云端打包。打包完成后下载 artifact：

```text
公文格式助手-win
```

下载后保持：

```text
公文格式助手.exe
public/
```

两者在同一目录，否则 Windows 打开网页可能出现 `404`。

详细上传和打包命令见源码包内：

```text
VPS_GITHUB_UPLOAD_COMMANDS.md
GITHUB_UPLOAD_CHECKLIST.md
```
