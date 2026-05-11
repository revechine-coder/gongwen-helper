# GitHub 上传清单

## 建议提交的文件

```text
.github/workflows/build-windows.yml
.gitignore
GITHUB_UPLOAD_CHECKLIST.md
README.md
desktop_app.py
formatter.py
public/app.js
public/index.html
public/styles.css
requirements-packaging.txt
server.py
tools/packaging/README.md
tools/packaging/build_macos.sh
tools/packaging/build_windows.ps1
tools/packaging/gongwen_helper.spec
```

## 不建议提交的内容

```text
.DS_Store
.venv-packaging/
.pyinstaller-cache/
build/
dist/
storage/uploads/
storage/outputs/
*.zip
*.spec
gongwen-format-helper-source/
```

这些内容已在 `.gitignore` 中忽略。

## 首次上传步骤

```bash
git init
git add .
git commit -m "Initial public version"
git branch -M main
git remote add origin <你的 GitHub 仓库地址>
git push -u origin main
```

推送后，GitHub Actions 会自动运行 Windows 打包流程。

## Windows 打包产物

下载 artifact `公文格式助手-win` 后，请保持：

```text
公文格式助手.exe
public/
```

两者在同一目录，否则 Windows 打开网页可能出现 `404`。
