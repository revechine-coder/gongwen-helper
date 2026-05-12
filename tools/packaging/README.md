# 公文格式助手打包说明

本项目使用 PyInstaller 分别生成 macOS 和 Windows 桌面版。两个系统需要分别在对应机器上打包，不能在 macOS 上直接生成 Windows `.exe`。

## macOS

在项目目录运行：

```bash
bash tools/packaging/build_macos.sh
```

生成文件：

```text
dist/公文格式助手.app
```

## Windows

把整个项目目录放到 Windows 电脑上，在 PowerShell 中进入项目目录后运行：

```powershell
powershell -ExecutionPolicy Bypass -File tools/packaging/build_windows.ps1
```

生成文件：

```text
dist\公文格式助手.exe
```

## GitHub Actions 云端打包 Windows

项目已内置工作流：

```text
.github/workflows/build-windows.yml
```

推送到 GitHub 的 `main` 分支后会自动打包，也可以在 GitHub Actions 页面手动点击 `Run workflow`。打包完成后，在 workflow 运行详情页下载 artifact：

```text
公文格式助手-win
```

里面就是 Windows `.exe`。

## 注意

- `.docx` 可直接处理。
- `.doc`、`.rtf`、`.txt`、`.html`、`.odt` 需要系统里有可用的转换工具。macOS 会优先使用系统 `textutil`；Windows 建议安装 LibreOffice，并确保 `soffice.exe` 在 PATH 中。
- 打包必须分别在对应系统上执行：macOS 生成 `.app`，Windows 生成 `.exe`。Windows 也可以直接使用 GitHub Actions 的 `windows-latest` runner 云端打包。
- `gongwen_helper.spec` 是备用 PyInstaller 配置，日常打包优先使用上面的两个脚本。
