# 公文格式助手

公文格式助手是一个本地运行的 Word 文稿格式整理工具。用户上传 `.docx`、`.doc`、`.rtf`、`.txt`、`.html`、`.odt` 文稿后，工具会按内置公文格式模板生成新的 `.docx` 文件。

处理流程会先清除原文稿已有段落和文字直接格式，仅保留文本内容，再按规则重新排版。全文段前、段后间距统一为 `0`。

## 功能概览

- 上传文稿并生成标准化 `.docx`
- 自动处理页面、页边距、标题、正文、主送单位、会议信息、附件、落款、页码
- 自动识别一级、二级、三级标题并做编号校验
- 支持 Word 自动编号转为真实标题编号
- 支持系统字体下拉选择
- 支持中文字号选择，例如 `二号`、`三号`
- 支持调整正文字号、标题字号、行距、字间距
- 默认字间距为加宽 `0.5pt`

## 本地运行

进入项目目录后运行：

```bash
python3 server.py
```

默认打开：

```text
http://127.0.0.1:8765
```

如需指定端口：

```bash
PORT=8778 python3 server.py
```

## 桌面封装

封装后的 macOS `.app` 和 Windows `.exe` 都会启动本地服务，并自动打开浏览器界面。

macOS 打包：

```bash
bash tools/packaging/build_macos.sh
```

生成文件：

```text
dist/公文格式助手.app
```

Windows 本地打包：

```powershell
powershell -ExecutionPolicy Bypass -File tools/packaging/build_windows.ps1
```

生成文件：

```text
dist\GongwenHelper.exe
dist\public\
```

Windows 运行时请保持 `GongwenHelper.exe` 和 `public` 文件夹在同一目录，避免网页资源 404。

## GitHub Actions 打包 Windows

项目已内置工作流：

```text
.github/workflows/build-windows.yml
```

推送到 GitHub 的 `main` 分支后会自动打包，也可以在 GitHub Actions 页面手动点击 `Run workflow`。

打包完成后下载 artifact：

```text
公文格式助手-win
```

artifact 中应包含：

```text
GongwenHelper.exe
public/
```

## 当前默认格式

- 页面：A4
- 页边距：上 `3.7cm`，下 `3cm`，左 `2.8cm`，右 `2.6cm`，页脚 `2.5cm`
- 主标题：二号 `方正公文小标宋`，居中，行距 `33pt`
- 正文：三号 `方正公文仿宋`，固定行距 `29.5pt`
- 西文及数字：`Times New Roman`
- 字间距：加宽 `0.5pt`
- 一级标题：三号 `方正公文黑体`
- 二级标题：三号 `方正公文楷体`
- 三级标题：三号 `方正公文仿宋` 加粗
- 页码：页脚居中，样式 `— 1 —`，四号 `Times New Roman`

## 可选格式项

网页界面可调整：

- 西文及数字字体
- 主标题字体
- 正文字体
- 一级标题字体
- 二级标题字体
- 主标题字号
- 正文字号
- 正文行距
- 标题行距
- 字间距加宽值

默认值即当前公文格式模板。

## 非 `.docx` 文件转换

`.docx` 可直接处理。

其他格式需要系统有转换工具：

- macOS：优先使用系统 `textutil`
- Windows：建议安装 LibreOffice，并确保 `soffice.exe` 在 `PATH` 中

如果转换工具不可用，请先将文稿另存为 `.docx` 后再上传。

## 目录说明

```text
.
├── formatter.py                 # Word 格式处理核心逻辑
├── server.py                    # 本地 Web 服务
├── desktop_app.py               # 桌面入口，封装版会打开本地网页
├── public/                      # 网页界面
├── tools/packaging/             # macOS / Windows 打包脚本
├── .github/workflows/           # GitHub Actions Windows 打包流程
├── requirements-packaging.txt   # 打包依赖
└── README.md
```

`storage/uploads/`、`storage/outputs/`、`build/`、`dist/`、`.venv-packaging/` 都是本地产物，不应提交到 GitHub。
