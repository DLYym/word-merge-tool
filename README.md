# Word 批量替换合并工具

把一个 Word 模板按「替换词列表」批量复制并替换关键词，最后合并成一个 Word 文档（每份之间用分页符分隔）。

## 功能

- 选择模板 `.docx`
- 指定「待替换的词」（如：朋友）
- 输入「替换后的词」列表（每行一个，或从 `.txt`/`.csv` 导入）
- 一键生成合并后的 Word 文档

## 直接运行（需要 Python + Word）

在 **Windows** 电脑上：

```bat
pip install pywin32
python word_merge_tool.py
```

> 依赖 `win32com`（pywin32），只能在 Windows 上运行；且电脑需已安装 Microsoft Word。

## 打包成独立 exe

PyInstaller **不能跨平台打包**，Windows 的 exe 必须在 Windows 环境下生成。任选下面一种方式。

### 方式一：在 Windows 电脑上打包（推荐，最简单）

1. 把整个 `word-merge-tool` 文件夹拷到 Windows 电脑。
2. 双击运行 `build.bat`（或手动执行下面的命令）。

```bat
pip install pywin32 pyinstaller
pyinstaller --noconsole --onefile --name WordMergeTool --icon app.ico word_merge_tool.py
```

3. 打包完成后，exe 在 `dist\WordMergeTool.exe`，可拷贝到任意 Windows 电脑直接双击使用（目标电脑仍需装有 Word，无需装 Python）。

### 方式二：云端打包（本机是 macOS/Linux 也能出 exe）

用 GitHub Actions 在 Windows 云环境里自动打包，不用本地装 Windows：

1. 把本文件夹推送到一个 GitHub 仓库。
2. 打开仓库的 **Actions** 标签页 → 选择 **Build Windows EXE** 工作流 → **Run workflow**。
3. 跑完后，在该次运行的 **Artifacts** 里下载 `WordMergeTool.exe`。

工作流文件已包含在 `.github/workflows/build.yml`。

## 文件说明

| 文件 | 说明 |
| --- | --- |
| `word_merge_tool.py` | 主程序（界面 + 合并逻辑） |
| `build.bat` | Windows 一键打包脚本 |
| `.github/workflows/build.yml` | GitHub Actions 云端打包工作流 |
| `app.ico` | 程序图标 |
| `gen_icon.py` | 图标生成脚本（一般用不到，可删） |

## 常见问题

- **运行报错缺 pywin32**：`pip install pywin32`。
- **打包后双击没反应**：先用「带控制台」的方式打包调试（去掉 `--noconsole`），看报错信息。
- **目标电脑没有 Word**：exe 只是免装 Python，仍需要本机装有 Microsoft Word 才能操作文档。
