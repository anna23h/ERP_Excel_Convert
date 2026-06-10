# 打包成 Windows exe（给办公室员工双击使用）

目标：把脚本打包成一个 **单文件 exe**，员工**双击运行、弹窗选文件、点按钮出结果**，无需装 Python、无需命令行。

---

## 一、在一台 Windows 机器上打包（一次性）

### 前提
- 一台 Windows 电脑。
- 安装 Python 3：到 <https://www.python.org/downloads/> 下载安装，**安装时务必勾选「Add Python to PATH」**。

### 步骤
1. 把整个项目文件夹拷到这台 Windows 机器（`gui.py`、`build_excel.py`、`stage2.py`、`step4_merge.py`、`requirements.txt`、`build_exe.bat` 必须在一起）。**`raw_data/` 不用拷**。
2. 双击 `build_exe.bat`（或在命令提示符 `cd` 到项目目录后运行它）。
3. 等待完成，产物在 `dist\VOTool.exe`（可在资源管理器里重命名为中文，如 `VO拉单工具.exe`）。

> 手动打包命令（等价于 bat 内容）：
> ```bat
> py -3 -m pip install -r requirements.txt
> py -3 -m PyInstaller --onefile --windowed --name VOTool --collect-all openpyxl gui.py
> ```
> （若 `py` 不可用，把 `py -3` 换成 `python`）

### 体积与速度
exe 含 pandas，约 100–300 MB，**首次启动慢几秒属正常**（单文件需先解压到临时目录）。

---

## 二、发给员工使用

把 `VOTool.exe`（可重命名为中文）单独发给员工（U盘/共享盘/IM 均可）。员工**双击**即可，界面操作：

1. **① 输入文件**：选「ERP 导出」「天猫全量」，设「输出目录」。
2. **② 阶段一**：点「生成 拣货表+面单」→ 打印给仓库。
3. 仓库返回、在「无货勾选」页把无货行填 `1` 并另存。
4. **③ 阶段二**：选「返回文件」（账单模板选填）、填日期 → 点「生成 B/C/D/缺货记录」。
5. 点「打开输出文件夹」取结果。

---

## 三、故障排查

| 现象 | 处理 |
|---|---|
| **打包卡在 `Installing dependencies` 就失败** | 多半是 `python` 指向了**微软应用商店占位程序**或没进 PATH。解决：① 重装 Python 勾选「Add to PATH」；② 关掉商店别名：设置 → 应用 → 高级应用设置 → 应用执行别名 → 关闭 `python.exe`/`python3.exe`。新版 `build_exe.bat` 会自动优先用 `py` 启动器并给出提示 |
| `pyinstaller` 不是内部或外部命令 | 已改用 `python -m PyInstaller`，无需 `pyinstaller` 在 PATH |
| 双击 exe 闪退、看不到报错 | 重新打包时**去掉 `--windowed`**，会保留黑窗口显示错误信息 |
| 提示缺少 pandas 子模块 | 打包命令末尾加 `--collect-all pandas` 再打 |
| 杀毒软件误报 | PyInstaller 单文件 exe 常被误报，加信任白名单即可 |
| 打开 Excel 日期列显示 `######` | 列太窄，拉宽即可（脚本已自适应，正常不出现） |

---

## 四、数据合规

- exe 和真实订单数据**全程在本地电脑**，不联网、不上传，符合个人信息保护要求。
- 升级逻辑后需**重新打包**一次，再把新 exe 发给员工。
