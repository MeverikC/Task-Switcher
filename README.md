<div align="center">

# Task Switcher

<img src="./img/preview.png" alt="Task Switcher Preview" width="600" style="border-radius: 12px; box-shadow: 0 10px 30px rgba(0,0,0,0.2);" />

<br><br>

**一款基于 Python & PyQt6 构建的 Windows 现代化窗口切换工具。** *轻量、美观、高度可定制，专为替代原生 Alt+Tab 而生。*

</div>

---

## ✨ 特性 (Features)

- **🎨 高度可定制 UI**：
  - 支持 **列表模式 (List)** 与 **Win10 网格模式 (Grid)** 切换。
  - 自定义背景色、文字颜色、高亮色及不透明度。
  - **真·圆角窗口**：无黑底、无边框的现代化圆角设计。
  
- **🚀 智能排版**：
  - **自适应布局**：无滚动条设计，窗口高度随应用数量自动撑开，拒绝繁琐滚动。
  - 智能计算列宽与行高，完美利用屏幕空间。

- **⚡ 强力核心**：
  - **核弹级切换算法**：通过模拟输入 + 线程附着 (AttachThreadInput) + 底层 API 组合拳，彻底解决 Windows "拒绝访问" 和无法抢占焦点的问题。
  - 支持最小化窗口自动还原。
  - 系统托盘常驻，资源占用极低。

## 🛠️ 安装与运行 (Installation)

### 1. 环境准备
确保已安装 Python 3.9+。

```bash
# 克隆仓库
git clone https://github.com/MeverikC/Task-Switcher.git
cd task-switcher

# 创建虚拟环境 (推荐)
python -m venv .venv
source .venv\Scripts\activate  # windows
```

### 2. 安装依赖
```bash
pip install -r requirements.txt
```

### 3. 运行
⚠️ **注意：为了确保能切换到任务管理器等高权限窗口，建议以【管理员身份】运行终端。**
```bash
python app.py
```

## 🎮 使用说明 (Usage)

* **Alt + Tab**: 呼出切换器 / 选中下一个窗口。
* **松开 Alt**: 激活当前选中的窗口。
* **右键托盘图标**:
* `设置...` : 打开外观配置面板。
* `退出` : 关闭程序。

## ⚙️ 设置面板 (Configuration)

点击托盘区的设置按钮，即可实时调整：

* **外观颜色**：支持 Hex 颜色输入与取色器。
* **透明度**：0% - 100% 实时预览。
* **布局模式**：一键切换列表或网格。
* **阈值控制**：控制每行显示的最大数量，多余自动换行。

## 📦 打包为 EXE (Build)

如果你想将其打包为独立可执行文件，推荐使用 `PyInstaller`。

```bash
pip install pyinstaller

# 执行打包命令 (需要管理员权限 manifest)
pyinstaller -F -w --uac-admin --icon=icon.ico --add-data "icon.png;." --name="Task Switcher" app.py
# 或
pyinstaller *.spec
```

* `-F`: 单文件模式
* `-w`: 隐藏控制台窗口
* `--uac-admin`: **关键参数**，自动请求管理员权限，确保键盘钩子和窗口切换功能正常。
---

<div align="center">
<sub>Made with 💻 and ☕</sub>

<b>vibe coding</b>
</div>
