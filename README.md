# Windows-Pin

Windows-Pin 是一个用于设置 Windows 应用窗口置顶优先级的简单程序，旨在满足多任务显示排布需求。

## 功能特点

- 动态显示当前所有可见窗口
- 通过拖拽调整窗口优先级
- 一键设置窗口置顶或置底
- 支持最小化到系统托盘，便于随时调用
- 保持自身界面始终置顶，方便操作

## 技术实现

Windows-Pin 主要使用以下技术：

- Python：作为主要编程语言
- Tkinter：用于创建图形用户界面
- Win32 API：通过 `win32gui` 和 `win32con` 模块实现窗口操作

## 使用方法

1. 运行程序后，会显示当前所有可见窗口的列表。
2. 通过拖拽列表项来调整窗口优先级。
3. 选中某个窗口后，可以使用"Set Top"或"Set Bottom"按钮快速调整其置顶状态。
4. 使用"Minimize"按钮可将程序最小化到系统托盘，点击托盘图标可再次打开主界面。

## 安装

1. 确保您的系统已安装 Python 3.x。
2. 克隆此仓库：
   ```bash
   git clone https://github.com/your-username/Windows-Pin.git
   ```
3. 进入项目目录：
   ```bash
   cd Windows-Pin
   ```
4. 安装所需依赖：
   ```bash
   pip install -r requirements.txt
   ```

## 运行

在项目目录下运行以下命令：

```bash
    python app.py
```

