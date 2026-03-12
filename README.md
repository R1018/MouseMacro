# MouseMacro - 鼠标操作录制工具

<div align="center">
![MouseMacro](https://img.shields.io/badge/MouseMacro-v2.0.0-blue?style=for-the-badge)
![Python](https://img.shields.io/badge/Python-3.7+-green?style=for-the-badge&logo=python)
![Windows](https://img.shields.io/badge/Windows-10/11-success?style=for-the-badge&logo=windows)
![License](https://img.shields.io/badge/License-MIT-yellow?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Active-success?style=for-the-badge)

**一个简单易用的 Windows 鼠标操作录制与回放工具，支持录制指定窗口内的鼠标移动和点击操作，并可重复回放。**

[功能特点](#功能特点) • [快速开始](#快速开始) • [使用说明](#使用说明) • [常见问题](#常见问题) • [更新日志](CHANGELOG.md)

</div>

---

## 功能特点

- 🎯 **窗口录制**：录制指定窗口内的鼠标操作（移动、点击、拖动）
- 🔄 **重复回放**：可设置回放次数，自动重复执行录制好的操作
- 🎨 **主题切换**：支持明亮、暗黑、跟随系统三种主题模式，并记住用户选择
- 🖥️ **界面友好**：简洁直观的 GUI 界面，操作简单
- 🔍 **窗口选择**：支持从窗口列表中点击选择，支持搜索过滤，无需手动输入
- 💾 **脚本管理**：可保存多个录制脚本，方便管理和调用
- ⚙️ **可配置**：支持自定义拖动记录间隔
- ✅ **窗口验证**：录制前可验证目标窗口是否存在
- 🛑 **停止控制**：支持中途停止录制和回放
- 📊 **状态显示**：实时显示录制/回放状态
- 💾 **配置保存**：自动保存主题、最后使用的脚本等设置

## 环境要求

- **操作系统**：Windows 10/11
- **Python 版本**：Python 3.7 或更高版本
- **权限**：需要管理员权限（用于访问其他窗口）

## 快速开始

### 1. 克隆或下载项目

```bash
git clone https://github.com/yourusername/MouseMacro.git
cd MouseMacro
```

或直接下载 ZIP 文件并解压。

### 2. 安装依赖

使用管理员权限打开命令行（PowerShell 或 CMD），在项目目录执行：

```bash
pip install -r requirements.txt
```

如果安装失败，可以尝试分别安装：

```bash
pip install pywin32>=305
pip install pynput>=1.7.6
```

### 3. 运行程序

**方法一：使用批处理文件（Windows 推荐）**

双击 `启动程序.bat` 文件即可运行。

**方法二：使用命令行**

```bash
python main.py
```

程序会自动隐藏控制台窗口，只显示 GUI 界面。

> 💡 **提示**：首次运行建议先执行 `安装依赖.bat` 来安装所有必需的包。

## 使用说明

### 1. 启动程序

根据上面的"快速开始"部分运行程序。

### 2. 录制操作

1. 点击"录制"标签页
2. 在"窗口标题"框中输入要录制的窗口名称，或点击"选择窗口"按钮从列表中选择
3. 设置"拖动记录间隔"（默认 0.01 秒）
4. 选择录制文件保存位置（默认保存在 `Scripts` 目录）
5. 点击"开始录制"按钮
6. 在目标窗口中执行鼠标操作
7. 按 `ESC` 键停止录制

### 3. 回放操作

1. 点击"回放"标签页
2. 选择要回放的录制文件
3. 设置回放次数
4. 点击"开始回放"按钮

### 4. 主题切换

点击顶部主题按钮切换：
- ☀ 明亮主题
- ☾ 暗黑主题
- 🔄 跟随系统主题

主题选择会自动保存，下次启动时自动应用。

### 5. 高级功能

- **检查窗口**：点击"检查窗口"按钮验证目标窗口是否存在
- **停止录制**：录制过程中可随时点击"停止录制"按钮
- **停止回放**：回放过程中可随时点击"停止回放"按钮
- **脚本列表**：点击"刷新脚本列表"快速选择已保存的脚本
- **日志控制**：顶部按钮可开启/关闭控制台日志输出

## 项目结构

```
MouseMacro/
├── main.py              # 主程序文件
├── README.md            # 使用说明文档
├── requirements.txt     # Python 依赖包列表
├── .gitignore          # Git 忽略文件配置
├── Scripts/            # 录制脚本目录（自动创建）
│   └── .gitkeep       # 目录占位文件
├── config.json         # 配置文件（运行时自动生成）
└── debug.log          # 调试日志（运行时自动生成）
```

## 配置说明

程序首次运行时会自动在当前目录创建 `config.json` 文件，包含以下配置：

```json
{
  "theme": "auto",              // 主题: auto/light/dark
  "last_script": "",            // 最后使用的脚本路径
  "last_window": "",            // 最后录制的窗口标题
  "log_enabled": true           // 是否启用日志输出
}
```

可以手动修改此文件来更改默认设置。

## 故障排除

## 录制文件格式

录制文件以 JSON 格式保存在 `Scripts` 目录下，包含以下信息：

```json
{
  "window_title": "窗口标题",
  "initial_rect": [left, top, right, bottom],
  "drag_interval": 0.01,
  "events": [
    {
      "type": "click|move",
      "x": 0.5,
      "y": 0.5,
      "button": "left",
      "state": "pressed|released",
      "time": 1.23
    }
  ]
}
```

**注意**：坐标使用相对值（0.0-1.0），确保在不同窗口大小下都能正确回放。

## 键盘快捷键

| 快捷键 | 功能 |
|--------|------|
| ESC | 停止录制 |
| Enter | 确认选择（窗口/脚本） |
| 双击 | 快速选择（窗口/脚本） |

## 注意事项

- 录制时确保目标窗口在前台且可见
- 回放时不要移动鼠标，以免干扰
- 不同分辨率下可能需要重新录制
- 建议录制前先测试目标窗口位置

## 目录结构

```
MouseMacro/
├── main.py              # 主程序
├── README.md            # 使用说明
├── Scripts/             # 录制脚本目录
│   ├── mouse_recording.json
│   └── ...
└── debug.log            # 调试日志（运行时生成）
```

## 常见问题

### 运行问题

**Q: 程序一闪而过，无法启动？**

A:
1. 确保已安装所有依赖：`pip install -r requirements.txt`
2. 尝试使用管理员权限运行
3. 检查 Python 版本是否为 3.7 或更高

**Q: 提示 "ModuleNotFoundError: No module named 'pywin32'"？**

A: 使用管理员权限运行命令行，然后执行：
```bash
pip install pywin32
```

**Q: 提示权限错误？**

A: 该程序需要访问其他窗口，请以管理员身份运行：
- 右键点击命令行（CMD 或 PowerShell）
- 选择"以管理员身份运行"
- 在该窗口中执行 `python main.py`

### 使用问题

**Q: 录制时提示"窗口未找到"？**

A: 请确保目标窗口已打开，且窗口标题准确。建议使用"选择窗口"功能，并点击"检查窗口"验证。

**Q: 回放时位置不准确？**

A: 可能是窗口位置或大小发生了变化。请确保回放时目标窗口的位置和大小与录制时一致。

**Q: 如何中途停止录制或回放？**

A:
- 录制：按 `ESC` 键或点击"停止录制"按钮
- 回放：点击"停止回放"按钮

**Q: 控制台窗口一直显示，如何隐藏？**

A: 程序默认会隐藏控制台窗口。如果控制台一直显示，可能是运行方式的问题。直接双击 `main.py` 文件或在命令行运行 `pythonw main.py` 可以确保无控制台窗口。

### 配置问题

**Q: 配置文件保存在哪里？**

A: 配置文件 `config.json` 保存在程序目录下，包含主题、最后使用的脚本等设置。可以手动编辑此文件来更改默认设置。

**Q: 如何清空日志文件？**

A: 程序会自动管理日志文件大小，超过 5MB 会自动清空。也可以手动删除 `debug.log` 文件。

**Q: 如何重置所有设置？**

A: 删除 `config.json` 文件，程序会重新创建默认配置。

## 开发与贡献

### 开发环境设置

1. Fork 本仓库
2. 创建特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 开启 Pull Request

### 代码规范

- 遵循 PEP 8 代码风格
- 添加适当的注释和文档字符串
- 确保代码在不同环境下可正常运行

## 许可证

本项目采用 MIT 许可证 - 详见 [LICENSE](LICENSE) 文件

## 致谢

- [pywin32](https://github.com/mhammond/pywin32) - Windows API 的 Python 封装
- [pynput](https://github.com/moses-palmer/pynput) - 控制输入设备（鼠标和键盘）

---

<div align="center">
**如果这个项目对你有帮助，请给个 ⭐ Star 支持一下！**

</div>

## 许可证

本项目仅供学习和个人使用。

## 版本历史

### v2.0

- 新增主题切换功能（明亮/暗黑/跟随系统）
- 隐藏控制台窗口
- 改进窗口选择功能，支持点击选择
- 优化用户界面

### v1.0
- 基础录制和回放功能
- 支持 JSON 格式保存
- GUI 界面
