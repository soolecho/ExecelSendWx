# 表格自动发送By春风予Lu

基于 PyQt6 开发的 Excel 表格数据自动提取并通过微信发送的自动化工具。

## 功能特点

- **本地 Excel 处理**：支持本地 Excel 文件读取，无需在线文档
- **Sheet 切换**：自动识别所有 Sheet，支持快速切换
- **列选择**：支持多选列进行数据提取（弹出式多选对话框）
- **人员筛选**：根据选择的人员列提取联系人信息
- **微信自动发送**：通过 wxauto4 实现微信消息自动发送
- **发送控制**：支持开始、暂停、停止发送
- **实时日志**：详细的操作日志记录，方便排查问题
- **后台线程**：Excel 读取和发送操作均在后台线程，不阻塞界面

## 环境要求

- Windows 10/11
- Python 3.10+
- 微信 PC 客户端（已登录）

## 安装依赖

```bash
pip install PyQt6 pandas openpyxl wxauto4
```

## 运行方式

### 源码运行

```bash
python main.py
```

### Nuitka 打包

使用 Nuitka 编译为独立 exe 文件：

```bash
python -m nuitka --standalone --windows-icon-from-ico=love.ico --windows-console-mode=disable --enable-plugin=pyqt6 --output-dir=dist main.py
```

## 使用说明

1. **选择 Excel 文件**：点击"选择文件"按钮，选择本地 Excel 文件
2. **选择 Sheet**：在 Sheet 下拉框中选择要处理的工作表
3. **选择提取列**：点击"选择要提取的列"按钮，多选需要的列
4. **选择人员列**：选择包含微信联系人名称的列
5. **加载数据**：点击"加载数据"按钮，提取人员列表
6. **编辑消息**：在消息框中编辑要发送的内容
7. **开始发送**：点击"开始发送"按钮，自动向选中的联系人发送消息

## 项目结构

```
wxauto/
├── main.py              # 程序入口
├── love.ico             # 应用图标
├── modules/
│   ├── __init__.py
│   ├── gui.py           # GUI 界面
│   ├── table_processor.py  # 表格处理
│   └── wechat_sender.py    # 微信发送
└── README.md
```

## 技术栈

- **GUI 框架**：PyQt6
- **数据处理**：pandas + openpyxl
- **微信自动化**：wxauto4
- **打包工具**：Nuitka

## 注意事项

1. 使用前请确保微信 PC 客户端已登录
2. 首次发送时可能需要手动确认微信窗口
3. 建议先使用少量联系人测试功能
4. 发送过程中请勿操作鼠标和键盘

## License

MIT
