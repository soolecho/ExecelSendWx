# 表格自动发送By春风予Lu

基于 PyQt6 开发的 Excel 表格数据自动提取并通过微信发送的自动化工具。

## 功能特点

- **本地 Excel 处理**：支持本地 Excel 文件读取，无需在线文档
- **Sheet 切换**：自动识别所有 Sheet，支持快速切换
- **列选择**：支持多选列进行数据提取（弹出式多选对话框）
- **人员筛选**：根据选择的人员列提取联系人信息
- **微信自动发送**：通过 wxauto4 实现微信消息自动发送
- **三种发送形式**：支持文字、表格图片、图片后再发文字
- **表格图片分页**：自动生成带表头的表格图片，分页发送后清理临时文件
- **发送控制**：支持开始、暂停、停止发送
- **多配置管理**：保存 Excel、列选择和发送设置，支持最近配置双击加载
- **自动重试**：聊天窗口切换失败时自动重试，最多尝试 3 次
- **断点重试**：发送失败后从未完成的图片页、文字段或自定义消息继续
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

安装 Nuitka：

```bash
pip install nuitka
```

使用多文件模式编译，不使用运行时解压的单文件模式：

```bash
python -m nuitka --standalone --windows-icon-from-ico=love.ico --windows-console-mode=disable --enable-plugin=pyqt6 --enable-plugin=tk-inter --output-dir=dist_nuitka_config --output-filename=表格自动发送By春风予Lu.exe main.py
```

编译结果位于 `dist_nuitka_config/main.dist/`。发布时必须复制整个
`main.dist` 目录，不能只复制其中的 EXE。

## 使用说明

1. **选择 Excel 文件**：选择本地 `.xlsx` 或 `.xls` 文件
2. **选择 Sheet**：从下拉框选择需要处理的工作表
3. **选择列**：设置人员列、提取列，以及可选的微信昵称列
4. **加载数据**：读取人员列表并选择本次需要发送的人员
5. **设置发送方式**：选择文字、图片或图片后文字，并设置发送延迟
6. **设置自定义消息**：按需启用数据发送后的追加消息
7. **保存配置**：保存当前文件、Sheet、列和发送设置
8. **开始发送**：确认人员列表后开始发送；失败任务可单独重试

## 配置管理

左侧“配置管理”支持：

- **保存配置**：更新当前已经加载的配置文件
- **另存为配置**：创建新的 JSON 配置，可保存多个不同任务
- **加载配置**：读取配置后自动打开 Excel、切换 Sheet、恢复列和发送设置
- **最近配置**：双击最近列表中的配置即可加载

配置中会保存：

- Excel 文件路径和 Sheet
- 人员列、微信昵称列和提取列
- 手动指定的微信接收人
- 文字、图片或图片后文字的发送形式
- 人员发送间隔和聊天窗口延迟
- 自定义消息开关及内容

默认配置目录为 `%LOCALAPPDATA%\ExcelSendWx\profiles`。如果 Excel 文件被
移动，加载配置时可以重新定位文件；如果表格列发生变化，程序会提示缺失列。

## 日志与临时文件

- 运行日志：`%LOCALAPPDATA%\ExcelSendWx\logs\app.log`
- 日志单文件最大 5 MB，保留 3 个历史文件
- 表格图片临时目录：`%TEMP%\wxauto_images`
- 临时图片在发送完成、失败或停止后自动删除

## 项目结构

```
wxauto/
├── main.py              # 程序入口
├── love.ico             # 应用图标
├── modules/
│   ├── __init__.py
│   ├── config_manager.py  # 配置文件和最近配置管理
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
2. 微信联系人使用模糊搜索，请确保保存的昵称能够正确匹配目标
3. 建议先使用少量联系人测试配置和发送内容
4. 发送过程中尽量不要操作微信窗口、鼠标和键盘
5. Nuitka 版本必须携带整个 `main.dist` 目录运行

## License

MIT
