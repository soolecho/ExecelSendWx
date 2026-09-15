# 表格自动发送By春风予Lu

基于 PyQt6 开发的 Excel 表格数据自动提取并通过微信发送的自动化工具，
内置定时调度器、托盘驻留、单实例锁、Nuitka standalone 打包与
Inno Setup + Authenticode 自签名安装包发布流程。

当前版本：**v1.1.0**（春风予Lu 自签名 SHA256 证书，指纹
`191C64E4EC07377CA032878878D0A45F554C8146`）。

## 功能特点

- **本地 Excel 处理**：支持本地 Excel 文件读取，无需在线文档
- **Sheet 切换**：自动识别所有 Sheet，支持快速切换
- **列选择**：支持多选列进行数据提取（弹出式多选对话框）
- **人员筛选**：根据选择的人员列提取联系人信息
- **微信自动发送**：通过 wxauto4 实现微信消息自动发送
- **三种发送形式**：支持文字、表格图片、图片后再发文字
- **表格图片分页**：自动生成带表头的表格图片，分页发送后清理临时文件
- **定时发送任务**：独立 Tab 管理周期性发送任务，支持周几 / 多时间点 /
  延迟与间隔，持久化在 `%LOCALAPPDATA%\ExcelSendWx\schedules\schedules.json`
- **发送控制**：支持开始、暂停、停止发送
- **托盘与单实例**：关闭窗口后驻留托盘，重复启动时提示程序正在运行
- **多配置管理**：保存 Excel、列选择和发送设置，支持最近配置双击加载；
  定时任务可单独导出/导入 JSON
- **自动重试**：聊天窗口切换失败时自动重试，最多尝试 3 次
- **断点重试**：发送失败后从未完成的图片页、文字段或自定义消息继续
- **实时日志**：详细的操作日志记录，方便排查问题
- **后台线程**：Excel 读取、发送与定时调度均在后台线程，不阻塞界面
- **稳定性优化**：WeChatSender 进程级单例 + 句柄失效自动重连；
  调度器空 tick 不刷日志；fired_log 增量写盘；日志面板原生
  `setMaximumBlockCount` 自动裁剪
- **精致交互界面**：分组全部改为可折叠芯片按钮排，点击以果冻回弹
  动画展开/收起，窗口尺寸跟随内容弹性自适应；某栏面板全部收起时
  自动收缩为芯片导轨，宽度按比例分配给其他栏
- **任务栏进度显示**：发送进度以绿色进度条显示在 Windows 任务栏
  图标上（失败时变红），配合窗口标题和托盘气泡，进度条带平滑
  增长动画，数字显示真实人数（如 40/40）而非百分比
- **发送逻辑优化**：数据发送与定时任务均实现每个收件人只 ChatWith
  切换一次聊天窗口，后续文字/图片/附件/自定义消息在当前窗口直接
  发送；发送前校验窗口是否仍为目标人，若被切走自动恢复，避免发错
- **智能面板联动**：加载数据后（手动点击「加载数据」或双击加载配置
  自动加载）自动展开人员列表、数据预览、发送进度、发送控制；
  选中人员时自动展开数据预览；点击发送后
  自动收起非必要面板，只保留发送控制和进度
- **定时任务手风琴**：右侧表单（任务基础/重复规则/发送内容/锁定
  行为/日志）一次只展开一组；配置保存区默认收起，表单有改动或
  新建任务时自动弹出，保存/切换后自动收起

## 环境要求

- Windows 10/11
- Python 3.10+（推荐 3.12）
- 微信 PC 客户端（已登录）

## 安装依赖

```bash
pip install PyQt6 pandas openpyxl wxauto4 nuitka
```

## 运行方式

### 源码运行

```bash
python main.py
```

### Nuitka 打包

```bash
python -m nuitka --standalone \
  --windows-icon-from-ico=love.ico \
  --windows-console-mode=disable \
  --enable-plugin=pyqt6 \
  --enable-plugin=tk-inter \
  --output-dir=dist_nuitka_config \
  --output-filename=表格自动发送By春风予Lu.exe \
  main.py
```

或直接运行项目内置的封装脚本（自动开启 `--jobs`、`--lto=no`、
`--assume-yes-for-downloads`，并把 `NUITKA_CACHE_DIR` 指向
`nuitka_cache`，日志落盘到 `logs/`）：

```powershell
python build_nuitka_now.py
```

编译结果位于 `dist_nuitka_config/main.dist/`。发布时必须复制整个
`main.dist` 目录，不能只复制其中的 EXE。

### Inno Setup 安装包

打包好 `main.dist` 后，使用 Inno Setup 6 + 本仓库的 `installer.iss`
生成中文安装包（含桌面/开始菜单快捷方式、开机自启动选项、
安装介绍页 `installer_info.txt`、签名卸载器、签名安装包）：

```powershell
.\build_installer.ps1
```

脚本会自动执行两阶段编译（生成未签名 uninstaller → 签名 → 重编 →
签名安装包），日志写入 `logs/build_installer_YYYYMMDD_HHMMSS.log`，
最终产物在 `installer_output/表格自动发送安装包_v1.1.0.exe`。

签名使用 `CurrentUser\My` 中指纹为
`191C64E4EC07377CA032878878D0A45F554C8146` 的春风予Lu 自签名证书，
时间戳服务器 `http://timestamp.digicert.com`，哈希算法 SHA256。
公钥同时导出到 `installer_output/code_signing_public.cer`。

## 使用说明

### 数据发送 Tab

1. **选择 Excel 文件**：选择本地 `.xlsx` 或 `.xls` 文件
2. **选择 Sheet**：从下拉框选择需要处理的工作表
3. **选择列**：设置人员列、提取列，以及可选的微信昵称列
4. **加载数据**：读取人员列表并选择本次需要发送的人员
5. **设置发送方式**：选择文字、图片或图片后文字，并设置发送延迟
6. **设置自定义消息**：按需启用数据发送后的追加消息
7. **保存配置**：保存当前文件、Sheet、列和发送设置
8. **开始发送**：确认人员列表后开始发送；失败任务可单独重试

### 定时发送 Tab

1. **新建任务**：填写任务名称，选择来源（已保存的 Excel 配置 +
   接收人；或直接指定微信接收人）
2. **设置周期**：勾选周一至周日，添加多个时间点（`QTimeEdit`）
3. **设置发送参数**：人员间隔、聊天窗口延迟、是否启用自定义消息
4. **启用任务**：勾选启用后，后台 `ScheduleDispatcher` 每 10 秒
   tick 一次，命中时间窗口（±20 秒）且当日未触发过即执行
5. **导入/导出**：支持把全部任务导出为
   `{version:1, type:schedule_profile, saved_at, tasks:[...]}` JSON，
   在另一台机器上覆盖或追加导入

调度器在程序退出时会等待当前正在发送的子任务完成，再安全停止；
重复触发的 slot 不会重写 JSON，只新增 slot 才落盘并刷新 `updated_at`。

## 配置管理

左侧「配置管理」支持：

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

默认配置目录为 `%LOCALAPPDATA%\ExcelSendWx\profiles`，定时任务目录为
`%LOCALAPPDATA%\ExcelSendWx\schedules`，两者独立互不污染。
如果 Excel 文件被移动，加载配置时可以重新定位文件；如果表格列发生
变化，程序会提示缺失列。

## 日志与临时文件

- 运行日志：`%LOCALAPPDATA%\ExcelSendWx\logs\app.log`
- 日志单文件最大 5 MB，保留 3 个历史文件
- 打包日志：`logs/build_nuitka_YYYYMMDD_HHMMSS.log`、
  `logs/build_installer_YYYYMMDD_HHMMSS.log`
- 表格图片临时目录：`%TEMP%\wxauto_images`
- 临时图片在发送完成、失败或停止后自动删除

## 项目结构

```
wxauto/
├── main.py                      # 程序入口（日志、异常钩子、信号处理）
├── love.ico                     # 应用图标
├── README.md
├── .gitignore
├── modules/
│   ├── __init__.py
│   ├── config_manager.py        # 配置文件和最近配置管理
│   ├── gui.py                   # GUI 界面（数据发送 Tab + 定时发送 Tab）
│   ├── schedule_manager.py      # 定时任务存储 + 调度器（10s tick）
│   ├── table_processor.py       # 表格处理
│   └── wechat_sender.py         # 微信发送（v2 单例 + 自动重连）
├── build_nuitka_now.py          # Nuitka standalone 编译入口
├── build_installer.ps1          # Inno Setup + 签名流水线
├── installer.iss                # Inno Setup 脚本
├── installer_info.txt           # 安装介绍页文案
└── installer_assets/
    └── ChineseSimplified.isl    # Inno Setup 简体中文语言包
```

## 技术栈

- **GUI 框架**：PyQt6
- **数据处理**：pandas + openpyxl
- **微信自动化**：wxauto4
- **打包工具**：Nuitka 4.1.x（standalone + pyqt6/tk-inter 插件）
- **安装包**：Inno Setup 6 + Authenticode SHA256 自签名 +
  Digicert 时间戳

## 注意事项

1. 使用前请确保微信 PC 客户端已登录
2. 微信联系人使用模糊搜索，请确保保存的昵称能够正确匹配目标
3. 建议先使用少量联系人测试配置和发送内容
4. 发送过程中尽量不要操作微信窗口、鼠标和键盘
5. Nuitka 版本必须携带整个 `main.dist` 目录运行
6. 点击窗口右上角关闭按钮只会隐藏到托盘
7. 完全退出程序请右键托盘图标并选择「退出程序」
8. 安装包通过 `表格自动发送安装包_v1.1.0.exe` 安装，可在安装时勾选
   「开机自启动」，安装后通过控制面板卸载；卸载器同样已签名

## License

MIT

## 作者

春风予Lu
