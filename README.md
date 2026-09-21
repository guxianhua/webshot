# webshot

批量网站可用性巡检与截图相册系统。

与传统漏洞扫描器不同，webshot 的核心目标是**快速确认一批网站"能不能打开、长什么样"**：管理员上传一个 Excel 清单，点一下开始巡检，几分钟后得到一组截图相册，肉眼浏览即可定位打不开的、空白的、跳到登录页的网站。

## 功能特性

- 📊 **Excel 驱动** — 上传 .xlsx 网站清单（名称、URL、责任单位），支持清单内直接编辑和删除
- 🔍 **智能状态识别** — 自动区分「正常 / 无法访问 / 空白页 / 登录页面 / 正常但无截图」五种状态
- 📸 **并发截图** — Playwright 驱动 Chromium 无头浏览器，8 并发同时截图，200 站点约 3-5 分钟完成
- 🖼️ **相册式展示** — 巡检端以卡片网格展示截图，支持悬浮标签切换（放大 / 打开网页 / 复制名称），点击放大、任意点关闭
- ⏯️ **实时进度** — 首页播放器风格按钮 + 进度段，显示距离上次巡检时间，巡检端边跑边加载图片
- ⛔ **随时停止** — 巡检过程中可一键停止，线程安全退出，不会卡死
- 🔄 **数据隔离** — 更换 Excel 或编辑列表后，自动清空相册和历史任务，保证数据一致性
- 💻 **本地运行** — 单台机器即可运行，可选 PyInstaller 打包为 exe 开箱即用

## 快速开始

`ash
# 1. 安装依赖
pip install flask openpyxl requests urllib3 pillow playwright

# 2. 安装 Playwright Chromium 浏览器
playwright install chromium

# 3. 启动服务（开发模式）
python app.py

# 4. 浏览器访问
# http://127.0.0.1:5000
`

生产模式（PyInstaller 打包 exe）：

`ash
# 打包
pyinstaller launcher.py --windowed --name webshot --add-data "templates;templates" --add-data "playwright-browsers;playwright-browsers"

# 运行 dist/webshot/webshot.exe
# 在弹出的 GUI 中选择工作目录和端口，点击"启动服务"
`

## 代码结构

`
webshot/
├── app.py              # Flask 后端（路由 + 巡检核心逻辑）
├── launcher.py         # Tkinter GUI 启动器（可选）
├── .gitignore
└── templates/
    ├── index.html      # 首页（巡检任务控制 + 模式选择）
    ├── admin.html      # 管理端（Excel 上传 / 网站列表编辑）
    └── user.html       # 巡检端（截图相册）
`

## 核心逻辑

### 1. 巡检任务生命周期

`
用户点击"开始巡检"
    │
    ▼
POST /api/inspect  →  创建 task_id，启动后台线程 run_inspect_task()
    │
    ▼
asyncio.run(run_inspect_async())
    ├─ 创建 8 并发 Semaphore
    ├─ 启动单个 Chromium 浏览器实例
    ├─ 并发执行每个网站：
    │   ├─ HTTP 探测（requests.get，8s 超时）
    │   ├─ 截图（Playwright，domcontentloaded 等待 + 1.5s 稳定）
    │   └─ 失败时重试一次（commit 等待 + 快速截图）
    │
    ▼
截图后自动标记状态：
    ├─ HTTP 2xx + 截图成功 + 非空白 + 非登录页 → 正常
    ├─ HTTP 2xx + 空白图片（99% 白色像素）     → 空白页
    ├─ HTTP 2xx + 跳转 auth.shsmu.edu.cn      → 登录页面
    ├─ HTTP 2xx + 两次截图都失败               → 正常但无截图
    └─ 连接超时 / 4xx / 5xx                    → 无法访问
    │
    ▼
每完成一个网站，加锁写入 tasks/{task_id}.json
前端轮询 /api/latest-task 实时更新进度条和相册
`

### 2. 停止机制

使用 	hreading.Event 作为停止信号：

- 用户点"停止" → POST /api/inspect/stop → stop_flags[task_id].set()
- asyncio 中创建 watch_stop 协程，每 0.5s 检查一次，发现停止信号后：
  - syncio.gather 中的剩余任务被 cancel
  - 最终关闭浏览器、写入终态 status: 'stopped'

### 3. 数据一致性

- **sites_data.json** — 当前网站清单（上传或编辑后更新）
- **sites_meta.json** — 清单元数据（原始文件名、数据更新时间戳）
- **tasks/*.json** — 巡检任务记录（只保留最新一次，列表更新后过期）
- **screenshots/** — 截图文件（文件名带 task_id 前缀，列表更新后整体清理）

更新 Excel 或编辑列表时，	ouch_data_updated_at() 记录新时间戳，前端发现 	ask.data_updated_at != meta.data_updated_at 时显示"网站列表已更新"并清空相册。

### 4. 路径与部署

`
开发模式：
  BASE_DIR = 项目根目录
  WORK_DIR = 项目根目录（默认）或 WORK_DIR 环境变量
  PLAYWRIGHT_BROWSERS_PATH = 系统默认 ~/ms-playwright

PyInstaller 打包后：
  BASE_DIR = sys._MEIPASS（临时解压目录，只读）
  WORK_DIR = 用户在 GUI 中选择的工作目录
  PLAYWRIGHT_BROWSERS_PATH = BASE_DIR/playwright-browsers（内嵌）
`

### 5. 任务状态清理

Flask debug reloader 重启或进程崩溃后，	asks/*.json 中可能残留 status: 'running' 的孤儿任务。pp.py 启动时自动将这些孤儿标记为 stopped，避免前端永远显示进度中。

## API 接口

| 方法 | 路径 | 说明 |
|------|------|------|
| POST | /api/upload | 上传 .xlsx，返回网站列表 |
| GET  | /api/sites | 获取当前网站清单 |
| GET  | /api/sites/meta | 获取清单元数据 |
| POST | /api/sites/save | 保存编辑后的清单 |
| DELETE | /api/sites/delete/{id} | 删除单个网站 |
| POST | /api/inspect | 发起巡检，返回 task_id |
| POST | /api/inspect/stop | 停止当前巡检 |
| GET  | /api/latest-task | 获取最新一次巡检状态（含过期判断） |
| GET  | /api/screenshot/{filename} | 返回截图文件 |

## 技术栈

| 组件 | 技术 |
|------|------|
| 后端 | Flask 3.x |
| 异步截图 | Playwright (Chromium headless, 8 并发) |
| HTTP 探测 | requests |
| 图片判断 | Pillow |
| Excel 解析 | openpyxl |
| 前端 | 原生 HTML/CSS/JS（无构建工具） |
| 可选 GUI | Tkinter |
| 可选打包 | PyInstaller |

## Excel 格式

上传的 .xlsx 文件需要包含以下列：

| 列 | 内容 |
|----|------|
| B | 网站名称 |
| C | URL（支持单元格内换行分隔多个） |
| D | 责任单位 |
| E | 访问方式（可选） |

## License

MIT
