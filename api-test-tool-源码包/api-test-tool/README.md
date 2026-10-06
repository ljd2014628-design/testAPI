# 接口测试工具

一个本地运行的 Windows 桌面接口测试工具。用 Excel 管理接口用例，一键批量执行，
支持多接口串联、变量传递、断言校验，结果可回写 Excel。

双击 exe 就能用，不依赖浏览器，不需要装 Python。

## 功能

- **Excel 驱动**：接口用例在 Excel 里维护，导入执行，结果回写；每个用例保留最近 5 次执行结果
- **cURL 导入**：粘贴浏览器 F12「复制为 cURL」(bash / cmd)、Postman / Apifox 导出的 cURL，自动解析成用例
- **请求参数表**：GET 的查询参数以「参数名 / 参数值 / 备注」表格维护，可单独停用某个参数
- **字段备注**：给 Body 里的每个字段写注释，只本地保存，不参与请求
- **变量传递**：从响应提取变量（点响应树即可生成 JSON 路径），后续接口用 `${变量名}` 引用
- **组合场景**：多接口按步骤串联执行，结果单独记录到 Excel Sheet2
- **断言**：40+ 比较器，支持「较上次减少 1」这类相对比较，适合领券 / 扣库存等每次都变的数据
- **筛选**：按接口名称、执行结果、断言结果、关键字筛选，几千条用例也不卡
- **请求头优先级**：接口级请求头覆盖公共请求头（同名不区分大小写）

## 界面

主界面分四块：公共配置（BaseURL / 编码 / 超时 / 公共请求头）、执行区、
接口用例列表（带筛选栏）、组合场景与执行日志页签。

接口编辑窗口分三段：接口级请求头、请求参数或 Body（互斥，另有 Body 字段备注页）、
提取变量与断言。

## 快速开始

### 方式一：直接跑源码

```bash
git clone <your-repo-url>
cd api-test-tool
pip install -r requirements.txt
python gui.py
```

需要 Python 3.8+，且自带 tkinter（Windows 官方安装包默认包含；
Ubuntu 需要 `sudo apt install python3-tk`；macOS 建议用 python.org 的安装包）。

### 方式二：用打包好的 exe

从 Releases 下载 `接口测试工具.exe`，双击运行。首次启动解压内置运行环境需要 5~15 秒。

### 先跑通一遍

仓库自带一个假接口服务，不用连真实环境也能体验完整流程：

```bash
python tools/mock_server.py      # 监听 127.0.0.1:9911
```

然后在工具里把 BaseURL 填成 `http://127.0.0.1:9911`，
点「生成Excel模板」存下来再「导入Excel」，点「全部执行」即可看到效果。
mock 提供了 `/coupon/receive`（每调一次库存减一，用来演示「较上次减少 1」的断言）、
`/coupon/wallet/list`、`/marketing/promotion/findAll` 等接口。

## 打包成 exe

在 Windows 上：

```bat
scripts\build_exe.bat
```

或手动：

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name ApiTestTool gui.py
```

产物在 `dist/ApiTestTool.exe`，改名即可分发。

## 目录结构

```
api-test-tool/
├── gui.py                    # 界面层（Tkinter），程序入口
├── core.py                   # 核心层：请求执行、变量提取、断言判定、cURL 解析
├── excel_io.py               # Excel 模板生成、导入解析、结果导出
├── requirements.txt
├── docs/
│   ├── 使用说明.md            # 完整使用文档
│   ├── 断言使用手册.md         # 断言比较器速查与写法示例
│   └── 接口测试模板.xlsx       # Excel 模板（含示例行）
├── tools/
│   └── mock_server.py        # 本地假接口服务，用于演示和自测
├── tests/
│   └── test_core.py          # 核心逻辑自测，73 项断言
└── scripts/
    └── build_exe.bat         # Windows 打包脚本
```

## 运行自测

```bash
python tests/test_core.py
```

覆盖 JSON 路径解析、变量替换、请求头合并优先级、URL 参数拼接、cURL 解析（bash / cmd / 多条）、
26 种断言比较器、Excel 导出导入往返。全部通过时退出码为 0，可直接接 CI。

## 数据存放

所有配置、用例、执行结果存在程序同级的 `data/state.json`，关掉再开自动恢复。
放在只读目录时会自动改存到 `~/ApiTestTool/data`。该目录已在 `.gitignore` 中排除。

## Excel 结构

| Sheet | 作用 |
|---|---|
| 接口用例 | 接口清单 + URL参数 + 字段备注 + 提取变量 + 断言 + 最近 5 次返回结果 |
| 组合场景结果 | 多接口场景的最新一次结果，按步骤按行记录 |
| 场景配置 | 组合场景的步骤编排 |
| 公共配置 | BaseURL / 编码 / 超时 / 校验SSL / 公共请求头 |

详见 [docs/使用说明.md](docs/使用说明.md)。

## 技术说明

- 界面：Tkinter + ttk，无第三方 UI 依赖
- 请求：requests，请求在后台线程执行，界面不阻塞
- Excel：openpyxl
- 打包：PyInstaller onefile + windowed
- 状态持久化：延迟合并写盘（1.2 秒内的多次改动合并成一次），关闭和导出前强制落盘

## 已知限制

- cURL 的 `-F` 表单 / 文件上传暂不支持（解析时会给出提示）
- 单条响应留存上限 6 万字符，超出部分在查看结果时标注截断
- 目前只支持 HTTP(S)，不支持 WebSocket / gRPC

## License

MIT
