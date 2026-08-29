# Fluent Tools CLI — 把 ansys-fluent-mcp 的工具面变成独立 Python 脚本

`fluent_cli.py` 把 `ansys-fluent-mcp` MCP 服务器暴露的全部工具提取成**单入口命令行工具**,
不启动 MCP 服务器、不依赖 MCP 客户端。每个子命令与 MCP 工具一一对应,**直接复用项目后端层**
(`solve/backends/pyfluent.py` + `solve/lib/*`)的同一份实现,因此返回的 JSON 信封与 MCP 工具
完全一致。输出为 stdout 上的纯 JSON(或用 `--json-out` 直接写入文件),便于 skill / agent 解析。

## 工作原理

MCP 服务器在一个常驻进程里持有 PyFluent 会话;CLI 每次调用都是一个新进程,靠**会话文件**衔接:

1. `connect` 启动(或附着)Fluent,把活动会话的 `connection_properties`(ip/port/password)
   持久化到会话文件。启动时强制 `cleanup_on_exit=False`,Fluent 进程在脚本退出后继续存活。
2. 其余命令读取会话文件,通过 gRPC **附着**到正在运行的 Fluent(秒级),执行后即退出,
   不触碰 Fluent 进程本身。
3. `disconnect` 以 `cleanup_on_exit=True` 附着并退出 Fluent,校验进程真正终止后删除会话文件。

会话文件默认位于 `~/.fluent_tools/session.json`,可用 `--session-file` 或环境变量
`FLUENT_TOOLS_SESSION` 覆盖。Fluent 启动日志与控制台输出重定向到
`~/.fluent_tools/fluent_launch.log`。

## 环境要求

- Python ≥ 3.12,且环境里装有 `ansys-fluent-core`(PyFluent)与 `pydantic`。
  **本机已验证的解释器**:`c:\programdata\anaconda3\envs\pyansys\python.exe`
  (PyFluent 0.41.0 + ANSYS 2026R1,与 fluent-cfd-skill 的 `check_env.py` 检测结果一致)。
- 工具代码来源二选一:本仓库 checkout(脚本自动把 `../src` 加入 `sys.path`,
  优先于已安装包),或 `pip install ansys-fluent-mcp` 后从任意目录运行。
- `find_api`、离线 `validate_code` 不需要 Fluent;其余命令需要 `connect` 后的活动会话。

## 快速开始

```bash
PY="c:/programdata/anaconda3/envs/pyansys/python.exe"
CLI="C:/Users/Administrator/Desktop/AI/pyfluent-mcp-main/tools/fluent_cli.py"

# 1. 启动 Fluent(无 GUI 但可截图;必填维度与网格一致,否则读 case 报
#    "File has wrong dimensions")
"$PY" "$CLI" connect --dimension 2 --processor-count 2

# 2. 任意工具命令:自动附着到上面的会话
"$PY" "$CLI" run_code --code "solver.settings.file.read_case(file_name=r'D:\work\case.cas.h5')"
"$PY" "$CLI" mesh_quality
"$PY" "$CLI" mesh_quality --pretty --json-out result/mesh_quality.json   # 结果直接落盘,父目录自动创建
"$PY" "$CLI" screenshot --out mesh.png

# 3. 结束后关闭 Fluent(会校验进程终止)
"$PY" "$CLI" disconnect
```

## 子命令一览(MCP 工具 ↔ CLI 对照)

| MCP 工具 | CLI 子命令 | 需要会话 | 说明 |
|---|---|---|---|
| `connect` | `connect` | — | 启动/附着 Fluent 并持久化会话;已有活动会话时直接报告(`--force` 强制重启) |
| `disconnect` | `disconnect` | ✓ | 退出 Fluent、删除会话文件(`--timeout` 秒后强杀) |
| `session_status` | `session_status` | — | 会话文件状态 + gRPC 探活(`--no-ping` 跳过) |
| `solver_status` | `solver_status` | ✓ | 初始化/迭代数/残差摘要 |
| `find_api` | `find_api --query ...` | 离线可用 | 内置目录 BM25 检索(`--kind` `--under` `--top-k` `--compact`) |
| `get_help` | `get_help --path ...` | ✓ | 路径的 docstring/子名/allowed values |
| `get_state` | `get_state [--path ...] [--key ...]` | ✓ | 读活动设置状态;`--key` 取单个命名对象 |
| `get_targeted_context` | `get_targeted_context --path ...` | ✓ | 批量 active+state+子名+allowed 一次取齐 |
| `list_named_objects` | `list_named_objects [--limit --offset]` | ✓ | 命名对象集合分页枚举 |
| `find_named_object` | `find_named_object --name ...` | ✓ | 符号名跨集合解析 |
| `select_named_objects` | `select_named_objects --collection ...` | ✓ | glob 展开选择(`--pattern` `--exclude` `--no-include-shadows`) |
| `run_code` | `run_code --code / --code-file` | ✓ | 沙箱 Python,预注入 `solver`/`session` |
| `validate_code` | `validate_code --code / --code-file` | 离线可用 | 语法/安全/路径预检,无副作用 |
| `screenshot` | `screenshot --out x.png` | ✓ | 截图**保存为 PNG 文件**(MCP 返回 base64,CLI 落盘更适合脚本) |
| `summarize_setup` | `summarize_setup` | ✓ | Report > Summary 全文 |
| `simulation_report` | `simulation_report --action ...` | ✓ | generate / export_html / export_pdf / export_pptx / list |
| `mesh_quality` | `mesh_quality [--include-check]` | ✓ | 网格数量 + 偏斜度/正交质量 |
| `list_fields` | `list_fields [--scope]` | ✓ | 场变量枚举 |
| `compare_files` | `compare_files --a x.cas.h5 --b y.cas.h5` | — | 两个临时会话中对比 case/网格文件 |
| `probe_path` | `probe_path --path ...` | ✓ | 批量 exists/active/creatable/kind 预检 |
| `get_active_status` | `get_active_status --path ...` | ✓ | 批量 `{path: bool}` |
| `get_allowed_values` | `get_allowed_values --path ...` | ✓ | 批量枚举合法值 |
| `describe_named_object_template` | `describe_named_object_template --path ...` | ✓ | 命名对象子项字段模板 |
| `describe_path` | `describe_path --path ...` | ✓ | probe+values+模板+命令参数统一描述符 |

> MCP 的 `manage_component`(activate/deactivate/update/refresh)属于 Fluids One 产品层,
> PyFluent 后端不支持,故未纳入 CLI。

除表中各子命令自己的参数外,所有子命令还共享 `--session-file`、`--pretty` 与
`--json-out`(JSON 结果落盘,见"输出契约")。

## 输出契约

- **stdout 只输出一个 JSON 文档**(默认紧凑,`--pretty` 缩进)。Fluent 的控制台/转录噪声
  已被吞掉;启动日志在 `~/.fluent_tools/fluent_launch.log`;库日志在 stderr,仅
  `FLUENT_CLI_DEBUG=1` 时开启(同时保留原始噪声、异常带栈)。
- **结果落盘(`--json-out <file>`)**:JSON 信封改写指定文件(UTF-8,结尾带换行,
  父目录自动创建),stdout 保持为空,stderr 打印一行 `result written to <file>`;
  `--pretty` 同样作用于文件内容。**错误信封也写入该文件**(如无会话时的退出码 3
  错误),因此文件内容始终完整,脚本可只看退出码判断成败、失败时读文件取详情。
  注意 `screenshot` 的 `--out` 仍指 PNG 图片路径,不受影响(其 JSON 元数据可另用
  `--json-out` 保存)。
- **退出码**:`0` 成功;`1` 工具/运行错误(含结果信封 `status:"error"` 或带 `error` 键,
  如 run_code 执行异常、validate_code 校验失败);`2` 参数错误;`3` 无活动会话或会话失活
  (应重新 `connect`)。shell 自动化可只看退出码,需要细节再读 JSON。

## 在 skill 中使用

1. 解释器:用装有 PyFluent 的环境(本机为 `c:\programdata\anaconda3\envs\pyansys\python.exe`,
   可复用 skill 的 `check_env.py`/`env_detected.json` 产物)。
2. 流程:`session_status` → (无会话则)`connect --dimension <2|3> --processor-count <n>` →
   按需调用工具 → `disconnect`。`connect` 自带防重复启动守卫,重复调用安全。
3. `run_code` 是沙箱:仅允许 `math/json/itertools/functools/collections/dataclasses/typing/
   ansys.fluent.core` 导入,预注入 `solver`(别名 `session`),禁止反射写入。
   需要完整自由 Python 的场景,仍走 skill 原有的"生成独立 .py 脚本直接运行"模式。
4. `--code-file path/to/x.py` 适合长代码;`--path` 可重复传参实现批量调用。
5. 写操作前的标准姿势:`probe_path`/`get_allowed_values`/`describe_path` 先探后写,
   与 MCP 工具的用法说明一致。

## 已实测项(本机 ANSYS 2026R1 + PyFluent 0.41)

connect/守卫、disconnect(进程确认终止)、session_status、run_code(含错误信封与
转录捕获)、validate_code(离线)、find_api(离线)、solver_status、get_state、
get_allowed_values、get_targeted_context、describe_path、
describe_named_object_template、probe_path、list_named_objects(分页)、
find_named_object、select_named_objects、summarize_setup、simulation_report(list)、
mesh_quality、list_fields、screenshot(hidden_gui 下对已 display 的图形对象出图成功)。

## 注意事项与差异

- **维度必须匹配**:`connect --dimension` 要与网格一致,否则 `read_case` 报
  "File has wrong dimensions"——与 skill 文档中的坑位一致。
- **截图前提**:默认 `--ui-mode hidden_gui`(无窗口但可离屏渲染)。`no_gui` 更轻但
  无法出图;且须先用 `run_code` 创建并 `display()` 图形对象(设置 `surfaces_list`),
  再调用 `screenshot`。
- **每次调用重新附着**:CLI 每条命令新建 gRPC 连接(约 2–5 秒),无 MCP 的进程内缓存;
  `list_named_objects`/`find_api` 等的缓存行为由后端内部 TTL 决定,单次调用内仍有效。
- **compare_files 未在本机实测**:实现与 MCP 完全同源(会另起两个 headless 临时会话,
  较重),逻辑为仓库自带单元测试覆盖的纯库代码。
- **会话漂移**:若 Fluent 被手动关闭,会话文件仍在,后续命令报退出码 3 并提示重新
  `connect`;`connect` 探测到死会话会自动重新启动。
- **运行产物**:Fluent 会在**当前工作目录**写 `fluent-*.trn` 转录与 `cleanup-fluent-*.bat`;
  skill 收尾阶段应按原流程清理这些临时文件。
