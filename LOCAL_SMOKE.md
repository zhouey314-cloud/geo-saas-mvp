# localhost 验收补强 · 2026-10-23 北京时间 00:18

## 结果

真实 Streamlit 启动、资产大盘和生产车间浏览器 smoke 均通过。原提交 `869ee8a491fc572d430f4f03134a41b991fbfb1a` 保留为当前修复的父提交，不做重写。最终源码与实际启动副本的 SHA-256 完全相同：`7bd231e7a556fdf7d49650180f822adb6a4808efd6024940e27b5fcd364c56d6`。

现有环境：macOS，`/opt/homebrew/bin/python3` 3.13.3、Streamlit 1.58.0、OpenAI SDK 2.38.0、python-docx 1.2.0、pdfplumber 0.11.10；浏览器为现有 Google Chrome，由已安装的 agent-browser 技能操作。没有安装软件、修改全局依赖或扩展 Python 3.12 支持范围。

## 本次发现及修正

真实浏览器中，从 A 确认变量后切换 B，服务器状态已清空，但相同控件标识导致“企业主体”输入框仍带入 A 的旧值。证据为 `evidence/live-check-b-before-fix.txt`。AppTest 的原始状态操作没有模拟这个前端重放问题，补充控件标识隔离断言后先失败，见 `evidence/tests-widget-identity-before-fix.log`。

现在变量输入、文件上传、资料预览、案例编辑和文章选择控件均按工作区使用独立标识。真实浏览器完成 A→B→A：原始资料只来自当前项目，旧变量和确认状态不跟随切换。去除外部字体与图标后，再次验证最终源码 A→B，见 `evidence/live-final-state-switch.txt`。

还发现历史界面包含外部 Google Fonts 和 Icons8 图标加载；测试浏览器 allowlist 已拦截初次请求。本地源码现已移除这两处默认资源，并在 `.streamlit/config.toml` 默认关闭使用统计、绑定 `127.0.0.1`。在线模型地址配置仍保留，默认不会调用。

## 实际启动方式与边界

应用运行副本位于：

`/Users/ethanyu/Documents/Codex/2026-10-22/task-2/localhost-smoke/runtime`

只复制本次应用代码与 Streamlit 配置，添加 `SMOKE_A` / `SMOKE_B` 两个合成资料项目；A 的 30 篇文章仅用于展示已满足切片前置条件时的 UGC 禁用状态。没有使用真实客户资料。

启动命令：

```sh
env -i PATH=/opt/homebrew/bin:/usr/bin:/bin \
  HOME=/Users/ethanyu/Documents/Codex/2026-10-22/task-2/localhost-smoke/home \
  TMPDIR=/tmp PYTHONDONTWRITEBYTECODE=1 \
  /opt/homebrew/bin/python3 \
  /Users/ethanyu/Documents/Codex/2026-10-22/task-2/localhost-smoke/launch.py
```

该本地验证入口调用真实 Streamlit CLI，使用：

```text
streamlit run <runtime>/app.py
--server.address=127.0.0.1 --server.port=18763
--server.headless=true --server.fileWatcherType=none
--browser.gatherUsageStats=false --browser.serverAddress=127.0.0.1
```

验证入口清空继承环境、使用独立 HOME，在连接层拒绝非本机连接和 DNS，在 SDK 创建层禁止真实模型客户端。应用文件没有为了 smoke 添加测试分支。浏览器使用独立临时环境、独立会话、现有 Chrome，并限定页面网络为 `127.0.0.1,localhost`。

## 命令及证据

| 验证 | 实际结果 | 证据 |
| --- | --- | --- |
| `curl --noproxy '*' --max-time 5 -fsS http://127.0.0.1:18763/_stcore/health` | `ok` | `evidence/live-health.txt` |
| `lsof -nP -iTCP:18763 -sTCP:LISTEN` | 只监听 `127.0.0.1:18763` | `evidence/live-listener.txt` |
| 真实浏览器双页面渲染 | 无页面异常，默认模拟模式 | `evidence/live-final-dashboard.png`、`evidence/live-final-assertions.txt` |
| 缺少 UGC 提示词且存在 30 篇合成切片 | 显示缺失警告，按钮 `disabled=true`，没有崩溃 | `evidence/live-final-ugc-disabled.png`、`evidence/live-final-production-controls.txt` |
| 工作区切换 | A→B→A 双向通过；最终源码 A→B 再验通过 | `evidence/live-check-a-return-after-fix.txt`、`evidence/live-final-state-switch.txt` |
| 默认页面网络 | 50 个 HTTP 请求全部为本机；非本机请求 0，浏览器错误 0，服务端连接/模型拦截事件 0 | `evidence/live-network-summary.json`、`evidence/live-final-network.txt` |
| `python3 scripts/audit_offline_source.py` | 硬编码凭据候选 0、凭据模式命中 0、默认外部资源候选 0；只输出数量和路径 | `evidence/credential-scan.json` |
| `sh scripts/test_offline.sh` | 16 项通过，5.798 秒 | `evidence/tests-python313-after-live-smoke.log` |
| `python3 -m py_compile app.py tests/test_app.py scripts/audit_offline_source.py` | 退出码 0 | 本地命令完成 |
| `git diff --check` | 退出码 0 | 本地命令完成 |
| 关闭测试会话与服务后检查端口 | 已关闭；18763 无监听 | `evidence/live-listener-after-stop.txt` |

凭据扫描仅读取当前应用源码和本地配置范围，不读取密钥文件或历史密钥值。静态扫描是有限规则检查，运行时的默认离线行为另由真实浏览器网络记录与模型客户端禁用保护交叉验证。

## 未测范围与已知限制

- `ugc_160_prompts*.txt` 仍缺少，未编造或从其他工程寻找模板；真实 160 篇 UGC 生成未执行。
- 未请求真实模型服务，未验证文章质量、费用、并发长任务或 200 篇端到端生产；未执行真实文件上传全流程。
- 服务日志有 30 次历史文章复选框空标签的可访问性提示，非运行异常，见 `evidence/live-server.log`；本次未扩展界面重设计。
- 原工程仍在 `24f682cb` 且工作区干净。所有服务和浏览器测试会话已经关闭。没有推送、PR、部署、权限变更、读取客户数据或密钥，也没有任何未解决的审批拒绝。
