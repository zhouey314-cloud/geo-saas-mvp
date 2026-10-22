# 本地修复记录 · 2026-10-23 北京时间 00:03

## 版本与隔离

- 目标仅为个人公开仓库 `zhouey314-cloud/geo-saas-mvp`。
- 只读远端核验：2026-10-22 15:54 UTC，`HEAD` / `main` 均为 `24f682cb3791d115bab653a89f4b2aa59a58f9b0`；远端仅发现 main 分支。
- 用户原工程：`/Users/ethanyu/Documents/OpenSource/geo-saas-mvp`，本轮前后均在上述提交，工作区干净。
- 隔离工程：`/Users/ethanyu/Documents/Codex/2026-10-22/task-2/geo-saas-mvp-repair`。
- 本地分支：`fix/local-startup-validation`。使用 `git clone --no-hardlinks`，没有修改原工程的 Git 历史。
- 已检查相关上级目录、仓库内的 `AGENTS.md`、`.agents` / `.codex` 和项目技能文件，未发现适用项目指令。
- 工作前检查了匹配的进程名、工作目录和现有副本，未发现该仓库的另一轮活动修复。没有读取进程参数或环境密钥。

## 已复现与修复

1. 原版 `app.py` 为 49,968 字节，793 行，`python3 -m py_compile app.py` 在第 793 行报 `SyntaxError: unterminated string literal`。见 `evidence/baseline-compile.log`。
2. 取回公开 Git 历史后，直接父提交 `dd06cd005cad47baa43341e25be6e73610af0e49` 的 2,363 行应用完整可解析。保留最新提交第 1–792 行，从父提交第 792 行起恢复缺失后半段；未恢复旧版硬编码凭据。来源与文件哈希见 `evidence/recovery.json`。
3. 恢复资产大盘、生产车间、模拟生成、导出、扫描与后台函数。恢复后没有未定义的全局引用；未根据猜测补写业务流程。
4. 拦截空、`.`、`..`、分隔符和控制字符项目名；校验解析后的路径范围，拒绝工作区根目录、项目目录和子目录的符号链接。写文件拒绝符号链接目标；资产扫描跳过越界或符号链接文件。
5. 切换项目清空旧资料、变量确认、选中文章及相关控件状态，防止跨项目沿用。
6. 切片线程显式接收已校验基石目录，避免后续切换项目使后台任务依赖错误的全局目录。模拟内容使用显式传入的模型名称。
7. 默认离线模拟；真实模式缺少凭据返回失败，不再静默伪装为模拟成功。请求传递调用者指定的 `max_tokens`。离线模式可手动确认变量。
8. 缺少 `ugc_160_prompts*.txt` 时显示缺失提示并禁用 UGC 按钮，防止启动无法工作的任务；未编造模板。
9. 新增离线回归脚本、14 项测试、GitHub Actions 配置、数据与凭据忽略规则，并将直接依赖固定为本轮已验证版本。

## 命令与真实结果

| 命令 | 结果 |
| --- | --- |
| 原版 `python3 -m py_compile app.py` | 失败，第 793 行未闭合字符串，已记录 |
| 修复版 `python3 -m py_compile app.py tests/test_app.py` | 退出码 0 |
| 修复版 `python3.12 -m py_compile app.py tests/test_app.py` | 退出码 0 |
| `sh scripts/test_offline.sh` | Python 3.13.3，14 项测试全部通过，5.126 秒；见 `evidence/tests-python313.log` |
| `PYTHON_BIN=python3.12 sh scripts/test_offline.sh` | 失败，本机 Python 3.12 环境未安装 Streamlit；见 `evidence/tests-python312.log`，没有为此安装软件 |
| `git diff --check` | 退出码 0 |

测试通过 Streamlit AppTest 执行实际应用脚本，分别渲染两个页面；测试副本、合成事实和 30 篇模拟切片都位于自动清理的临时目录。测试启动时清空继承环境变量并使用临时 HOME，屏蔽 socket 连接和真实 OpenAI 客户端。唯一的请求参数测试使用内存 Mock，未请求任何模型服务。覆盖工作区名称、符号链接、项目切换、模拟输出、凭据缺失、token 参数、线程来源隔离、UGC 缺失提示、手动变量确认、文本读取、DOCX 导出和资产扫描。

日志含 Streamlit AppTest 的 `missing ScriptRunContext` 提示；测试结果为通过，不等同于生产服务验证。本轮没有启动监听服务，没有安装新软件。

## 剩余范围与阻塞

- **待提供原始素材**：目标为 160 篇 UGC 生成；缺少所有支持名称的 `ugc_160_prompts*.txt`。生成操作未执行，按钮已禁用。需用户提供确认过的模板，不从客户资料或其他私有工程寻找替代。
- **本地环境缺依赖**：Python 3.12 完整回归因缺 Streamlit 失败；Python 3.13 完整回归及两版本语法检查已完成。这不是审批拒绝。
- **按限制未执行**：真实 API、真实文章质量/事实准确性、费用、并发长任务与 200 篇端到端生产均未验证；恢复历史实现不代表这些范围已经通过验收。
- GitHub Actions 文件仅在本地；没有推送或触发远端 CI，不能声称线上 CI 已通过。
- 初始沙箱内进程读取受限、GitHub DNS 失败；只读进程核对和公开源码获取经环境授权后成功，随后一次 TLS 失败的历史读取重试成功。不存在未解决的审批拒绝或远端变更不确定性。

本轮没有访问公司私有 Node GEO、客户数据、密钥文件或密钥值；没有推送、创建 PR、合并、部署、公开新代码或改变权限。工作在北京时间 01:00 前完成，未为消耗额度延长任务。
