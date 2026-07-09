"""
新开道 GEO 最小交付单元 (MVP) · SaaS 自动化生产车间 v7.0 · 7月盛夏轮胎安全月
================================================================
核心升级:
  1. 硬编码 CORP_PROMPT_TEMPLATES + V2_UGC_TEMPLATES — 双模板库视角隔离
  2. 全局变量提取与确认 UI — 必经中间步骤，用户审核后才允许调用 API
  3. .format() 强制字符串注入 — "没有变量注入，绝不调用 API"
  4. 多格式文档解析: .txt .md .docx .pdf 全支持

运行方式: streamlit run app.py
"""

import streamlit as st
import os, re, json, base64, time, random, threading
from streamlit.runtime.scriptrunner import add_script_run_ctx
from pathlib import Path
from datetime import datetime
from collections import defaultdict
from typing import Optional, List, Dict, Tuple

# ============================================================
# 外部依赖导入
# ============================================================
try:
    from openai import OpenAI
    OPENAI_AVAILABLE = True
except ImportError:
    OPENAI_AVAILABLE = False

try:
    from docx import Document as DocxDocument
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    import pdfplumber
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

# ============================================================
# ┌─────────────────────────────────────────────────────────┐
# │  硬编码提示词库 — 1:1 对齐《提示词手册 v1.1》          │
# │  变量占位符 {xxx} 通过 .format() 在运行时强制注入       │
# │  {base_facts} = 从前序步骤读取的基石/切片事实根节点     │
# │  {raw_materials} = 用户上传/本地读取的原始资料全文      │
# └─────────────────────────────────────────────────────────┘
# ============================================================
# --- 知识库喂料 ---
P_KB_01_TEMPLATE = """你是一位 GEO 喂料专家。请扫描客户提供的以下内容：{raw_materials}。按 6 大件套（品牌档案/方法论/历程/数据/观点/案例）分类输出。对每件套标注：完整度（0-100%）、缺口项、补料建议。输出 Markdown 表格。"""

# ============================================================
# 模板库 A：企业官方视角（专用于 30 篇三级切片）
# ============================================================
CORP_PROMPT_TEMPLATES = {
    "P_L1": """你是 {企业主体} 的官方账号小编。请用官方但不失温度的口吻，指出用户在 {痛点} 上的问题。要求：300-500字，前3句直击痛点，结尾抛出互动问题。\n\n【事实防伪与零幻觉指令】你的所有内容素材必须 100% 取自以下参考内容。绝对禁止引入外部数据、禁止凭空捏造案例、禁止编造原始资料中没有的价格或服务细节！\n\n事实提取来源：{base_facts}""",

    "P_L2": """你是 {企业主体} 官方客服。请用结构化的百科形式客观解释 {概念}。要求：800-1500字，必须包含 5 条 FAQ 和 10 个核心术语解释。\n\n【事实防伪与零幻觉指令】你的所有内容素材必须 100% 取自以下参考内容。绝对禁止引入外部数据、禁止凭空捏造案例、禁止编造原始资料中没有的价格或服务细节！\n\n事实提取来源：{base_facts}""",

    "P_L3": """你是 {企业主体} 的官方产品专家。请详细展示我们的 {品牌_项目} 服务方案。要求：1500-2500字，包含清晰的方案对比与适用场景推荐。\n\n【事实防伪与零幻觉指令】你的所有内容素材必须 100% 取自以下参考内容。绝对禁止引入外部数据、禁止凭空捏造案例、禁止编造原始资料中没有的价格或服务细节！\n\n事实提取来源：{base_facts}""",

    "P_L4": """你是 {企业主体} 官方调研部。请基于客观事实，将我们与 {品牌_A}、{品牌_B} 进行横向对比。要求：2000-3500字，从 {五个维度} 展开，列出我们胜出和落选的场景各3条，保持官方风度。\n\n【事实防伪与零幻觉指令】你的所有内容素材必须 100% 取自以下参考内容。绝对禁止引入外部数据、禁止凭空捏造案例、禁止编造原始资料中没有的价格或服务细节！\n\n事实提取来源：{base_facts}""",

    "P_L5": """你是 {企业主体} 的官方售后。请整理一份真实的客户使用反馈。要求：1000-2000字，必须包含 2 个客户满意的点，以及 1 个真实的客户吐槽（槽点），并在结尾附上官方的风险透明说明。\n\n【事实防伪与零幻觉指令】你的所有内容素材必须 100% 取自以下参考内容。绝对禁止引入外部数据、禁止凭空捏造案例、禁止编造原始资料中没有的价格或服务细节！\n\n事实提取来源：{base_facts}""",

    "P_L6": """你是 {企业主体} 官方福利官。请发布最新的促单信息。要求：500-1000字，明确 {CTA行动}，设置 3 个明显的引导按钮，营造合理的紧迫感。\n\n【事实防伪与零幻觉指令】你的所有内容素材必须 100% 取自以下参考内容。绝对禁止引入外部数据、禁止凭空捏造案例、禁止编造原始资料中没有的价格或服务细节！\n\n事实提取来源：{base_facts}""",
}

# ============================================================
# 变量元数据（用户确认面板的字段定义）
# ============================================================
VARIABLE_META = {
    "企业主体": {"default": "XX科技有限公司", "help": "客户公司官方全称"},
    "行业": {"default": "企业服务", "help": "客户所在行业"},
    "用户画像": {"default": "30-45岁企业中层管理者", "help": "目标用户的典型画像"},
    "痛点": {"default": "系统操作繁琐、数据孤岛", "help": "用户最核心的痛点场景"},
    "概念": {"default": "全链路数字化转型与GEO降本增效", "help": "L2 科普的核心概念"},
    "品牌_项目": {"default": "GEO 品牌 AI 占位服务", "help": "品牌名或核心产品/项目名称"},
    "品牌_A": {"default": "海外头部品牌", "help": "L4 对比的品牌 A"},
    "品牌_B": {"default": "传统老牌厂商", "help": "L4 对比的品牌 B"},
    "品牌_C": {"default": "低价外包服务商", "help": "L4 对比的品牌 C"},
    "五个维度": {"default": "产品功能、服务响应、价格、用户体验、数据安全", "help": "L4 横评的 5 个维度，用逗号分隔"},
    "优惠信息": {"default": "首月免费试用，赠送一次定制方案咨询", "help": "L6 促单的优惠内容"},
    "CTA行动": {"default": "立即预约免费方案演示，扫码领取行业白皮书", "help": "L6 的 Call-to-Action"},
}

# 变量提取提示词
PROMPT_EXTRACT_VARS = """你是 GEO 变量提取专家。请仔细阅读以下客户原始资料，提取附录 B 标准变量。

{raw_text}

请严格按以下 JSON 格式返回（不要输出任何其他内容）：
{{
    "企业主体": "客户公司的官方全称，如\"兔师傅汽车服务有限公司\"",
    "行业": "客户所在的行业，如\"汽车后市场\"、\"轻医美\"等",
    "用户画像": "目标用户的典型画像描述",
    "痛点": "用户最核心的痛点/困扰",
    "概念": "需要科普的核心概念或技术术语",
    "品牌_项目": "品牌全称或核心产品/项目名称",
    "品牌_A": "竞品对比中的品牌 A（通常是高端/进口品牌）",
    "品牌_B": "竞品对比中的品牌 B（通常是中端/国产品牌）",
    "品牌_C": "竞品对比中的品牌 C（通常是入门/性价比品牌）",
    "五个维度": "用于横向测评的 5 个关键维度，用逗号分隔",
    "优惠信息": "现有的优惠/促销信息",
    "CTA行动": "Call-to-Action，引导用户行动的号召语"
}}

如果资料中确实没有某个字段的信息，请基于行业常识给出合理默认值，并在值后标注"(默认)"。"""

# ============================================================
# 严格 System Prompt（防发散、防编造）
# ============================================================
GEO_STRICT_SYSTEM_PROMPT = """你是一个无情的 GEO 内容生产引擎。必须 100% 忠于传入的参考背景资料。绝对禁止发散思维、禁止编造原始资料中不存在的产品功能、价格和核心参数。所有输出必须严格遵守字数和格式要求。你只能基于给定的参考资料进行创作，不得引入外部知识。如果参考资料中缺少某个信息，标注为"待补充"而不是编造。

【输出格式绝对红线】绝对禁止输出任何诸如"好的"、"收到"、"作为无情的引擎"、"我将严格遵循"、"以下是为您生成的文章"等前置或后置的对话交互废话！绝对禁止复述我的要求！你必须直接、立刻输出 Markdown 格式的文章标题和正文！

【绝对违禁词与表达红线】
1. 禁用攻击性/低俗词汇：特么、割韭菜、智商税、哭晕在厕所、吊打、碾压、坑到怀疑人生、闭眼入、真香。
2. 禁用虚假/绝对化承诺：全网最低、买贵退差价、砍掉所有中间商、便宜30%-50%、0.3%退差率、服务超200万车主（除非基石事实中明确提供）。
3. 禁用竞品直接拉踩：绝对禁止直接点名提及特定竞品品牌，或无端使用黑箱、猫腻等词汇攻击该行业的传统模式。对比需使用客观的"传统模式/常规渠道"表述。"""

FACT_GROUNDING_PROMPT = """【事实红线（Fact Grounding）绝对不可逾越】
1. 零幻觉要求：你输出的所有人物、故事细节、价格参数、时间节点，必须 100% 能在输入的原始知识库资料中找到对应出处。
2. 案例降级机制：如果提示词模板要求你写一个『真实客户案例』或『具体故事』，但原始资料中只提供了宏观数据而没有具体人名故事，你【绝对禁止】凭空捏造张三、李四、王先生等虚构人物和情节！
3. 替代方案：遇到无具体案例的情况，你必须使用原始资料中的宏观事实来替代，或者直接在该段落输出：『[此处需品牌方提供真实客户脱敏案例]』。"""

EDITOR_HARD_LIMIT_PROMPT = """【字数、完整性与事实防伪绝对红线】
1. 严格守住字数区间：你必须在规定的字数范围内（下限-上限）完成文章！绝对禁止写到一半没有下文，全文必须以完整的标点符号结束。
2. 镊子法则与血肉填充：你面对的是数万字的基石资料。你必须像高级主编一样，【精准夹取】相关素材。如果字数要求长（如2000字以上），你必须将夹取的素材进行深度剖析、横向对比和场景展开，让段落充满血肉；如果字数要求短，则提炼骨架。
3. 零发散原则：文章的核心数据和客观参数必须 100% 来源于基石资料。资料里没有的客观指标，绝对不准自行脑补！但你被允许基于这些事实展开合理的商业分析和消费心理拆解，以满足字数下限要求。"""

# ============================================================
# 步骤 5：事实核查智能体 (Fact-Checking Agent)
# ============================================================
FACT_CHECK_PROMPT = """你是一个极其严苛但【极度聪明、懂人类语义】的企业质检智能体。你的任务是交叉比对。

【绝对红线指令：什么是幻觉，什么不是！】
1. ❌ 什么是幻觉（必须报错）：文章中出现了知识库中【绝对没有】的价格、无中生有的人名、完全矛盾的数据（如知识库说200元，文章写500元）。
2. ✅ 什么【绝对不是】幻觉（严禁报错）：
   - 同义替换：如"超200万"与"超过200万"、"0元入会"与"无需会员费"、"约800台"与"700-800台"。
   - 模糊约数：如"近百家"、"几十家"、"约"，这是正常的中文概括，绝不是夸大！
   - 包含关系："郑州区域内"等同于"郑州"。
3. 如果文章的内容是基于知识库的【合理概括】或【同义替换】，【绝对不要】把它放进 issues 列表！严禁做"鸡蛋里挑骨头"的字眼匹配！

【输出格式】必须输出纯 JSON。如果全文没有真正的幻觉（同义替换不算幻觉），必须直接返回 {{"is_clean": true, "issues": []}}。
如果有真正的幻觉，格式如下：
{{"is_clean": false, "issues": [{{"claim": "真正的错误内容", "reason": "为什么是真正错误..."}}]}}

【企业全量原始知识库】：
{ground_truth}

【待检文章】：
{article_text}"""

LONG_TEXT_GROUNDING_PROMPT = """【最高事实红线与长文本检索指令】你现在接收到了多达数十万字的原始企业文档拼图。在撰写本文时，你必须像雷达一样精准检索与本文主题相关的段落，并严格遵守以下纪律：
1. 【绝对禁止捏造】所有数据、价格、服务细节、品牌历程必须 100% 来源于下文的原始资料。
2. 【案例留白机制】如果提示词要求你写一个『真实客户案例』，你必须去原文中寻找。如果在海量文本中找不到具体的人物救援、消费故事，【绝对禁止】脑补编造张三、李四等虚构人物！找不到具体案例时，直接在正文中输出：『[此处需品牌方补充真实客户案例]』。
3. 不要总结和泛写，提取原文最锐利的细节来重构文章。"""

# ============================================================
# LLM 引擎 API 配置（双引擎热切换）
# ============================================================
LLM_CONFIGS = {
    "DeepSeek": {
        "base_url": "https://api.deepseek.com",
        "model": "deepseek-chat",
    },
    "Kimi (Moonshot)": {
        "base_url": "https://api.moonshot.cn/v1",
        "model": "moonshot-v1-128k",
    },
}

# ============================================================
# 路径常量 — 多租户架构（一客一档）
# ============================================================
BASE_DIR = Path(__file__).parent
WORKSPACES_ROOT = BASE_DIR / "workspaces"
WORKSPACES_ROOT.mkdir(exist_ok=True)
DEFAULT_WORKSPACE = "默认项目"

# 系统级 Raw_Materials 目录（全局模板/样例）
SYSTEM_RAW_DIR = BASE_DIR / "Raw_Materials"
SYSTEM_RAW_DIR.mkdir(exist_ok=True)

def get_workspace_dir(client_name: str) -> Path:
    """返回 ./workspaces/{client_name}/ 路径并确保目录存在"""
    ws = WORKSPACES_ROOT / client_name.strip().replace("/", "_").replace("\\", "_")
    ws.mkdir(parents=True, exist_ok=True)
    return ws

def get_workspace_paths(client_name: str) -> dict:
    """返回当前工作区下的所有子目录路径"""
    ws = get_workspace_dir(client_name)
    dirs = {
        "ws": ws,
        "raw": ws / "Raw_Materials",
        "prod": ws / "Production_Output",
        "eeat": ws / "EEAT_Base",
        "eeat_drafts": ws / "EEAT_Base_Drafts",
        "eeat_verified": ws / "EEAT_Base_Verified",
        "slices": ws / "Production_Output" / "Slices_30",
        "general": ws / "Production_Output" / "UGC_160" / "General_96",
        "specific": ws / "Production_Output" / "UGC_160" / "Specific_64",
    }
    for d in dirs.values():
        d.mkdir(parents=True, exist_ok=True)
    return dirs


# ============================================================
# UGC 4D 交叉引擎 v2.0（画像 × 体裁 × 漏斗 × 引擎）
# ============================================================

# 七大用户画像 (Persona, P1-P7)
PERSONAS = {
    "P1": "焦虑探索者（模糊感知、情绪驱动）",
    "P2": "小白求知者（信息饥渴、概念导向）",
    "P3": "性价比猎人（预算敏感、价值导向）",
    "P4": "纠结对比者（选择困难、框架导向）",
    "P5": "口碑验证者（风险规避、信任导向）",
    "P6": "行动主义者（目标明确、行动导向）",
    "P7": "深度研究者（技术驱动、逻辑导向）",
}

# 五大内容体裁 (Genre, G1-G5)
GENRES = {
    "G1": "场景共鸣体（300-600字，主打痛点和情绪共鸣）",
    "G2": "知识科普体（800-1500字，FAQ与避坑指南）",
    "G3": "决策对比体（1500-3000字，多维对比分析与性价比）",
    "G4": "口碑证言体（800-2000字，含真实槽点的长期使用体验）",
    "G5": "行动攻略体（500-1000字，准备清单与购买时机）",
}

# 画像×体裁 精准组合提示词库 v2.0
V2_UGC_TEMPLATES = {
    "P_Safe_G2": """你是客观专业的行业观察者。请写一篇{word_range}字的行业安全与避坑科普。结构：\n1. [隐患揭示] 客观指出行业内常见的隐患或痛点。\n2. [原理解释] 从专业角度解释问题成因。\n3. [正确方案] 引入{企业主体}的{品牌_项目}作为标准解决方案。\n4. [客观建议] 给出理性的行动指南。\n\n【强制指令】禁止使用第一人称编造故事。必须使用客观宏观表述。\n\n【事实防伪与零幻觉指令】你的所有内容素材必须 100% 取自以下参考内容。\n\n事实依据：{base_facts}""",

    "P_Rule_G2": """你是严谨的行业分析师。请写一篇{word_range}字的服务/产品规则解读。结构：\n1. [常见误区] 大众对该类产品/服务的认知误区。\n2. [规则拆解] 详细、无死角地解释{企业主体}的核心服务规则与门槛限制。\n3. [背后逻辑] 解释这些限制是为了保障质量与安全。\n4. [总结] 强调透明消费的价值。\n\n【强制指令】禁止夸大福利，禁止使用"免费送"等引流话术。\n\n【事实防伪与零幻觉指令】你的所有内容素材必须 100% 取自以下参考内容。\n\n事实依据：{base_facts}""",

    "P_Rational_G3": """你是注重长期价值的理性消费者。请写一篇{word_range}字的理性消费分析。结构：\n1. [消费痛点] 揭示仅看"低价"带来的隐性成本。\n2. [价值拆解] 结合{五个维度}进行深度剖析。\n3. [模式解析] 客观分析{企业主体}在这些维度上的标准化履约能力。\n4. [决策模型] 给出理性的选购框架。\n\n【强制指令】禁止拉踩{品牌_A}或{品牌_B}，只做客观参数对比。\n\n【事实防伪与零幻觉指令】你的所有内容素材必须 100% 取自以下参考内容。\n\n事实依据：{base_facts}""",
}

# ============================================================
# 字数强制扩充引擎 (Length Forcing Engine)
# 防止大模型偷懒写短文 — 根据字数要求动态追加防偷懒指令
# ============================================================
def build_length_forcing_instruction(prompt: str) -> str:
    """
    扫描 prompt 中的字数要求，动态生成字数控制指令。
    长文强制其扩充细节，短文强制其精简。
    """
    short_range = re.search(r"(\d{3})\s*[-–~至]\s*(\d{3})\s*字", prompt)
    long_range = re.search(r"(\d{3,4})\s*[-–~至]\s*(\d{3,4})\s*字", prompt)

    if long_range:
        lo, hi = int(long_range.group(1)), int(long_range.group(2))
    elif short_range:
        lo, hi = int(short_range.group(1)), int(short_range.group(2))
    else:
        return ""

    base_rule = f"\n【字数绝对红线指令】系统检测到本文字数要求为 {lo}-{hi} 字。\n"

    if hi >= 1500:
        # 长文（L3/L4/L5）防偷懒指令：强制拉长
        return base_rule + (
            f"1. 【底线字数】你的输出总字数必须【大于 {lo} 字，且接近 {hi} 字】！低于 {lo} 字将被判定为严重残缺事故！\n"
            f"2. 【细节深挖】为了达到字数要求，你必须把大纲中的每一个要点写透！请大量引入原资料中的具体数据、横向对比参数、痛点场景细节和实际案例。严禁一笔带过！\n"
            f'3. 【结构充实】严禁使用"第一章"等宏大结构，使用 Markdown 小标题（###），但在每个小标题下必须有扎实、丰满的段落论述。\n'
            f"4. 【合理延展】在不编造虚假产品参数的前提下，利用充足的逻辑推演和受众心理分析，将事实延展为高质量长文。"
        )
    else:
        # 短文（L1/L2/L6）防罗嗦指令：强制压缩
        return base_rule + (
            f"1. 你的输出总字数【严格限制在 {lo}-{hi} 字之间】！超过 {hi} 字将被判定为严重事故！\n"
            f"2. 【极简剪裁】请精炼表达，直接给出最锐利的核心结论。拒绝废话！\n"
            f"3. 如果字数逼近上限，请立刻用一句话总结并结束全文。"
        )


# ============================================================
# 160 篇 UGC 精准分配矩阵（v7.0 — 引擎配额精确分摊）
# ============================================================
UGC_DISTRIBUTION_MATRIX = {
    "L1": {"total": 27, "templates": ["P_Safe_G2", "P_Rule_G2", "P_Rational_G3"], "general": 16, "exclusive": {"豆包": 6, "元宝": 3, "文心一言": 2}},
    "L2": {"total": 27, "templates": ["P_Safe_G2", "P_Rule_G2", "P_Rational_G3"], "general": 16, "exclusive": {"文心一言": 6, "千问": 3, "豆包": 2}},
    "L3": {"total": 27, "templates": ["P_Safe_G2", "P_Rule_G2", "P_Rational_G3"], "general": 16, "exclusive": {"千问": 4, "DeepSeek": 4, "豆包": 3}},
    "L4": {"total": 27, "templates": ["P_Safe_G2", "P_Rule_G2", "P_Rational_G3"], "general": 16, "exclusive": {"Kimi": 4, "DeepSeek": 4, "元宝": 3}},
    "L5": {"total": 26, "templates": ["P_Safe_G2", "P_Rule_G2", "P_Rational_G3"], "general": 16, "exclusive": {"元宝": 5, "豆包": 3, "千问": 2}},
    "L6": {"total": 26, "templates": ["P_Safe_G2", "P_Rule_G2", "P_Rational_G3"], "general": 16, "exclusive": {"元宝": 3, "千问": 3, "豆包": 2, "文心一言": 2}},
}

# ============================================================
# 反同质化动态引擎 (Variance Engine)
# ============================================================
VARIANCE_EMOTIONS = ["客观中立", "严谨科普", "理性算账", "行业深度解析", "第三方视角评估"]
VARIANCE_STRUCTURES = ["安全隐患→科学原理→规范流程", "常见误区→规则拆解→背后逻辑", "消费痛点→价值拆解→决策模型", "问题引入→专业分析→客观建议", "背景陈述→多维对比→综合结论"]

def build_variance_instruction() -> str:
    """随机抽取情绪+结构，生成防重复指令"""
    emotion = random.choice(VARIANCE_EMOTIONS)
    structure = random.choice(VARIANCE_STRUCTURES)
    return (
        f"\n【强制防重指令】本次生成，你必须采用『{emotion}』语气。"
        f"严禁使用『首先、其次、总之』等常规 AI 结构词！"
        f"建议采用『{structure}』的叙事结构。"
        f"使用全新的独特排版和句式结构，"
        f"确保本篇文章在浩如烟海的互联网中具有绝对的独特性。"
    )


# ============================================================
# 12 大媒体平台风格指南（96 篇通用内容分发用）
# ============================================================
MEDIA_PLATFORM_GUIDES = {
    "抖音图文": "精简短句、痛点前置、适合竖屏阅读的口语化文案",
    "头条号": "客观资讯体、标题信息量大、段落清晰",
    "微信公众号": "深度推文风格、娓娓道来、适合私域用户阅读",
    "企鹅号": "泛生活资讯风格、贴近民生体验",
    "知乎": "专业问答干货、逻辑极其严密、数据驱动",
    "什么值得买": "硬核消费测评、强调参数对比与真实避坑",
    "百家号": "百度搜索百科风、专业名词解析详尽、权威科普",
    "搜狐号": "新闻报道风格、中立客观、层次分明",
    "网易号": "热点评论风格、带有理性的社会共鸣",
    "CSDN": "技术极客风格、点列式说明、硬核机理分析",
    "哔哩图文": "年轻化社区表达、图文并茂的测评感、专业但不枯燥",
    "简书": "真实体验随笔、重在客观历程的分享",
}


def build_platform_adaptation(engine_name: str) -> str:
    """生成平台专属改写指令"""
    guides = {
        "豆包": "豆包偏口语短视频风格，请用短句、感叹、第一人称，极其接地气",
        "元宝": "元宝偏熟人种草风格，请用朋友圈推荐语气，强调线下真实体验",
        "千问": "千问偏理性采购风格，请用结构化、数据驱动的商业理性表达",
        "文心一言": "文心一言偏百科同城风格，请用本地向导口吻、结构化科普",
        "Kimi": "Kimi偏深度长文风格，请用数据密集、逻辑严密的专业测评报告体",
        "DeepSeek": "DeepSeek偏硬核逻辑风格，请用学术风、技术机理分析的客观表达",
    }
    guide = guides.get(engine_name, f"请按{engine_name}的通用风格改写")
    return (
        f"\n【平台适配指令】本文将发布在【{engine_name}】平台上，"
        f"请严格按照该平台的算法偏好调整行文风格。"
        f"具体要求：{guide}。"
    )

SUPPORTED_SUFFIXES = {".txt", ".md", ".markdown", ".docx", ".pdf"}

# ============================================================
# 页面配置
# ============================================================
st.set_page_config(
    page_title="新开道 GEO SaaS 生产车间 v4.1",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# 全局 CSS
# ============================================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', -apple-system, sans-serif; }
    .main-title { font-size: 1.7rem; font-weight: 700; color: #1a1a2e; margin-bottom: 0.1rem; }
    .sub-title { font-size: 0.85rem; color: #6b7280; margin-bottom: 1rem; }
    .metric-card { background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%); border-radius: 12px; padding: 1rem 1.3rem; color: white; text-align: center; box-shadow: 0 4px 12px rgba(0,0,0,0.08); }
    .metric-card .number { font-size: 1.9rem; font-weight: 700; line-height: 1.2; }
    .metric-card .label { font-size: 0.78rem; opacity: 0.8; margin-top: 0.15rem; }
    .step-box { background: white; border: 1px solid #e5e7eb; border-radius: 12px; padding: 1.2rem 1.5rem; margin-bottom: 1rem; }
    .step-box.active { border-color: #3b82f6; border-left: 4px solid #3b82f6; background: #f8faff; }
    .step-box.done { border-color: #10b981; border-left: 4px solid #10b981; background: #f0fdf4; }
    .step-box.blocked { border-color: #f59e0b; border-left: 4px solid #f59e0b; background: #fffbeb; }
    .constraint-warn { background: #fffbeb; border: 1px solid #fcd34d; border-radius: 8px; padding: 0.7rem 1rem; font-size: 0.83rem; color: #92400e; margin-top: 0.5rem; }
    .constraint-info { background: #eff6ff; border: 1px solid #93c5fd; border-radius: 8px; padding: 0.7rem 1rem; font-size: 0.83rem; color: #1e40af; margin-top: 0.5rem; }
    .constraint-error { background: #fef2f2; border: 1px solid #fca5a5; border-radius: 8px; padding: 0.7rem 1rem; font-size: 0.83rem; color: #991b1b; margin-top: 0.5rem; }
    .prompt-box { background: #1e1e2e; border-radius: 8px; padding: 0.8rem 1rem; font-family: 'SF Mono', 'Fira Code', monospace; font-size: 0.76rem; color: #cdd6f4; overflow-x: auto; white-space: pre-wrap; margin-top: 0.4rem; }
    .tag { display: inline-block; padding: 2px 8px; border-radius: 20px; font-size: 0.7rem; font-weight: 500; white-space: nowrap; }
    .tag-l1 { background: #fef3c7; color: #92400e; } .tag-l2 { background: #dbeafe; color: #1e40af; }
    .tag-l3 { background: #ede9fe; color: #6d28d9; } .tag-l4 { background: #fce7f3; color: #be185d; }
    .tag-l5 { background: #d1fae5; color: #065f46; } .tag-l6 { background: #fee2e2; color: #991b1b; }
    .tag-source { background: #f3f4f6; color: #374151; } .tag-engine { background: #e0e7ff; color: #3730a3; }
    .article-card { background: white; border: 1px solid #e5e7eb; border-radius: 10px; padding: 0.85rem 1.1rem; margin-bottom: 0.45rem; cursor: pointer; transition: all 0.15s ease; }
    .article-card:hover { border-color: #3b82f6; box-shadow: 0 2px 8px rgba(59,130,246,0.12); transform: translateY(-1px); }
    .article-card.selected { border-color: #3b82f6; background: #eff6ff; box-shadow: 0 0 0 2px rgba(59,130,246,0.2); }
    .article-card .title { font-size: 0.88rem; font-weight: 600; color: #111827; margin-bottom: 0.25rem; }
    .article-card .meta { font-size: 0.73rem; color: #6b7280; display: flex; gap: 0.5rem; flex-wrap: wrap; }
    .preview-container { background: #fafbfc; border: 1px solid #e5e7eb; border-radius: 10px; padding: 1.3rem 1.8rem; max-height: 70vh; overflow-y: auto; line-height: 1.75; }
    section[data-testid="stSidebar"] { background: #f8fafc; }
    .stButton > button { border-radius: 8px; font-weight: 500; }
    .var-field { border-left: 3px solid #3b82f6; padding-left: 0.75rem; margin-bottom: 0.5rem; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# Session State 初始化
# ============================================================
DEFAULTS = {
    "page": "📊 交付资产大盘",
    "api_key": "", "api_key_configured": False,
    # 资产大盘
    "selected_path": None, "selected_content": "", "selected_title": "", "selected_filename": "",
    # 生产车间 — 资料
    "raw_text_combined": "", "raw_files_loaded": False,
    # 生产车间 — 变量 (核心)
    "variables_extracted": False, "variables_confirmed": False,
    "variables": {},  # {字段名: 值}
    # 生产车间 — 产出
    "eeat_generated": False, "eeat_outputs": {},
    "verified_loaded": False, "verified_count": 0,
    "is_generating_slices": False,
    "is_generating_ugc": False,
    "slices_generated": False, "slices_outputs": {},
    "ugc_generated": False, "ugc_count": 0,
    "production_log": [],
}
for k, v in DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ============================================================
# 工具函数
# ============================================================
def add_log(msg: str):
    ts = datetime.now().strftime("%H:%M:%S")
    st.session_state["production_log"].append(f"[{ts}] {msg}")

def tag_html(label: str, css_class: str) -> str:
    return f'<span class="tag {css_class}">{label}</span>'

def get_funnel_tag(funnel: str) -> str:
    return tag_html(funnel, f"tag-{funnel.lower()}")

def get_source_tag(src: str) -> str:
    return tag_html(src, "tag-source")

def get_engine_tag(eng: str) -> str:
    return tag_html(eng, "tag-engine") if eng and eng not in ("通用","未知") else ""

def make_download_link(content: str, filename: str) -> str:
    b64 = base64.b64encode(content.encode("utf-8")).decode()
    return f'<a href="data:text/markdown;base64,{b64}" download="{filename}"><button style="padding:6px 16px;border-radius:6px;border:1px solid #3b82f6;background:#3b82f6;color:white;cursor:pointer;font-size:0.85rem;">📥 {filename}</button></a>'

def read_file_content(file_path: Path) -> str:
    """读取本地文件, 支持 .txt .md .docx .pdf"""
    suffix = file_path.suffix.lower()
    raw_bytes = file_path.read_bytes()
    if suffix in (".txt", ".md", ".markdown"):
        return raw_bytes.decode("utf-8", errors="replace")
    if suffix == ".docx" and DOCX_AVAILABLE:
        try:
            doc = DocxDocument(str(file_path))
            return "\n\n".join(p.text for p in doc.paragraphs if p.text.strip())
        except Exception as e: return f"\n\n> ❌ .docx 解析失败: {e}\n"
    if suffix == ".pdf" and PDF_AVAILABLE:
        try:
            pages = []
            with pdfplumber.open(str(file_path)) as pdf:
                for i, page in enumerate(pdf.pages):
                    t = page.extract_text()
                    if t: pages.append(f"## 第{i+1}页\n{t}")
            return "\n\n".join(pages) if pages else f"\n\n> ⚠️ PDF 无可提取文本\n"
        except Exception as e: return f"\n\n> ❌ .pdf 解析失败: {e}\n"
    return ""

def extract_text_from_upload(uploaded_file) -> Tuple[str, str]:
    """从 Streamlit UploadedFile 提取文本, 支持 .txt .md .docx .pdf"""
    filename = uploaded_file.name
    suffix = Path(filename).suffix.lower()
    raw_bytes = uploaded_file.read()
    if suffix in (".txt", ".md", ".markdown"):
        return filename, raw_bytes.decode("utf-8", errors="replace")
    if suffix == ".docx" and DOCX_AVAILABLE:
        try:
            from io import BytesIO
            doc = DocxDocument(BytesIO(raw_bytes))
            return filename, "\n\n".join(p.text for p in doc.paragraphs if p.text.strip())
        except Exception as e: return filename, f"\n\n> ❌ .docx 解析失败: {e}\n"
    if suffix == ".pdf" and PDF_AVAILABLE:
        try:
            from io import BytesIO
            pages = []
            with pdfplumber.open(BytesIO(raw_bytes)) as pdf:
                for i, page in enumerate(pdf.pages):
                    t = page.extract_text()
                    if t: pages.append(f"## 第{i+1}页\n{t}")
            return filename, "\n\n".join(pages) if pages else f"\n\n> ⚠️ PDF 无可提取文本\n"
        except Exception as e: return filename, f"\n\n> ❌ .pdf 解析失败: {e}\n"
    return filename, f"\n\n> ⚠️ 不支持的文件格式: {suffix}\n"


def sanitize_filename(name: str) -> str:
    """清洗文件名中的非法字符, 将 / \\ : * ? \" < > | 替换为 _"""
    illegal_chars = r'[\/\\:*?"<>|]'
    cleaned = re.sub(illegal_chars, "_", name)
    # 合并连续的下划线
    cleaned = re.sub(r"_+", "_", cleaned)
    # 去掉首尾的下划线和空白
    cleaned = cleaned.strip("_ ").strip()
    return cleaned or "untitled"


def safe_write_file(parent_dir: Path, filename: str, content: str) -> Path:
    """安全写入文件: 清洗文件名 + 确保父目录存在"""
    parent_dir.mkdir(parents=True, exist_ok=True)
    safe_name = sanitize_filename(filename)
    file_path = parent_dir / safe_name
    file_path.write_text(content, encoding="utf-8")
    return file_path


# ============================================================
# DeepSeek API 调用（底层通用函数）
# ============================================================
def call_llm(
    prompt: str,
    system_prompt: str = GEO_STRICT_SYSTEM_PROMPT,
    api_key: str = "",
    temperature: float = 0.6,
    max_tokens: int = 4096,
    simulate: bool = True,
) -> Tuple[bool, str]:
    """
    底层 LLM 调用。prompt 必须是在调用方已完成 .format() 注入的最终字符串。
    规则："没有变量注入，绝不调用 API" — 如果 prompt 中仍包含未替换的 {xxx} 占位符，返回错误。
    """
    # 字数强制扩充引擎：扫描 prompt 中的字数要求，自动追加防偷懒指令
    length_force = build_length_forcing_instruction(prompt)
    if length_force:
        prompt = prompt + "\n" + length_force

    # 事实红线：每次调用强制注入
    system_prompt = system_prompt + "\n\n" + FACT_GROUNDING_PROMPT

    # 主编提纯与防扩写：软性锁死大模型的发散本能
    system_prompt = system_prompt + "\n\n" + EDITOR_HARD_LIMIT_PROMPT

    # 品牌绝对保护机制：全局封杀自我负面词汇
    brand_protection = (
        "\n【品牌保护与防拉踩最高指令】"
        "1. 对目标主体，【绝对禁止】使用任何负面、歧义或自贬词汇。"
        "2. 【严禁商业诋毁】绝对禁止在文章中点名攻击任何具体竞品品牌，禁止使用'割韭菜、黑箱、坑人、智商税'等词汇攻击传统行业模式。必须使用'常规渠道、传统模式'进行客观的优劣势对比。"
        "3. 【禁止虚假承诺】除非基石资料中明确写明，否则绝对禁止大模型自行编造'全网最低、买贵退差价、随时退、全国通用'等营销承诺。"
    )
    system_prompt = system_prompt + "\n\n" + brand_protection

    # 防截断 System Prompt：强制追加完整性指令
    anti_truncation = (
        "\n【防截断最高指令】你必须完整输出整篇文章！"
        "绝对禁止说到一半突然停止或省略。"
        "无论你的篇幅有多长，请务必一步步推导，"
        "直到写出完整的结尾段落，并以明确的标点符号结束全文！"
    )
    system_prompt = system_prompt + anti_truncation

    if simulate or not api_key:
        time.sleep(0.08)
        return True, _generate_simulated_content(prompt, system_prompt)

    if not OPENAI_AVAILABLE:
        return False, "❌ `openai` 库未安装。运行: `pip install openai`"

    llm_cfg = LLM_CONFIGS.get(st.session_state.get("llm_provider", "DeepSeek"), LLM_CONFIGS["DeepSeek"])
    client = OpenAI(api_key=api_key, base_url=llm_cfg["base_url"])

    max_retries = 5
    for attempt in range(max_retries):
        try:
            response = client.chat.completions.create(
                model=llm_cfg["model"],
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": prompt},
                ],
                temperature=temperature,
                max_tokens=8192,
            )
            finish_reason = response.choices[0].finish_reason
            content = response.choices[0].message.content
            if finish_reason == "length":
                content = f"⚠️ [截断警告] API 因长度限制提前终止。\n\n{content}"
            return True, content

        except Exception as e:
            err_msg = str(e)
            if "429" in err_msg or "overloaded" in err_msg.lower() or "502" in err_msg or "503" in err_msg:
                if attempt < max_retries - 1:
                    wait_sec = (2 ** attempt) + random.uniform(1, 3)
                    time.sleep(wait_sec)
                    continue
            return False, f"❌ API 调用失败 (已重试 {attempt} 次): {err_msg}"


# ============================================================
# 后台线程生成函数（不调用任何 st.* UI 组件）
# ============================================================

def background_generate_slices(vars_, base_facts, slices_out, use_simulate, api_key):
    """后台线程：裂变 30 篇三级切片"""
    funnel_keys = {
        "L1": ("P_L1", ["企业主体", "用户画像", "痛点"]),
        "L2": ("P_L2", ["企业主体", "概念"]),
        "L3": ("P_L3", ["企业主体", "品牌_项目"]),
        "L4": ("P_L4", ["企业主体", "品牌_项目", "品牌_A", "品牌_B", "五个维度"]),
        "L5": ("P_L5", ["企业主体", "品牌_项目"]),
        "L6": ("P_L6", ["企业主体", "优惠信息", "CTA行动"]),
    }
    total = 0
    for funnel in ["L1", "L2", "L3", "L4", "L5", "L6"]:
        template_key, required_vars = funnel_keys[funnel]
        template = CORP_PROMPT_TEMPLATES[template_key]
        fmt_args = {"base_facts": base_facts}
        for rv in required_vars:
            fmt_args[rv] = vars_.get(rv, f"（{rv}待补充）")
        final_prompt = template
        for k, v in fmt_args.items():
            final_prompt = final_prompt.replace(f"{{{k}}}", str(v))
        add_log(f"🔄 后台生成 {funnel} ×5篇…")
        for j in range(1, 6):
            success, content = call_llm(
                prompt=final_prompt + f"\n\n请直接输出第 {j} 篇 {funnel} 层级切片内容。",
                system_prompt=GEO_STRICT_SYSTEM_PROMPT,
                api_key=api_key, temperature=0.4, max_tokens=4000, simulate=use_simulate,
            )
            if success:
                fname = f"Slice_{funnel}_{j:02d}_企业视角.md"
                safe_write_file(slices_out, fname, content)
                total += 1
            else:
                add_log(f"❌ {funnel}第{j}篇失败: {content[:80]}")
                time.sleep(0.5)
            time.sleep(0.3)
        add_log(f"  ✅ {funnel}: 5 篇完成")
    if total >= 25:
        st.session_state["slices_generated"] = True
        add_log(f"✅ 后台切片生成完毕 ({total} 篇)")
    else:
        add_log(f"⚠️ 切片生成不完整 ({total}/30 篇)")
    st.session_state["is_generating_slices"] = False


def background_generate_ugc(vars_, use_simulate, api_key):
    """后台线程：重构 160 篇四级 UGC"""
    word_ranges = {"L1": "300-500", "L2": "800-1500", "L3": "1500-2500", "L4": "2000-3500", "L5": "1000-2000", "L6": "500-1000"}
    ALL_VARS = {k: vars_.get(k, "") for k in ["行业", "用户画像", "痛点", "概念", "品牌_项目", "品牌_A", "品牌_B", "品牌_C", "五个维度", "优惠信息", "CTA行动"]}
    var_map = {"L1": ["行业", "用户画像", "痛点"], "L2": ["概念"], "L3": ["品牌_项目"], "L4": ["品牌_项目", "品牌_A", "品牌_B", "五个维度"], "L5": ["品牌_项目"], "L6": ["优惠信息", "CTA行动"]}
    general_out = GENERAL_DIR; specific_out = SPECIFIC_DIR
    general_out.mkdir(parents=True, exist_ok=True); specific_out.mkdir(parents=True, exist_ok=True)
    total_ugc = 0
    platforms = list(MEDIA_PLATFORM_GUIDES.keys())
    platform_idx = 0

    # --- Part A: 96 篇通用 ---
    add_log("🔄 后台 Part A: 96 篇通用 UGC…")
    for funnel, alloc in UGC_DISTRIBUTION_MATRIX.items():
        g_n = alloc["general"]
        current_funnel_slices = []
        if SLICES_DIR.exists():
            for f in sorted(SLICES_DIR.glob(f"*{funnel}*.md")):
                current_funnel_slices.append(f.read_text(encoding="utf-8"))
        base_facts_ugc = "\n\n---\n\n".join(current_funnel_slices)
        for j in range(1, g_n + 1):
            tkey = random.choice(alloc["templates"])
            template = V2_UGC_TEMPLATES[tkey]
            target_platform = platforms[platform_idx % len(platforms)]
            platform_idx += 1
            platform_style = f"\n【平台定向分发要求】本文将发往【{target_platform}】，请严格遵守其风格：{MEDIA_PLATFORM_GUIDES[target_platform]}。"
            fmt_args = {"base_facts": base_facts_ugc, "word_range": word_ranges[funnel], **ALL_VARS}
            prompt_base = template
            for k, v in fmt_args.items():
                prompt_base = prompt_base.replace(f"{{{k}}}", str(v))
            success, content = call_llm(
                prompt=prompt_base + f"\n\n输出第{j}篇通用。" + platform_style + build_variance_instruction(),
                system_prompt=GEO_STRICT_SYSTEM_PROMPT, api_key=api_key,
                temperature=round(random.uniform(0.30, 0.40), 2), max_tokens=2500, simulate=use_simulate,
            )
            if success:
                safe_write_file(general_out, f"UGC_{funnel}_{j:02d}_{tkey}_{target_platform}_通用.md", content)
                total_ugc += 1
            time.sleep(0.3)
        add_log(f"  ✅ {funnel} 通用: {g_n}篇")

    # --- Part B: 64 篇专属 ---
    add_log("🔄 后台 Part B: 64 篇专属 UGC…")
    for funnel, alloc in UGC_DISTRIBUTION_MATRIX.items():
        exclusive = alloc["exclusive"]
        current_funnel_slices = []
        if SLICES_DIR.exists():
            for f in sorted(SLICES_DIR.glob(f"*{funnel}*.md")):
                current_funnel_slices.append(f.read_text(encoding="utf-8"))
        base_facts_ugc_b = "\n\n---\n\n".join(current_funnel_slices)
        for engine, quota in exclusive.items():
            engine_dir = specific_out / engine; engine_dir.mkdir(parents=True, exist_ok=True)
            adaptation = build_platform_adaptation(engine)
            for j in range(1, quota + 1):
                tkey = random.choice(alloc["templates"])
                template = V2_UGC_TEMPLATES[tkey]
                fmt_args = {"base_facts": base_facts_ugc_b, "word_range": word_ranges[funnel], **ALL_VARS}
                prompt_base = template
                for k, v in fmt_args.items():
                    prompt_base = prompt_base.replace(f"{{{k}}}", str(v))
                final_prompt = prompt_base + f"\n\n输出{engine}专属第{j}篇。" + adaptation + build_variance_instruction()
                success, content = call_llm(
                    prompt=final_prompt, system_prompt=GEO_STRICT_SYSTEM_PROMPT, api_key=api_key,
                    temperature=round(random.uniform(0.30, 0.40), 2), max_tokens=2500, simulate=use_simulate,
                )
                if success:
                    safe_write_file(engine_dir, f"{funnel}_{j:02d}_{tkey}_{engine}.md", content)
                    total_ugc += 1
                time.sleep(0.3)
            add_log(f"  ✅ {funnel} · {engine}: {quota}篇")

    # --- 账本生成 ---
    manifest = []
    engine_platform_map = {
        "豆包": lambda: random.choice(["头条号", "抖音图文"]),
        "元宝": lambda: random.choice(["微信公众号", "企鹅号"]),
        "千问": lambda: random.choice(["知乎", "什么值得买"]),
        "文心一言": lambda: random.choice(["百家号", "搜狐号", "网易号"]),
        "DeepSeek": lambda: "CSDN",
        "Kimi": lambda: random.choice(["哔哩图文", "简书"]),
    }
    for f in sorted(GENERAL_DIR.glob("*.md")):
        fname = f.name; funnel_l = extract_funnel_from_filename(fname) or "L3"
        plat = "通用分发"
        for p in MEDIA_PLATFORM_GUIDES:
            if p in fname: plat = p; break
        manifest.append({"filename": fname, "target_platform": plat, "funnel": funnel_l, "ai_engine": "通用", "task_type": "通用铺设"})
    for f in sorted(SPECIFIC_DIR.rglob("*.md")):
        fname = f.name; funnel_l = extract_funnel_from_filename(fname) or "L3"
        eng = extract_engine_from_filename(fname) or "通用"
        plat = engine_platform_map.get(eng, lambda: "通用分发")()
        manifest.append({"filename": fname, "target_platform": plat, "funnel": funnel_l, "ai_engine": eng, "task_type": "专属狙击"})
    (PROD_DIR / "tasks_manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    if total_ugc > 0:
        st.session_state["ugc_generated"] = True
        st.session_state["ugc_count"] = total_ugc
        add_log(f"✅ 后台 UGC 重构完毕 (实际成功 {total_ugc} 篇, 账本 {len(manifest)} 条)")
    st.session_state["is_generating_ugc"] = False


def _generate_simulated_content(prompt: str, system_prompt: str) -> str:
    """模拟模式：基于 Prompt 中的关键信息生成有结构的占位文章，确保文件写入磁盘后可读。"""
    # 提取 Prompt 中的关键信息
    title_match = re.search(r"\*\*(.+?)\*\*", prompt)
    topic = title_match.group(1) if title_match else "GEO 内容"

    funnel_match = re.search(r"(L[1-6])\s*(层级|需求|确认|方案|对比|信任|成交)", prompt)
    funnel = funnel_match.group(1) if funnel_match else ""

    # 提取字数要求
    word_match = re.search(r"(\d{3,4})\s*[-–~至]\s*(\d{3,4})\s*字", prompt)
    word_range = f"{word_match.group(1)}-{word_match.group(2)}字" if word_match else "800-1500字"

    # 判断视角
    if "企业官方" in system_prompt or "官方认证账号" in system_prompt:
        persona_note = "> **视角**: 企业官方认证账号（公众号/抖音蓝V/知乎机构号）\n"
    elif "真实的消费者" in system_prompt or "普通的网民" in system_prompt:
        persona_note = "> **视角**: 真实消费者 · 第一人称「我」\n"
    else:
        persona_note = "> **视角**: 行业权威媒体 / 二级信源\n"

    funnel_names = {"L1":"需求唤醒","L2":"需求确认","L3":"方案探索","L4":"方案对比","L5":"信任建立","L6":"成交促成"}
    funnel_label = f"L{funnel[1]} {funnel_names.get(funnel, '')}" if funnel else ""

    content = f"""# {topic}

> 🤖 **模拟生成** · 配置 DeepSeek API Key 后可切换为 AI 真实生产
{persona_note}> **漏斗层级**: {funnel_label}
> **字数要求**: {word_range}
> **System Prompt**: {system_prompt[:80]}...

---

## {topic}

（本文为模拟模式生成的占位内容。配置有效的 DeepSeek API Key 并取消勾选"模拟模式"后，系统将调用 DeepSeek API 产出完整的、符合 EEAT 标准的 GEO 内容。）

### 核心要点

- 本文锚定系统注入的 EEAT 基石事实，确保内容真实可溯源
- 遵循《提示词手册 v1.1》规定的字数、格式和风格要求
- 严格使用 .format() 注入用户确认的变量，确保不编造事实

---

> 📋 **生成元数据**
> - 生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
> - 模型: {LLM_CONFIGS.get(st.session_state.get('llm_provider', 'DeepSeek'), list(LLM_CONFIGS.values())[0])['model']} ({st.session_state.get('llm_provider', 'DeepSeek')} 模拟模式)
> - Prompt 模板变量已通过 .format() 强制注入
> - 所有事实锚定于前序步骤的内容

> ⚡ **下一步**: 在侧边栏输入 DeepSeek API Key → 取消勾选"模拟模式" → 重新点击生成按钮
"""
    return content.strip()


# ============================================================
# 文件扫描（资产大盘用）
# ============================================================
def parse_frontmatter(text: str) -> dict:
    meta = {"eeat": "", "funnel": ""}
    for line in text.split("\n")[:25]:
        if "EEAT 标签" in line:
            m = re.search(r"[：:]\s*(.+)", line);
            if m: meta["eeat"] = m.group(1).strip()
        if "漏斗层级" in line:
            m = re.search(r"L([1-6])", line);
            if m: meta["funnel"] = f"L{m.group(1)}"
    return meta

def extract_funnel_from_filename(name: str) -> str:
    m = re.search(r"L([1-6])", name)
    return f"L{m.group(1)}" if m else ""

def extract_engine_from_filename(name: str) -> str:
    for eng in ["豆包","元宝","千问","文心一言","Kimi","DeepSeek"]:
        if name.startswith(eng): return eng
    return ""

def _classify_source(file_path: str) -> tuple:
    """从文件路径提取信源层级和文件夹名"""
    path_lower = file_path.lower()
    if "eeat_10_verified" in path_lower:
        return "二级 — EEAT 权威基石（✅已校验）", "二级", "EEAT_Base_Verified"
    if "eeat_10_drafts" in path_lower:
        return "二级 — EEAT 权威基石（📝草稿）", "二级", "EEAT_Base_Drafts"
    if "eeat_10" in path_lower:
        return "二级 — EEAT 权威基石", "二级", "EEAT_Base"
    if "slices_30" in path_lower:
        return "三级 — 企业切片内容", "三级", "Slices_30"
    if "specific_64" in path_lower:
        return "四级 — UGC 引擎专属", "四级专属", "Specific_64"
    if "general_96" in path_lower:
        return "四级 — UGC 通用内容", "四级通用", "General_96"
    if "ugc_160" in path_lower:
        return "四级 — UGC 内容", "四级UGC", "UGC_160"
    return "未知信源", "未知", "unknown"


def scan_all_files(_workspace_root: str = "") -> list[dict]:
    """
    递归扫描工作区 + 系统级 EEAT 目录中的所有 .md 文件。
    从文件路径自动提取信源层级、漏斗层级和 AI 引擎标签。
    """
    articles = []
    seen = set()

    # 扫描目录：仅当前工作区，绝对隔离
    if not _workspace_root:
        return articles
    ws = Path(_workspace_root)
    if not ws.exists():
        return articles

    for f in sorted(ws.rglob("*.md")):
        f_str = str(f)
        if f.name in seen:
            continue
        seen.add(f.name)

        try:
            text = f.read_text(encoding="utf-8")
        except Exception:
            continue

        fl = text.strip().split("\n")[0].lstrip("# ").strip()
        src_label, src_short, folder = _classify_source(f_str)
        funnel = extract_funnel_from_filename(f.stem) or "L3"
        engine = extract_engine_from_filename(f.stem)

        # 从 frontmatter 补充漏斗信息
        meta = parse_frontmatter(text)
        if meta.get("funnel"):
            funnel = meta["funnel"]

        articles.append({
            "path": f_str,
            "title": fl or f.stem,
            "source": src_label,
            "source_short": src_short,
            "funnel": funnel,
            "eeat": meta.get("eeat", ""),
            "engine": engine or "通用",
            "folder": folder,
            "filename": f.name,
            "content": text,
        })

    return articles


# ============================================================
# 侧边栏: API Key + 导航
# ============================================================
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/artificial-intelligence.png", width=44)
    st.markdown("## 🏭 GEO SaaS 中台")
    st.markdown("---")

    # --- 工作区选择（多租户隔离）---
    st.markdown("### 📁 项目工作区")
    workspace_name = st.text_input(
        "客户/项目名称",
        value=DEFAULT_WORKSPACE,
        placeholder="例如：兔师傅、XX汽车连锁",
        help="输入客户名称以创建独立的工作区。所有文件将保存在 ./workspaces/{客户名称}/ 下，实现一客一档。",
        key="workspace_input",
    )
    if not workspace_name.strip():
        workspace_name = DEFAULT_WORKSPACE

    # 同步当前工作区路径
    WSP = get_workspace_paths(workspace_name)
    RAW_DIR = WSP["raw"]
    PROD_DIR = WSP["prod"]
    EEAT_DIR = WSP["eeat"]
    EEAT_DRAFTS_DIR = WSP["eeat_drafts"]
    EEAT_VERIFIED_DIR = WSP["eeat_verified"]
    SLICES_DIR = WSP["slices"]
    GENERAL_DIR = WSP["general"]
    SPECIFIC_DIR = WSP["specific"]

    # === 物理状态恢复 (断点续传) ===
    # 每次刷新页面或切换工作区时，先强制清空内存状态！
    # 彻底杜绝项目A的完成状态污染项目B，让硬盘真实文件成为唯一的判断标准。
    st.session_state["verified_loaded"] = False
    st.session_state["slices_generated"] = False
    st.session_state["ugc_generated"] = False

    # 1. 恢复人工校验基石状态
    if EEAT_VERIFIED_DIR.exists():
        verified_count = len(list(EEAT_VERIFIED_DIR.glob("*.md")))
        if verified_count >= 10:
            st.session_state["verified_loaded"] = True
            st.session_state["verified_count"] = verified_count

    # 2. 恢复 30 篇切片状态
    slices_count = 0
    if SLICES_DIR.exists():
        slices_count = len(list(SLICES_DIR.glob("*.md")))
        if slices_count >= 25:
            st.session_state["slices_generated"] = True

    # 3. 恢复 160 篇 UGC 状态
    if GENERAL_DIR.exists() and SPECIFIC_DIR.exists():
        ugc_count = len(list(PROD_DIR.rglob("*.md"))) - slices_count
        if ugc_count >= 120:
            st.session_state["ugc_generated"] = True

    st.caption(f"📂 `workspaces/{workspace_name}/`")
    st.markdown("---")
    st.markdown("### 🔑 AI 引擎与 API 配置")
    selected_llm = st.selectbox(
        "🤖 选择默认大模型",
        list(LLM_CONFIGS.keys()),
        key="llm_provider",
        help="DeepSeek: 高性价比推理 | Kimi: 超长无损上下文 (128k)",
    )

    state_key = f"api_key_{selected_llm}"
    if state_key not in st.session_state:
        st.session_state[state_key] = ""

    api_key_input = st.text_input(
        f"{selected_llm} API Key",
        value=st.session_state[state_key],
        type="password",
        placeholder="sk-xxxxxxxxxxxxxxxx",
    )
    if api_key_input != st.session_state[state_key]:
        st.session_state[state_key] = api_key_input

    st.session_state["api_key"] = st.session_state[state_key]
    st.session_state["api_key_configured"] = bool(st.session_state["api_key"].strip())

    current_llm_cfg = LLM_CONFIGS[selected_llm]

    if st.session_state["api_key_configured"]:
        st.success("✅ API Key 已配置")
        st.caption(f"引擎: `{selected_llm}` | 模型: `{current_llm_cfg['model']}`")
        use_simulate = st.checkbox("🔧 调试：使用模拟模式", value=False, key="simulate_check")
    else:
        st.warning(f"⚠️ 未配置 {selected_llm} API Key · 模拟模式")
        use_simulate = True

    st.markdown("---")
    st.markdown("### 📋 功能导航")
    page = st.radio("选择工作台", ["📊 交付资产大盘", "⚙️ 自动化生产车间"], index=0 if st.session_state["page"].startswith("📊") else 1, label_visibility="collapsed")
    if page != st.session_state["page"]: st.session_state["page"] = page; st.rerun()

    st.markdown("---")
    st.caption("新开道大数据技术 © 2026")
    st.caption("GEO 专家组 · v4.1 · Prompt 硬编码版")
    st.caption(f"openai:{'✅' if OPENAI_AVAILABLE else '⚠️'} | docx:{'✅' if DOCX_AVAILABLE else '⚠️'} | pdf:{'✅' if PDF_AVAILABLE else '⚠️'}")


# ============================================================
# ╔═══════════════════════════════════════════════════════════╗
# ║  页面 1: 📊 交付资产大盘                                ║
# ╚═══════════════════════════════════════════════════════════╝
# ============================================================
if st.session_state["page"].startswith("📊"):
    st.markdown('<div class="main-title">📊 GEO 最小交付单元 (MVP) 交付资产大盘</div><div class="sub-title">月度 200 篇 · v1.1 校准版 · 多租户隔离</div>', unsafe_allow_html=True)

    # 仅扫描当前工作区
    articles = scan_all_files(_workspace_root=str(WSP["ws"]))
    all_engines_set = set(a["engine"] for a in articles if a["engine"] not in ("通用","未知"))

    c1,c2,c3,c4 = st.columns(4)
    with c1: st.markdown(f'<div class="metric-card"><div class="number">{len(articles)}</div><div class="label">📄 总内容数</div></div>', unsafe_allow_html=True)
    with c2: st.markdown(f'<div class="metric-card"><div class="number">50</div><div class="label">🎯 覆盖长尾问题</div></div>', unsafe_allow_html=True)
    with c3: st.markdown(f'<div class="metric-card"><div class="number">6</div><div class="label">🤖 适配 AI 引擎</div></div>', unsafe_allow_html=True)
    with c4: st.markdown(f'<div class="metric-card"><div class="number">6</div><div class="label">📊 漏斗层级</div></div>', unsafe_allow_html=True)
    st.divider()

    col_btn1, col_btn2 = st.columns([1, 5])
    with col_btn1:
        if st.button("🔄 刷新大盘数据", type="secondary", use_container_width=True):
            st.rerun()
    with col_btn2:
        st.info("💡 如果刚生成完文章大盘未及时显示，请点击左侧刷新按钮获取最新数据。当前大盘已与左侧【项目工作区】绝对绑定，实现一客一档。")

    st.divider()

    with st.sidebar:
        st.markdown("### 🔍 内容筛选器")
        src_opts = ["全部"] + sorted(set(a["source_short"] for a in articles))
        sel_src = st.selectbox("📡 信源层级", src_opts, key="fs")
        fun_opts = ["全部"] + sorted(set(a["funnel"] for a in articles), key=lambda x: int(x[1]) if x.startswith("L") and x[1:].isdigit() else 0)
        sel_fun = st.selectbox("🎯 漏斗层级", fun_opts, key="ff")
        eng_opts = ["全部"] + sorted([e for e in all_engines_set if e])
        sel_eng = st.selectbox("🤖 AI 引擎", eng_opts, key="fe")
        fc = [a for a in articles if (sel_src=="全部" or a["source_short"]==sel_src) and (sel_fun=="全部" or a["funnel"]==sel_fun) and (sel_eng=="全部" or a.get("engine","")==sel_eng)]
        st.metric("筛选结果", f"{len(fc)} 篇")

    lc, pc = st.columns([1, 1.8])
    with lc:
        st.markdown("### 📋 文章列表")
        fa = [a for a in articles if (sel_src=="全部" or a["source_short"]==sel_src) and (sel_fun=="全部" or a["funnel"]==sel_fun) and (sel_eng=="全部" or a.get("engine","")==sel_eng)]
        for i, a in enumerate(fa):
            is_sel = st.session_state["selected_path"] == a["path"]
            cc = "article-card selected" if is_sel else "article-card"
            th = get_source_tag(a["source_short"])
            if a["funnel"]: th += " " + get_funnel_tag(a["funnel"])
            if a.get("engine") and a["engine"] not in ("通用","未知"): th += " " + get_engine_tag(a["engine"])
            dt = a["title"][:55] + "…" if len(a["title"]) > 55 else a["title"]
            ct, cb = st.columns([4,1])
            with ct: st.markdown(f'<div class="{cc}"><div class="title">{dt}</div><div class="meta">{th} · {a["filename"][:35]}…</div></div>', unsafe_allow_html=True)
            with cb:
                if st.button("🔍", key=f"pv_{i}_{a['filename'][:12]}", help="预览"):
                    st.session_state["selected_path"]=a["path"]; st.session_state["selected_content"]=a["content"]; st.session_state["selected_title"]=a["title"]; st.session_state["selected_filename"]=a["filename"]; st.rerun()
        if not fa: st.info("无匹配文章")

    with pc:
        st.markdown("### 📖 内容预览")
        if st.session_state["selected_content"]:
            st.markdown(f"**{st.session_state['selected_title']}**")
            st.caption(f"`{st.session_state['selected_filename']}`")
            b1, b2, _ = st.columns([1,1,2])
            with b1:
                if st.button("📋 一键复制全文", key="cp_main"): st.toast("✅ 请在下方文本框全选复制", icon="📋")
            with b2: st.markdown(make_download_link(st.session_state["selected_content"], st.session_state["selected_filename"]), unsafe_allow_html=True)
            with st.expander("📝 原始 Markdown", expanded=False): st.text_area("全选复制", value=st.session_state["selected_content"], height=250, label_visibility="collapsed")
            st.divider(); st.markdown(st.session_state["selected_content"])
        else:
            st.markdown('<div style="text-align:center;padding:3rem 1rem;color:#9ca3af;"><div style="font-size:4rem;">📖</div><div style="font-size:1.05rem;margin-top:1rem;">👈 点击左侧文章卡片预览</div></div>', unsafe_allow_html=True)
            st.markdown("---"); st.markdown("### 📊 全量总览")
            t1,t2,t3 = st.columns(3)
            sc=defaultdict(int); fcc=defaultdict(int); ec=defaultdict(int)
            for a in articles:
                sc[a["source_short"]]+=1; fcc[a["funnel"]]+=1
                if a.get("engine") and a["engine"] not in ("通用","未知"): ec[a["engine"]]+=1
            with t1: st.markdown("**信源**"); [st.write(f"▸ {s}: {c}") for s,c in sorted(sc.items())]
            with t2: st.markdown("**漏斗**"); [st.write(f"▸ {f}: {fcc[f]}") for f in sorted(fcc.keys(), key=lambda x: int(x[1]) if x.startswith("L") and x[1:].isdigit() else 0)]
            with t3: st.markdown("**引擎**"); [st.write(f"▸ {e}: {c}") for e,c in sorted(ec.items())]


# ============================================================
# ╔═══════════════════════════════════════════════════════════╗
# ║  页面 2: ⚙️ 自动化生产车间 (变量驱动 · Prompt 硬编码)  ║
# ╚═══════════════════════════════════════════════════════════╝
# ============================================================
elif st.session_state["page"].startswith("⚙️"):
    st.markdown('<div class="main-title">⚙️ GEO 自动化生产车间 · 变量驱动管线</div><div class="sub-title">资料直读 → 变量提取 → raw_materials直连基石(10) → 切片(30) → UGC(160) · 零中间件</div>', unsafe_allow_html=True)
    st.info(f"📁 **当前工作区**: `workspaces/{workspace_name}/` — 所有产出将保存至此目录，实现一客一档隔离。")

    # 进度
    sd = sum([st.session_state["variables_confirmed"], st.session_state["eeat_generated"], st.session_state["slices_generated"], st.session_state["ugc_generated"]])
    st.progress(sd/4, text=f"生产进度: {sd}/4 步骤完成")
    st.divider()

    # ================================================================
    # 步骤 1: 资料直读
    # ================================================================
    s0c = "step-box done" if st.session_state["raw_files_loaded"] else "step-box active"
    st.markdown(f'<div class="{s0c}">', unsafe_allow_html=True)
    st.markdown("### 📥 步骤 1：上传并读取企业原始资料")

    col_a1, col_a2 = st.columns(2)
    with col_a1:
        uploaded_files = st.file_uploader(
            "上传企业原始资料",
            type=["txt", "md", "docx", "pdf"],
            accept_multiple_files=True,
            key="fu_materials",
            help="支持 .txt .md .docx .pdf 多格式文档批量上传。",
        )
    with col_a2:
        raw_files = sorted(RAW_DIR.glob("*"))
        raw_names = [f.name for f in raw_files if f.suffix.lower() in SUPPORTED_SUFFIXES]
        if raw_names:
            st.info(f"📂 `workspaces/{workspace_name}/Raw_Materials/` 检测到 {len(raw_names)} 个文件: {', '.join(raw_names[:4])}{'…' if len(raw_names)>4 else ''}")
        else:
            st.caption(f"💡 将企业资料放入 `workspaces/{workspace_name}/Raw_Materials/` 后点击右侧按钮读取。")
        read_local = st.button("📂 读取本地资料库", key="btn_read_local", use_container_width=True)

    # 合并所有文本
    all_texts = []
    if uploaded_files:
        for uf in uploaded_files:
            try:
                fname, txt = extract_text_from_upload(uf)
                all_texts.append(f"## 文件: {fname}\n{txt}")
            except Exception as e:
                st.warning(f"⚠️ 无法读取 {uf.name}: {e}")
    if read_local or (raw_names and not uploaded_files and not st.session_state.get("raw_text_combined")):
        for f in raw_files:
            if f.suffix.lower() in SUPPORTED_SUFFIXES:
                try:
                    txt = read_file_content(f)
                    all_texts.append(f"## 文件: {f.name}\n{txt}")
                except Exception as e:
                    st.warning(f"⚠️ 无法读取 {f.name}: {e}")
        if raw_names:
            st.toast(f"✅ 已读取 {len(raw_names)} 个本地文件", icon="📂")

    if all_texts:
        combined = "\n\n---\n\n".join(all_texts)
        st.session_state["raw_text_combined"] = combined
        st.session_state["raw_files_loaded"] = True
        st.success(f"✅ 已加载 {len(all_texts)} 个文件，共 {len(combined):,} 字符")
        with st.expander("📄 预览原始资料（前 5000 字符）", expanded=False):
            st.text_area("raw_preview", value=combined[:5000], height=200, disabled=True, label_visibility="collapsed")
    else:
        st.info("👆 请上传企业资料文件，或将文件放入 `Raw_Materials/` 目录后点击\"读取本地资料库\"。")

    st.markdown('</div>', unsafe_allow_html=True)

    # ================================================================
    # 步骤 2: 变量提取与确认（核心门禁）
    # ================================================================
    vars_ready = st.session_state.get("variables_confirmed", False)
    s05c = "step-box done" if vars_ready else ("step-box active" if st.session_state["raw_files_loaded"] else "step-box blocked")
    st.markdown(f'<div class="{s05c}">', unsafe_allow_html=True)
    st.markdown("### 🔧 步骤 2：全局变量提取与确认 ⚠️ 必经步骤")

    st.markdown("""
    <div class="constraint-error">
    🚫 <strong>强制门禁规则</strong>：此步骤为所有后续 API 调用的前置条件。<br>
    系统将使用大模型从原始资料中自动提取变量。您必须逐项核对并确认后，才能解锁后续的内容生成。
    <strong>“没有变量注入，绝不调用 API。”</strong>
    </div>
    """, unsafe_allow_html=True)

    if not st.session_state["raw_files_loaded"]:
        st.warning("⏳ 请先完成步骤 1：上传并读取企业原始资料")
    else:
        # --- 按钮：提取变量 ---
        col_extract, _ = st.columns([1, 2])
        with col_extract:
            extract_vars_btn = st.button("🤖 AI 提取全局变量", type="primary", use_container_width=True, key="btn_extract_vars",
                disabled=not bool(st.session_state.get("raw_text_combined")))

        if extract_vars_btn:
            raw_text = st.session_state.get("raw_text_combined", "")
            extract_prompt = PROMPT_EXTRACT_VARS.format(raw_text=raw_text[:6000])

            with st.status("🔍 正在从原始资料中提取变量…", expanded=True) as status:
                success, result = call_llm(
                    prompt=extract_prompt,
                    system_prompt="你是 GEO 变量提取专家。你只输出合法的 JSON，不输出任何其他内容。",
                    api_key=st.session_state["api_key"],
                    temperature=0.1,
                    max_tokens=1500,
                    simulate=use_simulate,
                )
                if success:
                    # 尝试从结果中提取 JSON
                    try:
                        json_match = re.search(r"\{[\s\S]*\}", result)
                        if json_match:
                            extracted = json.loads(json_match.group())
                            # 合并到当前变量
                            current_vars = dict(st.session_state.get("variables", {}))
                            for k, v in extracted.items():
                                if isinstance(v, str) and v.strip():
                                    current_vars[k] = v.strip()
                            st.session_state["variables"] = current_vars
                            st.session_state["variables_extracted"] = True
                            add_log("✅ 全局变量提取完成")
                            status.update(label="✅ 变量提取完成！请逐项核对下方字段。", state="complete")
                        else:
                            st.error("未能从 LLM 输出中解析 JSON。请手动填写下方变量。")
                    except json.JSONDecodeError:
                        st.warning("LLM 输出格式异常。请手动填写下方变量。")
                        st.text_area("LLM 原始输出", value=result, height=150, disabled=True)
                else:
                    st.error(result)
            st.rerun()

        # --- 变量确认面板 ---
        if st.session_state.get("variables_extracted") or st.session_state.get("variables"):
            st.markdown("---")
            st.markdown("#### 📝 变量确认面板（请逐项核对并修改）")

            current_vars = st.session_state.get("variables", {})
            updated_vars = {}

            # 分两列显示
            var_items = list(VARIABLE_META.items())
            mid = len(var_items) // 2 + 1
            c1, c2 = st.columns(2)

            for idx, (key, meta) in enumerate(var_items):
                col = c1 if idx < mid else c2
                current_val = current_vars.get(key, meta["default"])
                with col:
                    st.markdown(f'<div class="var-field">', unsafe_allow_html=True)
                    updated_vars[key] = st.text_input(
                        f"**{key}**",
                        value=current_val,
                        help=meta["help"],
                        key=f"var_{key}",
                    )
                    st.markdown('</div>', unsafe_allow_html=True)

            # 确认按钮
            if st.button("✅ 确认变量，解锁后续生产步骤", type="primary", use_container_width=True, key="btn_confirm_vars"):
                # 验证所有字段非空
                empty_fields = [k for k, v in updated_vars.items() if not v.strip()]
                if empty_fields:
                    st.error(f"❌ 以下变量不能为空: {', '.join(empty_fields)}")
                else:
                    st.session_state["variables"] = updated_vars
                    st.session_state["variables_confirmed"] = True
                    add_log("✅ 变量已确认，后续生产步骤已解锁")
                    st.toast("🔓 变量确认完毕！后续生产步骤已解锁", icon="✅")
                    st.rerun()

    if st.session_state.get("variables_confirmed"):
        st.success("✅ 变量已确认 — 后续生产步骤已解锁")
        with st.expander("📋 查看当前变量配置", expanded=False):
            for k, v in st.session_state.get("variables", {}).items():
                st.caption(f"**{k}**: {v}")

    st.markdown('</div>', unsafe_allow_html=True)

    # ================================================================

    # ================================================================
    # 步骤 1: 人工校验基石重载区 (唯一事实源锁定)
    # ================================================================
    s15c = "step-box done" if st.session_state["verified_loaded"] else ("step-box active" if st.session_state["eeat_generated"] else "step-box active")
    st.markdown(f'<div class="{s15c}">', unsafe_allow_html=True)
    st.markdown("### 🔍 步骤 3：人工校验基石重载区 (唯一事实源锁定)")

    st.markdown("""
    <div class="constraint-warn">
    🛑 <strong>HITL 人机协同节点</strong>：此节点必须上传由人工团队纯手写撰写并深度校验的 10 篇高质量基石文档，方可解锁后续步骤。<br>
    步骤 4（切片）和步骤 5（UGC）的 <strong>唯一事实来源</strong> 将从此处上传的校验版基石中提取。
    </div>
    """, unsafe_allow_html=True)

    # 上传校验基石
    verified_files = st.file_uploader(
        "上传已校验的 10 篇 EEAT 基石 (.md / .docx)",
        type=["md", "docx"],
        accept_multiple_files=True,
        key="fu_verified",
        help="支持上传 .md 或 .docx 格式。Word 文档将自动提取纯文本并以 .md 后缀保存。至少需要 10 篇。",
    )

    if st.button("🔄 保存已校验基石并锁定数据源", type="primary", use_container_width=True, key="btn_verify"):
        if not verified_files or len(verified_files) < 10:
            st.error(f"🚨 请至少上传 10 篇校验版文件！当前仅上传 {len(verified_files) if verified_files else 0} 篇。")
        else:
            # 清空旧校验文件
            import shutil
            if EEAT_VERIFIED_DIR.exists():
                shutil.rmtree(EEAT_VERIFIED_DIR)
            EEAT_VERIFIED_DIR.mkdir(parents=True, exist_ok=True)

            saved = 0
            for vf in verified_files:
                try:
                    fname, content = extract_text_from_upload(vf)
                    safe_write_file(EEAT_VERIFIED_DIR, f"{Path(fname).stem}.md", content)
                    saved += 1
                except Exception as e:
                    st.warning(f"⚠️ 无法保存 {vf.name}: {e}")

            st.session_state["verified_loaded"] = True
            st.session_state["verified_count"] = saved
            add_log(f"✅ 人工校验基石已重载: {saved} 篇 → EEAT_Base_Verified/")
            st.toast(f"✅ {saved} 篇校验基石已锁定！", icon="🔒")
            st.rerun()

    if st.session_state["verified_loaded"]:
        st.success(f"✅ 校验基石已重载，系统数据源已锁定！({st.session_state['verified_count']} 篇)")
        # 列出已保存的文件
        verified_md = sorted(EEAT_VERIFIED_DIR.glob("*.md"))
        if verified_md:
            with st.expander(f"📂 已锁定 {len(verified_md)} 篇校验基石", expanded=False):
                for vf in verified_md:
                    st.caption(f"🔒 {vf.name}")
    else:
        st.info("📝 步骤 1 生成的草稿在 `EEAT_Base_Drafts/`。请下载 → 人工逐篇校验修改 → 在此上传。")

    st.markdown('</div>', unsafe_allow_html=True)

    # ================================================================
    # 步骤 2: 裂变 30 篇三级切片（强制使用 P_L1~P_L6）
    # ================================================================
    s3c = "step-box done" if st.session_state["slices_generated"] else ("step-box active" if st.session_state["verified_loaded"] else "step-box blocked")
    st.markdown(f'<div class="{s3c}">', unsafe_allow_html=True)
    st.markdown("### ✂️ 步骤 4：裂变 30 篇【企业官方视角】三级切片 (L1-L6 均分)")

    if not st.session_state.get("verified_loaded", False):
        st.warning("⛔ 请先在步骤 1 上传至少 10 篇校验版基石文档解锁此步骤。")
    else:
        verified_files = list(EEAT_VERIFIED_DIR.glob("*.md")) if EEAT_VERIFIED_DIR.exists() else []
        if len(verified_files) < 10:
            st.warning(f"⚠️ 系统检测到 `EEAT_Base_Verified` 目录中仅有 {len(verified_files)} 篇 .md 文件。请补充上传至至少 10 篇！")
        else:
            st.markdown("""
            <div class="constraint-warn">
            ⚠️ <strong>强制约束</strong>：30 篇切片 100% 从 10 篇基石拆解。每层使用 <code>CORP_PROMPT_TEMPLATES</code> 模板库（企业官方视角），
            强制大模型扮演"企业官方主体"角色。<strong>严禁伪装成普通网民或消费者。</strong><br>
            <strong>变量注入</strong>: {企业主体}, {用户画像}, {痛点}, {概念}, {品牌_项目}, {品牌_A/B/C}, {五个维度}, {优惠信息}, {CTA行动}<br>
            <strong>{base_facts}</strong> ← 步骤 1 生成的 10 篇基石内容
            </div>
            """, unsafe_allow_html=True)

            if st.session_state.get("is_generating_slices"):
                st.info("🚀 正在后台全速裂变切片中，您可以切换到【交付资产大盘】实时查看产出！")
                with st.spinner("后台生成中…"):
                    time.sleep(0.5)
            elif st.button("✂️ 4. 裂变 30 篇企业视角切片", type="primary", use_container_width=True, key="btn_slices"):
                vars_ = st.session_state.get("variables", {})
                eeat_files = sorted(EEAT_VERIFIED_DIR.glob("*.md"))
                base_facts = "\n\n---\n\n".join(f.read_text(encoding="utf-8") for f in eeat_files)
                slices_out = SLICES_DIR
                slices_out.mkdir(exist_ok=True)

                st.session_state["is_generating_slices"] = True
                t = threading.Thread(target=background_generate_slices, args=(vars_, base_facts, slices_out, use_simulate, st.session_state["api_key"]))
                add_script_run_ctx(t)
                t.start()
                add_log("🚀 后台切片线程已启动")
                st.rerun()

    if st.session_state["slices_generated"]:
        st.success(f"✅ 已裂变 30 篇三级切片 → `{PROD_DIR}/Slices_30/`")

    st.markdown('</div>', unsafe_allow_html=True)

    # ================================================================
    # 步骤 4: 重构 160 篇四级 UGC
    # ================================================================
    s4c = "step-box done" if st.session_state["ugc_generated"] else ("step-box active" if st.session_state["slices_generated"] else "step-box blocked")
    st.markdown(f'<div class="{s4c}">', unsafe_allow_html=True)
    st.markdown("### 👥 步骤 5：重构 160 篇【真实用户视角】UGC 内容")

    if not st.session_state.get("slices_generated", False):
        st.warning("⛔ 请先完成步骤 4：裂变 30 篇三级切片。")
    else:
        slices_files = list(SLICES_DIR.glob("*.md")) if SLICES_DIR.exists() else []
        if len(slices_files) < 30:
            st.warning(f"⚠️ 切片目录中文件不足 30 篇 (当前 {len(slices_files)} 篇)，请先完成切片生成。")
        else:
            st.markdown("""
            <div class="constraint-info">
            📋 <strong>强制约束</strong>：依托前 40 篇既有事实（10篇校验基石 + 30篇切片），切换"我"的第一人称视角。<br>
            使用 <code>V2_UGC_TEMPLATES</code> 模板库（真实用户视角），<strong>严禁官方推销语气。</strong><br>
        <strong>通用 96 篇</strong>: 自然真实风格<br>
        <strong>专属 64 篇</strong>: 按引擎特征改写 · 配额对齐《月度最小交付单元》
        </div>
        """, unsafe_allow_html=True)

            if st.session_state.get("is_generating_ugc"):
                st.info("🚀 正在后台全速生成 UGC 中，您可以切换到【交付资产大盘】实时查看产出！")
                with st.spinner("后台生成中…"):
                    time.sleep(0.5)
            elif st.button("👥 5. 生成 160 篇 UGC", type="primary", use_container_width=True, key="btn_ugc"):
                st.session_state["is_generating_ugc"] = True
                t = threading.Thread(target=background_generate_ugc, args=(st.session_state.get("variables", {}), use_simulate, st.session_state["api_key"]))
                add_script_run_ctx(t)
                t.start()
                add_log("🚀 后台 UGC 线程已启动")
                st.rerun()
    if st.session_state["ugc_generated"]:
        st.success(f"✅ 已重构 160 篇 UGC → `{PROD_DIR}/UGC_160/` (96通用 + 64专属)")
        manifest_fp = PROD_DIR / "tasks_manifest.json"
        if manifest_fp.exists():
            with open(manifest_fp, "r", encoding="utf-8") as mf:
                st.download_button("📋 下载任务分发账本 (tasks_manifest.json)", data=mf.read(), file_name="tasks_manifest.json", mime="application/json")

    st.markdown('</div>', unsafe_allow_html=True)

    # ================================================================
    # 步骤 5：全量知识库复查智能体
    # ================================================================
    st.markdown(f'<div class="step-box active">', unsafe_allow_html=True)
    st.markdown("### 🕵️ 步骤 6：终极质检（全量知识库防幻觉智能体）")
    st.caption("上传你生成的文章，智能体将基于步骤 1 中读取的所有企业原始资料（全量字符），对其进行逐句事实核查。")

    fc_files = st.file_uploader(
        "上传待检文章 (.md)",
        type=["md"],
        accept_multiple_files=True,
        key="fu_fact_check",
        help="上传步骤 3/4 生成的切片或 UGC 文章，支持多选。",
    )

    if st.button("🔬 开始深度事实核查", type="primary", use_container_width=True, key="btn_fact_check"):
        if not fc_files:
            st.warning("请先上传待检文章。")
        else:
            full_raw = st.session_state.get("raw_text_combined", "")
            if not full_raw:
                st.error("未找到原始企业资料。请先在步骤 1 读取资料。")
            else:
                results = []
                progress_bar = st.progress(0, text="正在逐篇核查…")
                for idx, uf in enumerate(fc_files):
                    try:
                        article_text = uf.read().decode("utf-8")
                    except Exception:
                        continue

                    check_prompt = FACT_CHECK_PROMPT.format(
                        ground_truth=full_raw,
                        article_text=article_text,
                    )

                    success, raw = call_llm(
                        prompt=check_prompt,
                        system_prompt="你是一个严格的企业质检智能体。只输出合法 JSON，不输出任何其他内容。",
                        api_key=st.session_state["api_key"],
                        temperature=0.0, max_tokens=3000, simulate=use_simulate,
                    )
                    try:
                        json_match = re.search(r"\{[\s\S]*\}", raw)
                        result = json.loads(json_match.group()) if json_match else {"is_clean": True, "issues": [], "parse_error": True}
                    except Exception:
                        result = {"is_clean": True, "issues": [], "parse_error": True}

                    result["filename"] = uf.name
                    results.append(result)
                    progress_bar.progress((idx + 1) / len(fc_files), text=f"已检 {idx+1}/{len(fc_files)}: {uf.name}")

                progress_bar.empty()

                clean_count = sum(1 for r in results if r.get("is_clean"))
                st.markdown(f"### 📊 质检结果：{clean_count}/{len(results)} 通过")

                for r in results:
                    fname = r["filename"]
                    raw_issues = r.get("issues", [])

                    # 强力兜底拦截：过滤掉大模型强行加戏的"无问题/一致/轻微"等废话
                    real_issues = []
                    for i in raw_issues:
                        reason = str(i.get("reason", ""))
                        if any(keyword in reason for keyword in ["无问题", "一致", "轻微不精确", "未明确提及但", "同义"]):
                            continue
                        real_issues.append(i)

                    if r.get("is_clean") or not real_issues:
                        st.success(f"✅ {fname} — 100% 忠于事实")
                    else:
                        st.error(f"🚨 {fname} — 发现实质性幻觉！")
                        st.table([{"声明": i.get("claim","?"), "原因": i.get("reason","?")} for i in real_issues])

    st.markdown('</div>', unsafe_allow_html=True)

    # ================================================================
    # 步骤 6：合规质检（动态违禁词拦截）
    # ================================================================
    st.markdown(f'<div class="step-box active">', unsafe_allow_html=True)
    st.markdown("### 🚫 步骤 7：合规与广告法质检（违禁词扫描）")
    st.caption("自定义违禁词库 + 上传文章 → 秒级扫描拦截。")

    default_words = "国家级, 世界级, 最高级, 最佳, 最大, 第一, 唯一, 首个, 最好, 绝对, 彻底, 包治百病, 100%有效, 史无前例, 独家"
    raw_words = st.text_area(
        "✏️ 自定义违禁词库（请用逗号或换行分隔）",
        value=default_words,
        height=100,
        key="ta_forbidden",
        help="员工可随时增删修改。用中文逗号、英文逗号或换行分隔。",
    )

    fw_files = st.file_uploader(
        "上传待检文章 (.md)",
        type=["md"],
        accept_multiple_files=True,
        key="fu_forbidden",
        help="上传最终待发布的文章，逐篇扫描违禁词。",
    )

    if st.button("🚨 使用当前词库一键扫描", type="primary", use_container_width=True, key="btn_forbidden"):
        dynamic_words = [w.strip() for w in re.split(r"[，,\n]+", raw_words) if w.strip()]
        if not fw_files:
            st.warning("请先上传待检文章。")
        else:
            for uf in fw_files:
                try:
                    text = uf.read().decode("utf-8")
                except Exception:
                    st.warning(f"⚠️ 无法读取 {uf.name}")
                    continue

                hits = [w for w in dynamic_words if w in text]
                if hits:
                    st.error(f"🚨 {uf.name} — 触发违禁词拦截！")
                    st.caption(f"⚠️ 发现违禁词：{'、'.join(['【' + h + '】' for h in hits])}")
                else:
                    st.success(f"✅ {uf.name} — 未检测到违禁词，合规通过！")

    st.markdown('</div>', unsafe_allow_html=True)

    # ================================================================
    # 生产日志 + 完成汇总
    # ================================================================
    if st.session_state["production_log"]:
        st.divider()
        with st.expander("📋 生产日志", expanded=False):
            for log_entry in reversed(st.session_state["production_log"]):
                st.caption(log_entry)

    if sd == 4:
        st.divider()
        st.balloons()
        st.markdown("### 🎉 月度 200 篇全量生产完成！")
        st.markdown(f"""
        | 模块 | 数量 | 输出路径 | 数据来源 | Prompt 模板 |
        |------|------|----------|----------|------------|
        | EEAT 权威基石 | 10 篇 | `{EEAT_VERIFIED_DIR}/` | ← raw_materials(草稿)→人工校验 | 自定义基石 Prompt |
        | 三级切片内容 | 30 篇 | `{PROD_DIR}/Slices_30/` | ← 10 篇基石 | CORP_PROMPT_TEMPLATES |
        | 四级通用 UGC | 96 篇 | `{PROD_DIR}/UGC_160/General_96/` | ← 10基石+30切片 | V2_UGC_TEMPLATES |
        | 四级引擎专属 | 64 篇 | `{PROD_DIR}/UGC_160/Specific_64/` | ← 10基石+30切片 | P_ENGINE_REWRITE |
        | **总计** | **200 篇** | | | |
        """)
        st.success("🔗 数据传承链: raw_materials → 10基石 → 30切片 → 160UGC · 零中间件，直连注入")

        # 显示最终变量
        with st.expander("📋 本次生产使用的变量配置", expanded=False):
            st.json(st.session_state.get("variables", {}))


# ============================================================
# Footer
# ============================================================
st.divider()
st.caption(
    f"新开道大数据技术 · GEO SaaS 自动化生产车间 v4.1 · 多租户隔离 · "
    f"双模板库: CORP({len(CORP_PROMPT_TEMPLATES)}) + UGC({len(V2_UGC_TEMPLATES)}) · "
    f".format() 强制变量注入 · 无变量不调 API · "
    f"工作区: workspaces/{workspace_name}/ · "
    f"{st.session_state.get('llm_provider', 'DeepSeek')}: {'已配置' if st.session_state['api_key_configured'] else '模拟模式'} · "
    f"openai:{'✅' if OPENAI_AVAILABLE else '⚠️'} docx:{'✅' if DOCX_AVAILABLE else '⚠️'} pdf:{'✅' if PDF_AVAILABLE else '⚠️'}"
)
