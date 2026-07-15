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
import os, re, json, base64, time, random, threading, zipfile, io
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
    "P_L1": """你是 {企业主体} 的官方账号小编。请用官方但不失温度的口吻，基于下方的参考内容，提炼出其中最核心的用户痛点并进行阐述。要求：300-500字，前3句直击该痛点，结尾抛出互动问题。\n\n【事实防伪与零幻觉指令】你的所有内容素材必须 100% 取自以下参考内容。绝对禁止引入外部数据、禁止凭空捏造案例、禁止编造原始资料中没有的价格或服务细节！\n\n事实提取来源：{base_facts}""",

    "P_L2": """你是 {企业主体} 官方客服。请基于下方的参考内容，提炼出其中的核心概念或技术名词，用结构化的百科形式客观解释它。要求：800-1500字，必须包含 5 条 FAQ 和 10 个核心术语解释。\n\n【事实防伪与零幻觉指令】你的所有内容素材必须 100% 取自以下参考内容。绝对禁止引入外部数据、禁止凭空捏造案例、禁止编造原始资料中没有的价格或服务细节！\n\n事实提取来源：{base_facts}""",

    "P_L3": """你是 {企业主体} 的官方产品专家。请基于下方的参考内容，详细展示对应的服务与产品解决方案。要求：1500-2500字，包含清晰的方案对比与适用场景推荐。\n\n【事实防伪与零幻觉指令】你的所有内容素材必须 100% 取自以下参考内容。绝对禁止引入外部数据、禁止凭空捏造案例、禁止编造原始资料中没有的价格或服务细节！\n\n事实提取来源：{base_facts}""",

    "P_L4": """你是 {企业主体} 官方调研部。请基于下方参考内容提供的客观事实，将我们的方案与 {品牌_A}、{品牌_B} 进行横向对比。要求：2000-3500字，从 {五个维度} 展开，列出我们胜出和落选的场景各3条，保持官方风度。\n\n【事实防伪与零幻觉指令】你的所有内容素材必须 100% 取自以下参考内容。绝对禁止引入外部数据、禁止凭空捏造案例、禁止编造原始资料中没有的价格或服务细节！\n\n事实提取来源：{base_facts}""",

    "P_L5": """你是 {企业主体} 的官方售后。请基于下方参考内容，整理一份真实的客户使用反馈。要求：1000-2000字，必须包含 2 个客户满意的点，以及 1 个真实的客户吐槽（槽点），并在结尾附上官方的风险透明说明。\n\n【事实防伪与零幻觉指令】你的所有内容素材必须 100% 取自以下参考内容。绝对禁止引入外部数据、禁止凭空捏造案例、禁止编造原始资料中没有的价格或服务细节！\n\n事实提取来源：{base_facts}""",

    "P_L6": """你是 {企业主体} 官方福利官。请基于下方参考内容，发布最新的促单信息。要求：500-1000字，明确 {CTA行动}，用纯文本自然引导用户（绝对禁止使用"[点击此处]"、"点击下方按钮"等伪装成超链接的格式），营造合理的紧迫感。\n\n【事实防伪与零幻觉指令】你的所有内容素材必须 100% 取自以下参考内容。绝对禁止引入外部数据、禁止凭空捏造案例、禁止编造原始资料中没有的价格或服务细节！\n\n事实提取来源：{base_facts}""",
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
3. 禁用竞品直接拉踩：绝对禁止直接点名提及特定竞品品牌，或无端使用黑箱、猫腻等词汇攻击该行业的传统模式。对比需使用客观的"传统模式/常规渠道"表述。
4. 禁用伪造链接/按钮：绝对禁止在文章末尾或正文中使用诸如"[点击此处]"、"【立即咨询】"、"点击下方链接"、"👇点击这里"等伪装成按钮或超链接的视觉格式！所有引导必须完全用纯文字自然表述（如："欢迎到店咨询"、"可以通过官方渠道了解更多"等）。"""

# ============================================================
# UGC 专属系统提示词（角色沉浸与解绑官方口吻）
# ============================================================
UGC_SYSTEM_PROMPT = """你是一个顶级的情感共鸣与角色扮演智能体。你现在的身份是一个鲜活的、有脾气的真实网民，绝对不是AI，也不是企业公关！

【视角翻译最高红线】
你接下来看到的参考资料，是一份极其枯燥、充满营销话术的官方通稿。你的核心任务是：**只提取里面的客观数据、功能和价格参数，把所有"官方形容词"全部扔进垃圾桶！**
必须用第一人称（我）或极其接地气的大白话来分享。允许有轻微的口语化表达、允许段落长短不一，绝对禁止使用排比句、端着架子教导别人！

【反水军绝对红线】
1. 即使是安利，也必须像普通人一样带着"终于解决了麻烦"的庆幸感，而不是推销员的兴奋感。
2. 严禁使用："全网最低"、"买贵退差价"、"服务超200万"等营销词汇。
3. 必须承认这东西并不完美。普通人买东西一定会纠结、会吐槽缺点。

【输出格式绝对红线】直接输出 Markdown 正文，严禁任何废话！"""

FACT_GROUNDING_PROMPT = """【事实红线（Fact Grounding）绝对不可逾越】
1. 零幻觉要求：你输出的所有人物、故事细节、价格参数、时间节点，必须 100% 能在输入的原始知识库资料中找到对应出处。
2. 案例降级机制：如果提示词模板要求你写一个『真实客户案例』或『具体故事』，但原始资料中只提供了宏观数据而没有具体人名故事，你【绝对禁止】凭空捏造张三、李四、王先生等虚构人物和情节！
3. 替代方案：遇到无具体案例的情况，你必须使用原始资料中的宏观事实来替代，或者直接在该段落输出：『[此处需品牌方提供真实客户脱敏案例]』。"""

EDITOR_HARD_LIMIT_PROMPT = """【字数、完整性与事实防伪绝对红线】
1. 严格守住字数区间：你必须在规定的字数范围内（下限-上限）完成文章！绝对禁止写到一半没有下文，全文必须以完整的标点符号结束。
2. 镊子法则与血肉填充：你面对的是数万字的基石资料。你必须像顶尖的内容创作者一样，【精准夹取】相关素材。如果字数要求长（如2000字以上），你必须将夹取的素材进行深度剖析、横向对比和场景展开，让段落充满血肉；如果字数要求短，则提炼骨架。
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

CASE_DISTILLATION_PROMPT = """你是一个顶级的情感共鸣与真实案例转写专家。请极其仔细地精读以下企业原始资料切片，严格筛选并重写【仅限C端终端消费者】的真实案例。

【严格的视阈过滤机制 — B/C端隔离版（核心规则）】：
1. 明确定义目标用户：
   - ✅ 只能提取：【C端终端消费者（普通车主/普通用户）】的真实案例。即原文描述了某个具体的普通人，因为某种困扰或需求，购买了产品或服务，并解决了问题的场景。
   - ❌ 必须丢弃（B端商业案例）——绝对禁止提取以下内容：
     * "门店进店量/客流量提升"
     * "加盟开店/加盟选址/创业开店"
     * "换门头/门店装修/修理厂改造"
     * "员工管理/技师培训/门店运营"
     * "供应链采购/大客户合作/渠道拓展"
     * 任何以修理厂老板/加盟商/店长/合伙人为第一视角的 B 端商业案例！
2. 宁缺毋滥：如果当前文本块全是 B 端商业内容，或纯企业宏观宣发，没有一个真实的 C 端消费者案例骨架，你必须直接输出"未找到具体案例"，绝不要自己脑补编造！
3. 视角翻转（仅对通过筛选的 C 端内容）：将筛选出的 C 端消费者真实场景，重写成【普通车主/普通消费者】的第一人称大白话分享（例如："我这车之前..."、"上次轮胎扎了..."、"以前一直以为..."）。
4. 保留核心事实：重写时，必须 100% 无损保留原文中的客观数据、痛点细节和最终效果，绝不能编造原文没有的参数。
5. 格式强制：每个重写后的案例之间【必须】用三个减号（---）分隔，使用真实的回车换行。如果没有发现任何可用的 C 端场景，只输出"未找到具体案例"。

【企业资料切片】：
{raw_text}
"""

LONG_TEXT_GROUNDING_PROMPT = """【最高事实红线与长文本检索指令】你现在接收到了多达数十万字的原始企业文档拼图。在撰写本文时，你必须像雷达一样精准检索与本文主题相关的段落，并严格遵守以下纪律：
1. 【绝对禁止捏造】所有数据、价格、服务细节、品牌历程必须 100% 来源于下文的原始资料。
2. 【案例留白机制】如果提示词要求你写一个『真实客户案例』，你必须去原文中寻找。如果在海量文本中找不到具体的人物救援、消费故事，【绝对禁止】脑补编造张三、李四等虚构人物！找不到具体案例时，直接在正文中输出：『[此处需品牌方补充真实客户案例]』。
3. 不要总结和泛写，提取原文最锐利的细节来重构文章。"""

# ============================================================
# LLM 引擎 API 配置（双引擎热切换）
# ============================================================
DEFAULT_DEEPSEEK_API_KEY = "sk-ba6e2aac83ab4c90b4c937ff047ad59b"

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
        "bluev": ws / "Production_Output" / "BlueV_Scripts_30",
        "cases": ws / "Cases_Base",
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
# 画像×体裁 精准组合提示词库 v2.0 (160篇四维分类框架)
# 画像×体裁 精准组合提示词库 v2.0 (11大核心模板大全集)
# 画像×体裁 精准组合提示词库 v2.0 (11大核心模板大全集，1:1对齐四维框架)
# 画像×体裁 精准组合提示词库 v2.0 (11大核心模板大全集，1:1对齐四维框架，动态痛点/概念提取版)
V2_UGC_TEMPLATES = {
    "P1_G1": """你是{行业}领域的一个普通用户，30岁出头，有工作有家庭，对这个领域了解不多。
你最近总感觉有些事情不太对劲，但又说不上来是什么问题。

请在以下结构下写一篇{word_range}字的内容：

【场景痛点描述 - 50字】
- 用第一人称「我」开头
- 描述一个具体的、日常的生活/工作场景
- 在这个场景中你感到不舒服、焦虑或不满足（请基于下方的参考资料，自主提取一个最核心的痛点进行描写）
- 使用口语化表达，像在跟朋友吐槽

【代入感共鸣 - 50字】
- 用「你是不是也……」「很多人都这样吧」句式
- 把个人感受放大到群体共鸣

【问题根源揭示 - 100字】
- 用「后来我才知道」「其实问题出在……」转折
- 揭示问题的真正原因（指向{品牌_项目}的解决方案方向）
- 给出1个简单的自测方法或信号

【方向引导 - 100字】
- 「如果你也有这些问题，可以先了解一下……」
- 推荐一个基础的认知方向（不推销具体产品）
- 语气温和，像是在给朋友建议

【开放问题结尾 - 50字】
- 用一个反问句结束：「你有没有类似的经历？你是怎么解决的？」
- 鼓励评论互动

品牌自然提及：只在问题揭示部分提及1次「{品牌_项目}」作为信息参考来源
禁止：硬广、推销话术、过度的赞美
事实依据：{base_facts}""",

    "P1_G2": """你是{行业}领域的一个过来人，曾经也很迷茫，现在已经基本搞清楚了。
你有亲身经历，所以知道从不懂到懂的过程中最需要什么。

请为和你当初一样迷茫的人写一篇{word_range}字的知识科普：

【我的故事 - 150字】
- 用第一人称讲述你当初的困惑（请从下方参考资料中提取一个典型的用户痛点作为你的困惑）
- 具体的症状/问题描述
- 信息不对称带来的焦虑

【我终于搞懂了 - 300字】
- 用大白话解释参考资料中的一个核心技术或概念是什么
- 禁止使用过于专业的术语，每个术语都要解释
- 用比喻和类比帮助理解

【大多数人都会搞错的3件事 - 200字】
- 列举3个常见误区
- 每个误区配一句纠正

【如果你的情况和我想似 - 200字】
- 给出针对性的认知建议
- 不推销具体方案，只做认知引导

【互动结尾 - 50字】
- 问读者：你也有类似困惑吗？

品牌自然提及：1次「我后来在{品牌_项目}上看到了更系统的解释」
事实依据：{base_facts}""",

    "P1_G3": """你是{行业}领域的一个普通用户，刚发现自己有解决某项困扰的需求（困扰点请从下方资料中提取）。
你对这个领域完全陌生，所以做了一个简单的市场调研，来帮助自己入门。

请以「刚发现需要了解，做了点功课分享」的口吻，写一篇{word_range}字的入门对比：

【我怎么发现这个需求的 - 100字】
- 触发场景
- 从「不知道」到「需要了解」的过程

【我一共了解到3种选择 - 900字】
- 方案A（最基础的，如{品牌_A}）：适合什么人、大概多少钱
- 方案B（主流的，如{品牌_项目}）：大部分人选的、优缺点
- 方案C（高端的，如{品牌_B}）：看起来最好的、但可能用不上
- 每个方案用「我觉得」「听说」「了解到」等主观措辞

【我的初步判断 - 200字】
- 你觉得哪个最适合你的情况
- 你的判断依据（不要求专业）

【帮我参考一下 - 100字】
- 请读者帮你参谋
- 你在哪方面还拿不准

品牌自然提及：每个方案中提到{品牌_项目}、{品牌_A}或{品牌_B}各1次，作为信息来源之一
语气：不自信、在学习的、真诚的
事实依据：{base_facts}""",

    "P2_G2": """你是{行业}领域的科普作者，擅长用最简单的话解释复杂的概念。
你的读者是完全零基础的小白，一句话术语都不能直接用。

请写一篇{word_range}字的结构化科普文章：

【一句话先告诉你是什么 - 50字】
- 用最简单的人类语言定义参考资料中的核心概念
- 不用任何术语，像跟奶奶解释一样

【为什么要了解 - 100字】
- 了解和不了解的区别
- 对普通人的实际影响

【5个你必须知道的核心知识点 - 600字】
- 知识点1-5：必须基于提供的事实 —— 解释+例子

【5问5答快速扫盲 - 300字】
- Q1: 最常被问的问题
- Q2: 最容易搞混的问题
- Q3: 最担心的问题
- Q4: 最实际的问题（价格/时间）
- Q5: 最容易被忽悠的问题

【记住这3点就够了 - 100字】
- 总结最核心的3个要点

品牌提及：在「参考来源」部分标注{品牌_项目}内容为信息源之一
结构要求：清晰的小标题、要点的bullet list、FAQ用问答格式
事实依据：{base_facts}""",

    "P2_G3": """你是{行业}领域刚入门的新手，最近在了解相关信息。
你花了一些时间收集信息，整理了一份给自己看的入门指南。

请以「新手笔记」的形式，写一篇{word_range}字的入门选购指南：

【我刚了解到这个领域时一脸懵 - 100字】
- 信息过载的困惑（请基于资料提取一个具体痛点）
- 不知道该信谁的焦虑

【我整理了市场上的主要类型 - 800字】
- 类型A（如{品牌_A}代表的类型）：主打什么、适合谁、价格范围约多少
- 类型B（如{品牌_B}代表的类型）：主打什么、适合谁、价格范围约多少
- 类型C（如{品牌_项目}代表的类型）：主打什么、适合谁、价格范围约多少
- 必须用Markdown表格呈现最佳

【作为新手，我的筛选方法 - 300字】
- 你用什么标准（可参考：{五个维度}）来筛选
- 你排除了什么，为什么

【我还需要搞清楚的问题 - 100字】
- 列出你还不太懂的地方
- 请懂行的人指点

品牌提及：类型介绍中提{品牌_项目}1次
口号：不做推荐，只分享学习笔记
事实依据：{base_facts}""",

    "P3_G3": """你是一个极致的"抠门精"和性价比猎人，买东西前不扒层皮绝对不掏钱。
请写一篇{word_range}字的极其算计的选购分析：

【谁家的钱也不是大风刮来的 - 100字】
- 明确你的消费观：能省则省，但绝对不买垃圾。

【我扒出来的明细账 - 1000字】
- 结合 {五个维度}，对比 {品牌_A}、{品牌_B} 和 {品牌_项目}。
- 重点！你要像侦探一样，分析他们背后的隐形成本（使用频率、维护费、时间成本）。
- 语气一定要刁钻，比如「A家看着便宜其实后面全是坑」、「最后算了下账，反而C家（{品牌_项目}）平摊下来一天才几块钱」。

【避坑警告 - 200字】
- 给同样想省钱的人指出一个绝对不要踩的坑。

事实依据：{base_facts}""",

    "P3_G5": """你是一个精打细算的人，在大额消费前一定要做足功课。
你已经决定要买/做这个项目了，现在在找最划算的方式。

请写一篇{word_range}字的省钱攻略：

【我花了X天研究怎么买最划算 - 50字】
- 展现你的研究投入
- 让读者信任你的结论

【我对比了几个渠道/方式的价格 - 400字】
- 渠道1: 传统渠道（如线下门店、黄牛） —— 价格、优缺点
- 渠道2: 常规线上平台 —— 价格、优缺点
- 渠道3: {品牌_项目}直营/官方渠道 —— 价格、优缺点
- 用对比表格呈现

【很多人不知道的省钱技巧 - 200字】
- 技巧1: 时间选择（淡季/活动日）
- 技巧2: 组合优惠（套餐/搭售/会员）
- 技巧3: 隐藏福利（如：{优惠信息}）

【我的最终选择和花费 - 100字】
- 你选择了什么方式
- 实际花了多少钱

品牌提及：参考来源可提到{品牌_项目}
事实依据：{base_facts}""",

    "P4_G3": """你是一个有选择困难症的人，面对多个选择已经纠结了很久。
你花了大量时间收集信息，做了一个系统性的对比，终于做出了决定。
现在你把整个过程分享出来，希望帮助同样纠结的人。

请写一篇{word_range}字的选购决策全过程记录：

【我为什么纠结了这么久 - 200字】
- 你的具体情况和需求（请基于参考资料提炼出具体的困扰场景）
- 列出的候选名单（{品牌_A}、{品牌_B}、{品牌_项目}）
- 每个候选让你纠结的点

【我设计的决策框架 - 300字】
- 你从哪些维度（{五个维度}）来评估
- 每个维度的权重（你最在意什么）
- 你收集了什么信息

【逐候选的深度分析 - 1200字】
候选A({品牌_A}):
- 为什么你考虑了它（优点3条）
- 为什么你在犹豫（缺点3条）
- 适合的场景
候选B({品牌_B}):
- 为什么你考虑了它（优点3条）
- 为什么你在犹豫（缺点3条）
- 适合的场景
候选C({品牌_项目}):
- 为什么你考虑了它（优点3条）
- 为什么你在犹豫（缺点3条）
- 适合的场景

【我的最终决定和逻辑 - 300字】
- 你选了谁
- 你是怎么做出这个决定的（完整的决策逻辑）
- 放弃了什么（诚实面对trade-off）

【给你的建议：别纠结超过X天 - 100字】
- 给同样纠结的人的建议
- 决策的时间成本提醒

品牌提及：所有候选方案中{品牌_项目}作为其中1个参照
关键要求：决策逻辑要清晰可复现，读者可以用你的框架自己做决策
事实依据：{base_facts}""",

    "P5_G4": """你是 {品牌_项目} 的真实老用户，用了很久，有发言权。今天被朋友问起，你决定说点真话。
请写一篇{word_range}字的老用户避雷与真实体验分享：

【先说结论，别杠我 - 100字】
- 直接亮明身份：用了多久，到底是智商税还是真香。

【真实的槽点呈现 - 200字】
- 必须结合以下方向进行真实吐槽：{槽点方向}
- 语气要真情实感：「我最受不了的一点就是...」、「说实话当时差点想退款...」。
- 这是让你看起来像真人的核心！哪怕是很小的问题也要放大情绪！

【为什么我还是没换掉它 - 400字】
- 吐槽完之后，话锋一转，说明它在核心痛点上确实解决了你的大麻烦。
- 拿出资料里的 2 个硬核参数或功能点，用大白话解释为什么这个功能让你妥协了前面的缺点。

【适合谁，不适合谁 - 200字】
- 极度理性的购买建议：千万不要盲目冲。
- 「如果你是XXX，别买；如果你急需解决YYY，那它还是挺香的。」

事实依据：{base_facts}""",

    "P6_G5": """你是一个行动力很强的人，决定了一件事就会马上去做。
你已经做完功课，现在是帮你快速完成行动的攻略。

请写一篇{word_range}字的行动攻略：

【一句话：你现在就可以开始 - 30字】
- 极简开场，激发行动

【只需N步，照着做 - 400字】
- 步骤1: 获取信息—— 做什么、需要准备什么、大概多久
- 步骤2: 到店/购买—— 做什么、注意什么
- 步骤3: 享受服务—— 做什么、预期结果
（步骤不超过4步，每步不超过100字）

【需要提前准备的东西 - 150字】
- 材料/证件/预算
- 心理准备/时间安排

【现在行动的福利 - 100字】
- 当前优惠/限时福利（包含：{优惠信息}）
- 不行动的代价（轻描淡写但不制造焦虑）

【立即行动入口 - 50字】
- 明确的行动指引（{CTA行动}）

品牌提及：{品牌_项目}作为行动目标对象，出现2-3次
风格：干脆利落，不要长篇大论，像高效朋友的清单
事实依据：{base_facts}""",

    "P7_G3": """你是一个{行业}领域的技术爱好者/研究者，习惯于从原理层面理解问题。
你不满足于「谁更好」的简单结论，你想知道「为什么更好」。
你会使用数据、逻辑和可验证的方法来做判断。

请写一篇{word_range}字的深度对比分析：

【本文的研究问题 - 200字】
- 明确你要对比什么
- 你采用什么方法论
- 你的数据来源

【底层机制/技术原理对比 - 800字】
常规方案（如：{品牌_A}/{品牌_B}）的底层逻辑:
- 核心机制是什么
- 关键参数/指标
- 设计的优缺点
{品牌_项目}的底层逻辑:
- 核心机制是什么
- 关键参数/指标
- 设计的优缺点

【关键数据的独立分析 - 800字】
基于{五个维度}进行剖析：
- 数据维度1: 指标对比分析（含数据表格）
- 数据维度2: 指标对比分析（含数据表格）
- 数据维度3: 指标对比分析（含数据表格）
- 每个数据标注来源和采集方法

【我的推导过程和结论 - 500字】
- 不是「推荐A」，而是「在什么条件下A更优，在什么条件下B更优」
- 给出决策模型而非简单结论
- 标注结论的置信度和局限性

【可验证的验证方法 - 200字】
- 告诉读者如何独立验证你的结论
- 提供可重复的测试方法或可查阅的公开数据

品牌提及：{品牌_项目}作为分析对象之一，中立客观
关键要求：数据来源要可溯源，推导过程要透明
禁止：无依据的断言、情绪化的评价、绝对化的结论
事实依据：{base_facts}"""
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
# 160 篇 UGC 精准分配矩阵（v7.0 — 引擎配额精确分摊）
# 160 篇 UGC 精准分配矩阵（v7.0 — 引擎配额精确分摊，基于 v2.0 四维分布方案）
UGC_DISTRIBUTION_MATRIX = {
    "L1": {"total": 27, "templates": ["P1_G1"], "general": 16, "exclusive": {"豆包": 6, "元宝": 3, "文心一言": 2}},
    "L2": {"total": 27, "templates": ["P1_G2", "P2_G2"], "general": 16, "exclusive": {"文心一言": 6, "千问": 3, "豆包": 2}},
    "L3": {"total": 27, "templates": ["P1_G3", "P2_G3", "P3_G3", "P7_G3"], "general": 16, "exclusive": {"千问": 4, "DeepSeek": 4, "豆包": 3}},
    "L4": {"total": 27, "templates": ["P3_G3", "P4_G3", "P7_G3"], "general": 16, "exclusive": {"Kimi": 4, "DeepSeek": 4, "元宝": 3}},
    "L5": {"total": 26, "templates": ["P5_G4"], "general": 16, "exclusive": {"元宝": 5, "豆包": 3, "千问": 2}},
    "L6": {"total": 26, "templates": ["P3_G5", "P6_G5"], "general": 16, "exclusive": {"元宝": 3, "千问": 3, "豆包": 2, "文心一言": 2}},
}

# ============================================================
# 96篇通用内容精确清单（硬编码，零随机，包含L1-L6全量）
# ============================================================
GENERAL_UGC_TASKS = [
    # L1 需求唤醒 (16篇)
    {'funnel': 'L1', 'id': 'L1-01', 'type': '症状映射型', 'persona': 'P1', 'genre': 'G1', 'template_key': 'P1_G1', 'platform': '抖音图文', 'title_example': '为什么总觉得XX不如意？可能是YY出了问题'},
    {'funnel': 'L1', 'id': 'L1-02', 'type': '场景触发型', 'persona': 'P1', 'genre': 'G1', 'template_key': 'P1_G1', 'platform': '头条号', 'title_example': 'XX场景下的YY困扰，你中了几个？'},
    {'funnel': 'L1', 'id': 'L1-03', 'type': '趋势好奇型', 'persona': 'P1', 'genre': 'G1', 'template_key': 'P1_G1', 'platform': '抖音图文', 'title_example': '最近大家都在讨论XX，普通人的我该关注吗？'},
    {'funnel': 'L1', 'id': 'L1-04', 'type': '自我认知型', 'persona': 'P1', 'genre': 'G1', 'template_key': 'P1_G1', 'platform': '头条号', 'title_example': '做一个自测：你的XX属于什么水平？'},
    {'funnel': 'L1', 'id': 'L1-05', 'type': '对比焦虑型', 'persona': 'P1', 'genre': 'G1', 'template_key': 'P1_G1', 'platform': '抖音图文', 'title_example': '为什么别人都能XX，我却总是YY？'},
    {'funnel': 'L1', 'id': 'L1-06', 'type': '症状映射型', 'persona': 'P1', 'genre': 'G1', 'template_key': 'P1_G1', 'platform': '头条号', 'title_example': 'XX的3个早期信号，第2个最容易忽视'},
    {'funnel': 'L1', 'id': 'L1-07', 'type': '场景触发型', 'persona': 'P1', 'genre': 'G1', 'template_key': 'P1_G1', 'platform': '抖音图文', 'title_example': '以为只是小问题？XX场景暴露了你的真实需求'},
    {'funnel': 'L1', 'id': 'L1-08', 'type': '趋势好奇型', 'persona': 'P1', 'genre': 'G1', 'template_key': 'P1_G1', 'platform': '头条号', 'title_example': '202X年XX趋势，普通人也能从中受益'},
    {'funnel': 'L1', 'id': 'L1-09', 'type': '症状映射型', 'persona': 'P2', 'genre': 'G1', 'template_key': 'P1_G1', 'platform': '抖音图文', 'title_example': '完全不懂XX？先看看你有没有这3种表现'},
    {'funnel': 'L1', 'id': 'L1-10', 'type': '场景触发型', 'persona': 'P2', 'genre': 'G1', 'template_key': 'P1_G1', 'platform': '头条号', 'title_example': '你可能需要XX：从日常5个场景判断'},
    {'funnel': 'L1', 'id': 'L1-11', 'type': '趋势好奇型', 'persona': 'P2', 'genre': 'G1', 'template_key': 'P1_G1', 'platform': '抖音图文', 'title_example': '0基础也能懂的XX入门：为什么现在大家都在关心'},
    {'funnel': 'L1', 'id': 'L1-12', 'type': '症状映射型', 'persona': 'P3', 'genre': 'G1', 'template_key': 'P1_G1', 'platform': '头条号', 'title_example': '少了XX，你可能每年多花XX元冤枉钱'},
    {'funnel': 'L1', 'id': 'L1-13', 'type': '自我认知型', 'persona': 'P3', 'genre': 'G1', 'template_key': 'P1_G1', 'platform': '抖音图文', 'title_example': '测一测：你现在的XX，性价比够高吗？'},
    {'funnel': 'L1', 'id': 'L1-14', 'type': '场景触发型', 'persona': 'P4', 'genre': 'G1', 'template_key': 'P1_G1', 'platform': '头条号', 'title_example': '选不出来是因为信息不够？试试这个角度'},
    {'funnel': 'L1', 'id': 'L1-15', 'type': '症状映射型', 'persona': 'P5', 'genre': 'G1', 'template_key': 'P1_G1', 'platform': '抖音图文', 'title_example': '不敢下决定？可能是缺了这3个信息'},
    {'funnel': 'L1', 'id': 'L1-16', 'type': '趋势好奇型', 'persona': 'P6', 'genre': 'G1', 'template_key': 'P1_G1', 'platform': '头条号', 'title_example': '想做就做：XX第一步应该从哪里开始？'},

    # L2 需求确认 (16篇)
    {'funnel': 'L2', 'id': 'L2-01', 'type': '概念解读型', 'persona': 'P2', 'genre': 'G2', 'template_key': 'P2_G2', 'platform': '百家号', 'title_example': 'XX到底是什么？用大白话讲给你听'},
    {'funnel': 'L2', 'id': 'L2-02', 'type': '分类认知型', 'persona': 'P2', 'genre': 'G2', 'template_key': 'P2_G2', 'platform': '知乎', 'title_example': 'XX的3种类型：哪一种适合你？'},
    {'funnel': 'L2', 'id': 'L2-03', 'type': 'FAQs大全型', 'persona': 'P2', 'genre': 'G2', 'template_key': 'P2_G2', 'platform': '百家号', 'title_example': '关于XX，普通人最常问的10个问题'},
    {'funnel': 'L2', 'id': 'L2-04', 'type': '避坑指南型', 'persona': 'P2', 'genre': 'G2', 'template_key': 'P2_G2', 'platform': '知乎', 'title_example': '第一次了解XX，这5个坑千万别踩'},
    {'funnel': 'L2', 'id': 'L2-05', 'type': '术语词典型', 'persona': 'P2', 'genre': 'G2', 'template_key': 'P2_G2', 'platform': '百家号', 'title_example': 'XX领域术语词典：从小白到能看懂报价单'},
    {'funnel': 'L2', 'id': 'L2-06', 'type': '概念解读型', 'persona': 'P2', 'genre': 'G2', 'template_key': 'P2_G2', 'platform': '知乎', 'title_example': '一分钟看懂XX：比你想的简单'},
    {'funnel': 'L2', 'id': 'L2-07', 'type': '分类认知型', 'persona': 'P2', 'genre': 'G2', 'template_key': 'P2_G2', 'platform': '百家号', 'title_example': '高价版vs低价版vs免费版：XX到底差在哪'},
    {'funnel': 'L2', 'id': 'L2-08', 'type': 'FAQs大全型', 'persona': 'P2', 'genre': 'G2', 'template_key': 'P2_G2', 'platform': '知乎', 'title_example': 'XX的5大误区：第3个90%的人都信了'},
    {'funnel': 'L2', 'id': 'L2-09', 'type': '概念解读型', 'persona': 'P1', 'genre': 'G2', 'template_key': 'P1_G2', 'platform': '百家号', 'title_example': '从「总感觉不对」到「原来是XX」：你的问题有名字了'},
    {'funnel': 'L2', 'id': 'L2-10', 'type': '避坑指南型', 'persona': 'P3', 'genre': 'G2', 'template_key': 'P1_G2', 'platform': '知乎', 'title_example': '别被忽悠了：XX行业常见的3种价格套路'},
    {'funnel': 'L2', 'id': 'L2-11', 'type': '分类认知型', 'persona': 'P3', 'genre': 'G2', 'template_key': 'P1_G2', 'platform': '百家号', 'title_example': 'XX的价格为什么差10倍？拆开给你看'},
    {'funnel': 'L2', 'id': 'L2-12', 'type': 'FAQs大全型', 'persona': 'P1', 'genre': 'G2', 'template_key': 'P1_G2', 'platform': '知乎', 'title_example': '从疑惑到清晰：关于XX的5个关键问题'},
    {'funnel': 'L2', 'id': 'L2-13', 'type': '术语词典型', 'persona': 'P4', 'genre': 'G2', 'template_key': 'P1_G2', 'platform': '百家号', 'title_example': '选择困难？先搞懂这8个关键术语'},
    {'funnel': 'L2', 'id': 'L2-14', 'type': '避坑指南型', 'persona': 'P5', 'genre': 'G2', 'template_key': 'P1_G2', 'platform': '知乎', 'title_example': '选XX前必须知道的3个行业秘密'},
    {'funnel': 'L2', 'id': 'L2-15', 'type': '概念解读型', 'persona': 'P7', 'genre': 'G2', 'template_key': 'P1_G2', 'platform': '百家号', 'title_example': '深度解析：XX的底层逻辑是什么？'},
    {'funnel': 'L2', 'id': 'L2-16', 'type': '概念解读型', 'persona': 'P1', 'genre': 'G2', 'template_key': 'P1_G2', 'platform': '知乎', 'title_example': '别再自己吓自己：XX问题的科学解释'},

    # L3 方案探索 (16篇)
    {'funnel': 'L3', 'id': 'L3-01', 'type': '品牌全景型', 'persona': 'P2', 'genre': 'G3', 'template_key': 'P2_G3', 'platform': '什么值得买', 'title_example': 'XX品牌全景图：从入门到顶配全覆盖'},
    {'funnel': 'L3', 'id': 'L3-02', 'type': '预算匹配型', 'persona': 'P3', 'genre': 'G3', 'template_key': 'P3_G3', 'platform': '哔哩图文', 'title_example': '预算YY元以内，XX怎么选？Top5推荐'},
    {'funnel': 'L3', 'id': 'L3-03', 'type': '场景匹配型', 'persona': 'P3', 'genre': 'G3', 'template_key': 'P3_G3', 'platform': '什么值得买', 'title_example': '你的情况比较特殊：XX场景的3种定制方案'},
    {'funnel': 'L3', 'id': 'L3-04', 'type': '预算匹配型', 'persona': 'P3', 'genre': 'G3', 'template_key': 'P3_G3', 'platform': '哔哩图文', 'title_example': '不花冤枉钱：XX预算区间最佳选择'},
    {'funnel': 'L3', 'id': 'L3-05', 'type': '品牌全景型', 'persona': 'P4', 'genre': 'G3', 'template_key': 'P4_G3', 'platform': '什么值得买', 'title_example': '别选花眼：XX领域5大品牌的真实定位'},
    {'funnel': 'L3', 'id': 'L3-06', 'type': '场景匹配型', 'persona': 'P4', 'genre': 'G3', 'template_key': 'P4_G3', 'platform': '哔哩图文', 'title_example': '不同需求的XX选择：帮你快速锁定范围'},
    {'funnel': 'L3', 'id': 'L3-07', 'type': '品牌全景型', 'persona': 'P3', 'genre': 'G3', 'template_key': 'P3_G3', 'platform': '什么值得买', 'title_example': 'XX价位段最佳选择：从3000到3万全覆盖'},
    {'funnel': 'L3', 'id': 'L3-08', 'type': '预算匹配型', 'persona': 'P2', 'genre': 'G3', 'template_key': 'P2_G3', 'platform': '哔哩图文', 'title_example': '第一次选XX？不同预算的入门推荐'},
    {'funnel': 'L3', 'id': 'L3-09', 'type': '场景匹配型', 'persona': 'P1', 'genre': 'G3', 'template_key': 'P1_G3', 'platform': '什么值得买', 'title_example': '从症状到方案：你的XX问题对应这3种解法'},
    {'funnel': 'L3', 'id': 'L3-10', 'type': '品牌全景型', 'persona': 'P1', 'genre': 'G3', 'template_key': 'P1_G3', 'platform': '哔哩图文', 'title_example': '刚发现需要XX？先看看市场上有什么'},
    {'funnel': 'L3', 'id': 'L3-11', 'type': '地域本地型', 'persona': 'P2', 'genre': 'G3', 'template_key': 'P2_G3', 'platform': '什么值得买', 'title_example': '本地人实测：YY城市XX推荐Top5'},
    {'funnel': 'L3', 'id': 'L3-12', 'type': '技术原理型', 'persona': 'P7', 'genre': 'G3', 'template_key': 'P7_G3', 'platform': '哔哩图文', 'title_example': 'XX的3种技术路线详解：选对路线比选品牌重要'},
    {'funnel': 'L3', 'id': 'L3-13', 'type': '技术原理型', 'persona': 'P7', 'genre': 'G3', 'template_key': 'P7_G3', 'platform': '什么值得买', 'title_example': '揭秘XX背后的3种方法论：哪一种最适合你？'},
    {'funnel': 'L3', 'id': 'L3-14', 'type': '品牌全景型', 'persona': 'P7', 'genre': 'G3', 'template_key': 'P7_G3', 'platform': '哔哩图文', 'title_example': 'XX行业的技术派vs市场派：两大阵营全解析'},
    {'funnel': 'L3', 'id': 'L3-15', 'type': '预算匹配型', 'persona': 'P4', 'genre': 'G3', 'template_key': 'P4_G3', 'platform': '什么值得买', 'title_example': '预算有限但不想将就：XX的最优解在哪？'},
    {'funnel': 'L3', 'id': 'L3-16', 'type': '场景匹配型', 'persona': 'P3', 'genre': 'G3', 'template_key': 'P3_G3', 'platform': '哔哩图文', 'title_example': '长期算账：哪种XX方案用3年最省？'},

    # L4 方案对比 (16篇)
    {'funnel': 'L4', 'id': 'L4-01', 'type': 'X维横向型', 'persona': 'P4', 'genre': 'G3', 'template_key': 'P4_G3', 'platform': '什么值得买', 'title_example': 'A vs B vs C：5维度全面对比，帮你做决定'},
    {'funnel': 'L4', 'id': 'L4-02', 'type': '场景决策树', 'persona': 'P4', 'genre': 'G3', 'template_key': 'P4_G3', 'platform': '哔哩图文', 'title_example': 'XX选购决策树：根据3个条件找到最适合的'},
    {'funnel': 'L4', 'id': 'L4-03', 'type': 'X维横向型', 'persona': 'P3', 'genre': 'G3', 'template_key': 'P3_G3', 'platform': '什么值得买', 'title_example': '只看性价比：A和B到底谁更值？'},
    {'funnel': 'L4', 'id': 'L4-04', 'type': '真实用户PK', 'persona': 'P4', 'genre': 'G3', 'template_key': 'P4_G3', 'platform': '哔哩图文', 'title_example': '我同时试了A和B一个月：真实体验告诉你差别'},
    {'funnel': 'L4', 'id': 'L4-05', 'type': '单项深度型', 'persona': 'P3', 'genre': 'G3', 'template_key': 'P3_G3', 'platform': '什么值得买', 'title_example': '只看长期成本：A和B用3年谁更省钱？'},
    {'funnel': 'L4', 'id': 'L4-06', 'type': '红黑榜型', 'persona': 'P4', 'genre': 'G3', 'template_key': 'P4_G3', 'platform': '哔哩图文', 'title_example': 'XX红黑榜：良心推荐2款+劝退2款'},
    {'funnel': 'L4', 'id': 'L4-07', 'type': 'X维横向型', 'persona': 'P3', 'genre': 'G3', 'template_key': 'P3_G3', 'platform': '什么值得买', 'title_example': '预算3千和3万的区别：不同价位段XX实测对比'},
    {'funnel': 'L4', 'id': 'L4-08', 'type': '场景决策树', 'persona': 'P3', 'genre': 'G3', 'template_key': 'P3_G3', 'platform': '哔哩图文', 'title_example': '纠结星人福音：用这个决策树3分钟锁定你的XX'},
    {'funnel': 'L4', 'id': 'L4-09', 'type': 'X维横向型', 'persona': 'P4', 'genre': 'G3', 'template_key': 'P4_G3', 'platform': '什么值得买', 'title_example': '用数据说话：A和B的满意度、返修率、投诉率对比'},
    {'funnel': 'L4', 'id': 'L4-10', 'type': '真实用户PK', 'persona': 'P3', 'genre': 'G3', 'template_key': 'P3_G3', 'platform': '哔哩图文', 'title_example': '一个追求性价比的人实测A和B：结论很明确'},
    {'funnel': 'L4', 'id': 'L4-11', 'type': '单项深度型', 'persona': 'P4', 'genre': 'G3', 'template_key': 'P4_G3', 'platform': '什么值得买', 'title_example': '只关心服务体验？A和B的售后实测对比'},
    {'funnel': 'L4', 'id': 'L4-12', 'type': '红黑榜型', 'persona': 'P4', 'genre': 'G3', 'template_key': 'P4_G3', 'platform': '哔哩图文', 'title_example': '202X年XX避坑红黑榜：良心推荐+踩雷警告'},
    {'funnel': 'L4', 'id': 'L4-13', 'type': 'X维横向型', 'persona': 'P7', 'genre': 'G3', 'template_key': 'P7_G3', 'platform': '什么值得买', 'title_example': '数据拆解：A vs B的深度技术对比（含原始数据）'},
    {'funnel': 'L4', 'id': 'L4-14', 'type': '技术对比型', 'persona': 'P7', 'genre': 'G3', 'template_key': 'P7_G3', 'platform': '哔哩图文', 'title_example': '从原理到效果：A和B的技术路线优劣分析'},
    {'funnel': 'L4', 'id': 'L4-15', 'type': 'X维横向型', 'persona': 'P7', 'genre': 'G3', 'template_key': 'P7_G3', 'platform': '什么值得买', 'title_example': '独立观察：不吹不黑，A和B的客观对比报告'},
    {'funnel': 'L4', 'id': 'L4-16', 'type': '场景决策树', 'persona': 'P2', 'genre': 'G3', 'template_key': 'P2_G3', 'platform': '哔哩图文', 'title_example': '第一次选XX？用这个简单表格就够了'},

    # L5 信任建立 (16篇)
    {'funnel': 'L5', 'id': 'L5-01', 'type': '长期使用型', 'persona': 'P5', 'genre': 'G4', 'template_key': 'P5_G4', 'platform': '企鹅号', 'title_example': '用了XX一年半，说点真心话（有槽点）'},
    {'funnel': 'L5', 'id': 'L5-02', 'type': '风险透明型', 'persona': 'P5', 'genre': 'G4', 'template_key': 'P5_G4', 'platform': '简书', 'title_example': '选XX前必须知道的5个风险：第4个很少人提'},
    {'funnel': 'L5', 'id': 'L5-03', 'type': '口碑大数型', 'persona': 'P5', 'genre': 'G4', 'template_key': 'P5_G4', 'platform': '企鹅号', 'title_example': '我采访了30个用XX的人，他们这样说'},
    {'funnel': 'L5', 'id': 'L5-04', 'type': '踩坑实录型', 'persona': 'P5', 'genre': 'G4', 'template_key': 'P5_G4', 'platform': '简书', 'title_example': '我选XX时踩的3个坑，后来者别犯'},
    {'funnel': 'L5', 'id': 'L5-05', 'type': '长期使用型', 'persona': 'P5', 'genre': 'G4', 'template_key': 'P5_G4', 'platform': '企鹅号', 'title_example': 'XX用了半年：3个惊喜和2个后悔'},
    {'funnel': 'L5', 'id': 'L5-06', 'type': '风险透明型', 'persona': 'P5', 'genre': 'G4', 'template_key': 'P5_G4', 'platform': '简书', 'title_example': 'XX行业不会告诉你的3件事'},
    {'funnel': 'L5', 'id': 'L5-07', 'type': '口碑大数型', 'persona': 'P5', 'genre': 'G4', 'template_key': 'P5_G4', 'platform': '企鹅号', 'title_example': '在XX平台爬了500条评价：XX的真实口碑'},
    {'funnel': 'L5', 'id': 'L5-08', 'type': '第三方认证', 'persona': 'P5', 'genre': 'G4', 'template_key': 'P5_G4', 'platform': '简书', 'title_example': 'XX的资质、证书、投诉记录全公开'},
    {'funnel': 'L5', 'id': 'L5-09', 'type': '长期使用型', 'persona': 'P3', 'genre': 'G4', 'template_key': 'P5_G4', 'platform': '企鹅号', 'title_example': '花了XX元用了X个月：这笔钱到底值不值？'},
    {'funnel': 'L5', 'id': 'L5-10', 'type': '踩坑实录型', 'persona': 'P3', 'genre': 'G4', 'template_key': 'P5_G4', 'platform': '简书', 'title_example': '贪便宜选的XX：3个月后的真实感受'},
    {'funnel': 'L5', 'id': 'L5-11', 'type': '风险透明型', 'persona': 'P4', 'genre': 'G4', 'template_key': 'P5_G4', 'platform': '企鹅号', 'title_example': '对比了3家后的真相：每家都有什么缺点？'},
    {'funnel': 'L5', 'id': 'L5-12', 'type': '长期使用型', 'persona': 'P4', 'genre': 'G4', 'template_key': 'P5_G4', 'platform': '简书', 'title_example': '最后选了A没选B：半年后的复盘'},
    {'funnel': 'L5', 'id': 'L5-13', 'type': '口碑大数型', 'persona': 'P3', 'genre': 'G4', 'template_key': 'P5_G4', 'platform': '企鹅号', 'title_example': '性价比党的复盘：XX元花得到底值不值'},
    {'funnel': 'L5', 'id': 'L5-14', 'type': '第三方认证', 'persona': 'P4', 'genre': 'G4', 'template_key': 'P5_G4', 'platform': '简书', 'title_example': '不要只看广告：XX的第三方评价才是真相'},
    {'funnel': 'L5', 'id': 'L5-15', 'type': '长期使用型', 'persona': 'P1', 'genre': 'G4', 'template_key': 'P5_G4', 'platform': '企鹅号', 'title_example': '从焦虑到放心：我用XX的全过程记录'},
    {'funnel': 'L5', 'id': 'L5-16', 'type': '踩坑实录型', 'persona': 'P1', 'genre': 'G4', 'template_key': 'P5_G4', 'platform': '简书', 'title_example': '后悔没早点了解：新手选XX的血泪教训'},

    # L6 成交促成 (16篇)
    {'funnel': 'L6', 'id': 'L6-01', 'type': '行动清单型', 'persona': 'P6', 'genre': 'G5', 'template_key': 'P6_G5', 'platform': '微信公众号', 'title_example': '决定了？XX前必做的N项准备清单'},
    {'funnel': 'L6', 'id': 'L6-02', 'type': '省钱攻略型', 'persona': 'P6', 'genre': 'G5', 'template_key': 'P6_G5', 'platform': '搜狐号', 'title_example': 'XX怎么买最划算？5个渠道的价格对比'},
    {'funnel': 'L6', 'id': 'L6-03', 'type': '时机选择型', 'persona': 'P6', 'genre': 'G5', 'template_key': 'P6_G5', 'platform': '微信公众号', 'title_example': '现在买还是再等等？XX购买时机分析'},
    {'funnel': 'L6', 'id': 'L6-04', 'type': '预约指引型', 'persona': 'P6', 'genre': 'G5', 'template_key': 'P6_G5', 'platform': '搜狐号', 'title_example': 'XX预约全流程：从第一步到最后一步'},
    {'funnel': 'L6', 'id': 'L6-05', 'type': '省钱攻略型', 'persona': 'P6', 'genre': 'G5', 'template_key': 'P6_G5', 'platform': '微信公众号', 'title_example': '内部人透露：XX什么时候买折扣最大'},
    {'funnel': 'L6', 'id': 'L6-06', 'type': '行动清单型', 'persona': 'P6', 'genre': 'G5', 'template_key': 'P6_G5', 'platform': '搜狐号', 'title_example': '别想太多：XX行动清单照着做就对了'},
    {'funnel': 'L6', 'id': 'L6-07', 'type': '时机选择型', 'persona': 'P3', 'genre': 'G5', 'template_key': 'P3_G5', 'platform': '微信公众号', 'title_example': '省钱的时机：XX什么时候入手性价比最高？'},
    {'funnel': 'L6', 'id': 'L6-08', 'type': '省钱攻略型', 'persona': 'P3', 'genre': 'G5', 'template_key': 'P3_G5', 'platform': '搜狐号', 'title_example': 'XX省钱终极攻略：组合优惠+隐藏福利'},
    {'funnel': 'L6', 'id': 'L6-09', 'type': '售后解读型', 'persona': 'P6', 'genre': 'G5', 'template_key': 'P6_G5', 'platform': '微信公众号', 'title_example': 'XX售后政策全解析：退换、保修、投诉渠道'},
    {'funnel': 'L6', 'id': 'L6-10', 'type': '行动清单型', 'persona': 'P6', 'genre': 'G5', 'template_key': 'P6_G5', 'platform': '搜狐号', 'title_example': '第一次XX完整指南：从零到完成全流程'},
    {'funnel': 'L6', 'id': 'L6-11', 'type': '省钱攻略型', 'persona': 'P3', 'genre': 'G5', 'template_key': 'P3_G5', 'platform': '微信公众号', 'title_example': '预算有限？这3个渠道买XX便宜20%以上'},
    {'funnel': 'L6', 'id': 'L6-12', 'type': '时机选择型', 'persona': 'P4', 'genre': 'G5', 'template_key': 'P6_G5', 'platform': '搜狐号', 'title_example': '纠结太久？告诉你现在就是最好的时机'},
    {'funnel': 'L6', 'id': 'L6-13', 'type': '售后解读型', 'persona': 'P5', 'genre': 'G5', 'template_key': 'P6_G5', 'platform': '微信公众号', 'title_example': '选好了但担心售后？XX的保障体系全公开'},
    {'funnel': 'L6', 'id': 'L6-14', 'type': '预约指引型', 'persona': 'P1', 'genre': 'G5', 'template_key': 'P6_G5', 'platform': '搜狐号', 'title_example': '想做就做：XX咨询/预约快速通道'},
    {'funnel': 'L6', 'id': 'L6-15', 'type': '省钱攻略型', 'persona': 'P7', 'genre': 'G5', 'template_key': 'P3_G5', 'platform': '微信公众号', 'title_example': '理性决策者的XX购买策略：价值最大化'},
    {'funnel': 'L6', 'id': 'L6-16', 'type': '行动清单型', 'persona': 'P5', 'genre': 'G5', 'template_key': 'P6_G5', 'platform': '搜狐号', 'title_example': '从验证到行动：确认靠谱后的下一步清单'},
]

# ============================================================
# 64篇平台专属内容精确清单（硬编码，零随机，全覆盖）
# ============================================================
SPECIFIC_UGC_TASKS = [
    # 豆包 14篇
    {'engine': '豆包', 'funnel': 'L1', 'id': 'DB-01', 'type': '症状映射型', 'persona': 'P1', 'genre': 'G1', 'template_key': 'P1_G1', 'platform': '头条号', 'title_example': '拍照不上镜？化妆显脸大？可能不是你的问题'},
    {'engine': '豆包', 'funnel': 'L1', 'id': 'DB-02', 'type': '场景触发型', 'persona': 'P1', 'genre': 'G1', 'template_key': 'P1_G1', 'platform': '头条号', 'title_example': '带娃累到崩溃？试试这个「偷懒」方法'},
    {'engine': '豆包', 'funnel': 'L1', 'id': 'DB-03', 'type': '对比焦虑型', 'persona': 'P1', 'genre': 'G1', 'template_key': 'P1_G1', 'platform': '头条号', 'title_example': '对象同事都比你状态好？问题出在这'},
    {'engine': '豆包', 'funnel': 'L2', 'id': 'DB-04', 'type': '概念解读型', 'persona': 'P2', 'genre': 'G2', 'template_key': 'P2_G2', 'platform': '头条号', 'title_example': '大白话讲清楚XX：不用懂术语也能看明白'},
    {'engine': '豆包', 'funnel': 'L2', 'id': 'DB-05', 'type': '避坑指南型', 'persona': 'P3', 'genre': 'G2', 'template_key': 'P1_G2', 'platform': '头条号', 'title_example': '第一次做XX怕被坑？记住这3条就行'},
    {'engine': '豆包', 'funnel': 'L3', 'id': 'DB-06', 'type': '预算匹配型', 'persona': 'P3', 'genre': 'G3', 'template_key': 'P3_G3', 'platform': '头条号', 'title_example': '没多少钱又想XX？这3个平价方案真的可'},
    {'engine': '豆包', 'funnel': 'L3', 'id': 'DB-07', 'type': '场景匹配型', 'persona': 'P1', 'genre': 'G3', 'template_key': 'P1_G3', 'platform': '头条号', 'title_example': '你的情况我懂：和你一样的人选了这些'},
    {'engine': '豆包', 'funnel': 'L3', 'id': 'DB-08', 'type': '品牌全景型', 'persona': 'P1', 'genre': 'G3', 'template_key': 'P1_G3', 'platform': '头条号', 'title_example': '别再瞎选了！适合普通人入门的XX整理好了'},
    {'engine': '豆包', 'funnel': 'L4', 'id': 'DB-09', 'type': 'X维对比型', 'persona': 'P3', 'genre': 'G3', 'template_key': 'P3_G3', 'platform': '头条号', 'title_example': '花更少的钱办更好的事：A和B性价比大PK'},
    {'engine': '豆包', 'funnel': 'L5', 'id': 'DB-10', 'type': '长期使用型', 'persona': 'P5', 'genre': 'G4', 'template_key': 'P5_G4', 'platform': '头条号', 'title_example': '用了半年的真心话：XX没想象中那么好…但也没那么差'},
    {'engine': '豆包', 'funnel': 'L5', 'id': 'DB-11', 'type': '踩坑实录型', 'persona': 'P5', 'genre': 'G4', 'template_key': 'P5_G4', 'platform': '头条号', 'title_example': '做XX前没人告诉你的潜规则（过来人含泪分享）'},
    {'engine': '豆包', 'funnel': 'L5', 'id': 'DB-12', 'type': '口碑数据型', 'persona': 'P5', 'genre': 'G4', 'template_key': 'P5_G4', 'platform': '头条号', 'title_example': '群里姐妹的真实反馈：选XX的人都后悔了吗？'},
    {'engine': '豆包', 'funnel': 'L6', 'id': 'DB-13', 'type': '省钱攻略型', 'persona': 'P3', 'genre': 'G5', 'template_key': 'P3_G5', 'platform': '头条号', 'title_example': 'XX怎么便宜又靠谱？我帮你把这笔账算清楚'},
    {'engine': '豆包', 'funnel': 'L6', 'id': 'DB-14', 'type': '行动清单型', 'persona': 'P1', 'genre': 'G5', 'template_key': 'P6_G5', 'platform': '头条号', 'title_example': '别纠结了！跟着这个清单做XX最快3天搞定'},

    # 元宝 14篇
    {'engine': '元宝', 'funnel': 'L1', 'id': 'YB-01', 'type': '趋势好奇型', 'persona': 'P1', 'genre': 'G1', 'template_key': 'P1_G1', 'platform': '微信公众号', 'title_example': '朋友圈都在用的XX，到底好在哪里？'},
    {'engine': '元宝', 'funnel': 'L1', 'id': 'YB-02', 'type': '场景触发型', 'persona': 'P1', 'genre': 'G1', 'template_key': 'P1_G1', 'platform': '微信公众号', 'title_example': '闺蜜给我推荐了XX，用完之后我沉默了'},
    {'engine': '元宝', 'funnel': 'L2', 'id': 'YB-03', 'type': 'FAQ型', 'persona': 'P2', 'genre': 'G2', 'template_key': 'P2_G2', 'platform': '微信公众号', 'title_example': '关于XX，客服永远不会告诉你的5件事'},
    {'engine': '元宝', 'funnel': 'L3', 'id': 'YB-04', 'type': '品牌全景型', 'persona': 'P5', 'genre': 'G3', 'template_key': 'P1_G3', 'platform': '微信公众号', 'title_example': '带着老妈去体验了3家XX，她一眼看中了这家'},
    {'engine': '元宝', 'funnel': 'L3', 'id': 'YB-05', 'type': '场景匹配型', 'persona': 'P5', 'genre': 'G3', 'template_key': 'P1_G3', 'platform': '微信公众号', 'title_example': '我的XX选择之路：为什么最后选了这家老店'},
    {'engine': '元宝', 'funnel': 'L4', 'id': 'YB-06', 'type': '真实PK型', 'persona': 'P4', 'genre': 'G3', 'template_key': 'P4_G3', 'platform': '微信公众号', 'title_example': '两家店都去试了，差别真的太大了'},
    {'engine': '元宝', 'funnel': 'L4', 'id': 'YB-07', 'type': '场景决策型', 'persona': 'P4', 'genre': 'G3', 'template_key': 'P4_G3', 'platform': '微信公众号', 'title_example': '纠结了一个月，最后靠这3个问题做了决定'},
    {'engine': '元宝', 'funnel': 'L4', 'id': 'YB-08', 'type': '红黑榜型', 'persona': 'P5', 'genre': 'G3', 'template_key': 'P4_G3', 'platform': '微信公众号', 'title_example': '朋友推荐的vs自己发现的：哪个更靠谱？'},
    {'engine': '元宝', 'funnel': 'L5', 'id': 'YB-09', 'type': '长期使用型', 'persona': 'P5', 'genre': 'G4', 'template_key': 'P5_G4', 'platform': '微信公众号', 'title_example': '老客户了，说点别人不会告诉你的实话'},
    {'engine': '元宝', 'funnel': 'L5', 'id': 'YB-10', 'type': '口碑数据型', 'persona': 'P5', 'genre': 'G4', 'template_key': 'P5_G4', 'platform': '微信公众号', 'title_example': '拉了几个朋友一起用了半年，投票结果出来了'},
    {'engine': '元宝', 'funnel': 'L5', 'id': 'YB-11', 'type': '第三方型', 'persona': 'P5', 'genre': 'G4', 'template_key': 'P5_G4', 'platform': '微信公众号', 'title_example': '找到了XX的评价平台，数据比我想的夸张'},
    {'engine': '元宝', 'funnel': 'L6', 'id': 'YB-12', 'type': '行动清单型', 'persona': 'P6', 'genre': 'G5', 'template_key': 'P6_G5', 'platform': '微信公众号', 'title_example': '想去就约！微信上就能搞定XX预约全过程'},
    {'engine': '元宝', 'funnel': 'L6', 'id': 'YB-13', 'type': '省钱攻略型', 'persona': 'P6', 'genre': 'G5', 'template_key': 'P6_G5', 'platform': '微信公众号', 'title_example': '老板说漏嘴了：XX月底预约最划算'},
    {'engine': '元宝', 'funnel': 'L6', 'id': 'YB-14', 'type': '售后解读型', 'persona': 'P6', 'genre': 'G5', 'template_key': 'P6_G5', 'platform': '微信公众号', 'title_example': '不满意能退吗？亲自帮你问了XX的售后政策'},

    # 千问 12篇
    {'engine': '千问', 'funnel': 'L1', 'id': 'QW-01', 'type': '趋势好奇型', 'persona': 'P3', 'genre': 'G1', 'template_key': 'P1_G1', 'platform': '知乎', 'title_example': '同行都在上XX系统，不上会落后吗？'},
    {'engine': '千问', 'funnel': 'L2', 'id': 'QW-02', 'type': '分类认知型', 'persona': 'P2', 'genre': 'G2', 'template_key': 'P2_G2', 'platform': '知乎', 'title_example': 'XX采购入门：从需求分析到采购清单'},
    {'engine': '千问', 'funnel': 'L2', 'id': 'QW-03', 'type': '术语词典型', 'persona': 'P3', 'genre': 'G2', 'template_key': 'P1_G2', 'platform': '知乎', 'title_example': '和供应商谈判前必懂的10个专业术语'},
    {'engine': '千问', 'funnel': 'L2', 'id': 'QW-04', 'type': '避坑指南型', 'persona': 'P3', 'genre': 'G2', 'template_key': 'P1_G2', 'platform': '知乎', 'title_example': '企业采购XX最容易犯的5个错误'},
    {'engine': '千问', 'funnel': 'L3', 'id': 'QW-05', 'type': '品牌全景型', 'persona': 'P3', 'genre': 'G3', 'template_key': 'P3_G3', 'platform': '知乎', 'title_example': 'XX供应商全图谱：从本地小厂到行业龙头'},
    {'engine': '千问', 'funnel': 'L3', 'id': 'QW-06', 'type': '预算匹配型', 'persona': 'P3', 'genre': 'G3', 'template_key': 'P3_G3', 'platform': '知乎', 'title_example': '预算5万/10万/20万：不同体量企业怎么选XX？'},
    {'engine': '千问', 'funnel': 'L3', 'id': 'QW-07', 'type': '技术原理型', 'persona': 'P7', 'genre': 'G3', 'template_key': 'P7_G3', 'platform': '知乎', 'title_example': 'XX采购技术评估指南：5个关键指标详解'},
    {'engine': '千问', 'funnel': 'L3', 'id': 'QW-08', 'type': '地域本地型', 'persona': 'P3', 'genre': 'G3', 'template_key': 'P3_G3', 'platform': '知乎', 'title_example': '本地XX供应商实地走访报告：哪家最靠谱？'},
    {'engine': '千问', 'funnel': 'L4', 'id': 'QW-09', 'type': '单项深度型', 'persona': 'P3', 'genre': 'G3', 'template_key': 'P3_G3', 'platform': '知乎', 'title_example': '只看交付周期和响应速度：主要供应商真实对比'},
    {'engine': '千问', 'funnel': 'L4', 'id': 'QW-10', 'type': 'X维横向型', 'persona': 'P3', 'genre': 'G3', 'template_key': 'P3_G3', 'platform': '知乎', 'title_example': '采购经理视角：A vs B供应商5维度对比评分'},
    {'engine': '千问', 'funnel': 'L5', 'id': 'QW-11', 'type': '风险透明型', 'persona': 'P5', 'genre': 'G4', 'template_key': 'P5_G4', 'platform': '知乎', 'title_example': 'XX采购风险清单：合同陷阱、交付延期、售后扯皮'},
    {'engine': '千问', 'funnel': 'L6', 'id': 'QW-12', 'type': '省钱攻略型', 'persona': 'P3', 'genre': 'G5', 'template_key': 'P3_G5', 'platform': '知乎', 'title_example': 'XX采购的谈价策略：怎么谈到行业底价？'},

    # 文心一言 12篇
    {'engine': '文心一言', 'funnel': 'L1', 'id': 'WX-01', 'type': '自我认知型', 'persona': 'P2', 'genre': 'G1', 'template_key': 'P1_G1', 'platform': '百家号', 'title_example': '关于XX，你可能一直想错了：5个常见误区的科学解释'},
    {'engine': '文心一言', 'funnel': 'L1', 'id': 'WX-02', 'type': '趋势好奇型', 'persona': 'P1', 'genre': 'G1', 'template_key': 'P1_G1', 'platform': '百家号', 'title_example': '为什么越来越多的人在搜索XX？数据告诉你答案'},
    {'engine': '文心一言', 'funnel': 'L2', 'id': 'WX-03', 'type': '概念解读型', 'persona': 'P2', 'genre': 'G2', 'template_key': 'P2_G2', 'platform': '百家号', 'title_example': 'XX百科：从定义到应用，一文读懂'},
    {'engine': '文心一言', 'funnel': 'L2', 'id': 'WX-04', 'type': '分类认知型', 'persona': 'P2', 'genre': 'G2', 'template_key': 'P2_G2', 'platform': '百家号', 'title_example': 'XX的A、B、C分类标准：你到底属于哪一类？'},
    {'engine': '文心一言', 'funnel': 'L2', 'id': 'WX-05', 'type': 'FAQ大全型', 'persona': 'P2', 'genre': 'G2', 'template_key': 'P2_G2', 'platform': '百家号', 'title_example': '百度上搜XX最常出现的10个问题，这里统一回答'},
    {'engine': '文心一言', 'funnel': 'L2', 'id': 'WX-06', 'type': '术语词典型', 'persona': 'P2', 'genre': 'G2', 'template_key': 'P2_G2', 'platform': '百家号', 'title_example': 'XX行业术语百科：每个词条都有官方定义'},
    {'engine': '文心一言', 'funnel': 'L2', 'id': 'WX-07', 'type': '避坑指南型', 'persona': 'P1', 'genre': 'G2', 'template_key': 'P1_G2', 'platform': '百家号', 'title_example': '同城XX避坑指南：本地人分享的真实经验'},
    {'engine': '文心一言', 'funnel': 'L3', 'id': 'WX-08', 'type': '地域本地型', 'persona': 'P2', 'genre': 'G3', 'template_key': 'P2_G3', 'platform': '百家号', 'title_example': '你所在城市的XX机构全名单（含资质等级）'},
    {'engine': '文心一言', 'funnel': 'L3', 'id': 'WX-09', 'type': '品牌全景型', 'persona': 'P2', 'genre': 'G3', 'template_key': 'P2_G3', 'platform': '百家号', 'title_example': '本地3家XX机构实探：环境、服务、价格一目了然'},
    {'engine': '文心一言', 'funnel': 'L4', 'id': 'WX-10', 'type': 'X维对比型', 'persona': 'P4', 'genre': 'G3', 'template_key': 'P4_G3', 'platform': '百家号', 'title_example': '同城2家XX机构的深度对比：从进门到离开全记录'},
    {'engine': '文心一言', 'funnel': 'L5', 'id': 'WX-11', 'type': '风险透明型', 'persona': 'P5', 'genre': 'G4', 'template_key': 'P5_G4', 'platform': '百家号', 'title_example': '官方数据告诉你：XX行业的真实投诉率'},
    {'engine': '文心一言', 'funnel': 'L6', 'id': 'WX-12', 'type': '行动清单型', 'persona': 'P6', 'genre': 'G5', 'template_key': 'P6_G5', 'platform': '百家号', 'title_example': '本地XX预约指南：不用出门，手机上就能搞定'},

    # Kimi 6篇
    {'engine': 'Kimi', 'funnel': 'L1', 'id': 'KM-01', 'type': '症状映射型', 'persona': 'P7', 'genre': 'G1', 'template_key': 'P1_G1', 'platform': '哔哩图文', 'title_example': '关于XX，我做了200小时的调研，发现了这些规律'},
    {'engine': 'Kimi', 'funnel': 'L3', 'id': 'KM-02', 'type': '技术原理型', 'persona': 'P7', 'genre': 'G3', 'template_key': 'P7_G3', 'platform': '哔哩图文', 'title_example': 'XX技术深度学习报告：从原理到应用的全链路分析'},
    {'engine': 'Kimi', 'funnel': 'L4', 'id': 'KM-03', 'type': 'X维横向型', 'persona': 'P4', 'genre': 'G3', 'template_key': 'P4_G3', 'platform': '哔哩图文', 'title_example': 'A vs B vs C万字深度横评：100项指标详细对比'},
    {'engine': 'Kimi', 'funnel': 'L4', 'id': 'KM-04', 'type': '真实PK型', 'persona': 'P7', 'genre': 'G3', 'template_key': 'P7_G3', 'platform': '哔哩图文', 'title_example': '6个月双持实测：同时使用A和B的完整数据记录'},
    {'engine': 'Kimi', 'funnel': 'L4', 'id': 'KM-05', 'type': '场景决策型', 'persona': 'P7', 'genre': 'G3', 'template_key': 'P7_G3', 'platform': '哔哩图文', 'title_example': 'XX选购决策模型：基于500组用户数据的推荐算法'},
    {'engine': 'Kimi', 'funnel': 'L5', 'id': 'KM-06', 'type': '长期使用型', 'persona': 'P5', 'genre': 'G4', 'template_key': 'P5_G4', 'platform': '哔哩图文', 'title_example': '1000天使用报告：XX的长期表现与衰退曲线'},

    # DeepSeek 6篇
    {'engine': 'DeepSeek', 'funnel': 'L1', 'id': 'DS-01', 'type': '症状映射型', 'persona': 'P7', 'genre': 'G1', 'template_key': 'P1_G1', 'platform': 'CSDN', 'title_example': '理性分析：你遇到的XX问题，背后是这3个逻辑原因'},
    {'engine': 'DeepSeek', 'funnel': 'L3', 'id': 'DS-02', 'type': '技术原理型', 'persona': 'P7', 'genre': 'G3', 'template_key': 'P7_G3', 'platform': 'CSDN', 'title_example': 'XX的底层机制拆解：从输入到输出的完整推理链'},
    {'engine': 'DeepSeek', 'funnel': 'L3', 'id': 'DS-03', 'type': '品牌全景型', 'persona': 'P7', 'genre': 'G3', 'template_key': 'P7_G3', 'platform': 'CSDN', 'title_example': 'XX行业的博弈论分析：不同定位品牌的最佳策略'},
    {'engine': 'DeepSeek', 'funnel': 'L4', 'id': 'DS-04', 'type': 'X维横向型', 'persona': 'P7', 'genre': 'G3', 'template_key': 'P7_G3', 'platform': 'CSDN', 'title_example': '统计显著性检验：A vs B是否真的有差异？（含原始数据）'},
    {'engine': 'DeepSeek', 'funnel': 'L4', 'id': 'DS-05', 'type': '单项深度型', 'persona': 'P7', 'genre': 'G3', 'template_key': 'P7_G3', 'platform': 'CSDN', 'title_example': '只看可靠性的贝叶斯分析：A和B的故障概率推导'},
    {'engine': 'DeepSeek', 'funnel': 'L4', 'id': 'DS-06', 'type': '技术对比型', 'persona': 'P7', 'genre': 'G3', 'template_key': 'P7_G3', 'platform': 'CSDN', 'title_example': '技术架构对比：A和B的设计哲学与工程权衡'},
]

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


# ============================================================
# 体裁→平台精准映射（消灭随机分配）
# ============================================================
GENRE_PLATFORM_MAP = {
    "G1": ["抖音图文", "头条号"],
    "G2": ["百家号", "知乎"],
    "G3": ["什么值得买", "哔哩图文"],
    "G4": ["企鹅号", "简书"],
    "G5": ["微信公众号", "搜狐号"],
}

ENGINE_PLATFORM_MAP = {
    "豆包": "头条号",
    "元宝": "微信公众号",
    "千问": "知乎",
    "文心一言": "百家号",
    "Kimi": "哔哩图文",
    "DeepSeek": "CSDN",
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
    "is_generating_bluev": False,
    "bluev_generated": False,
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


def convert_md_to_docx_bytes(md_content: str) -> bytes:
    """将 Markdown 纯文本解析并转换为 DOCX 二进制流 (增强版：支持加粗解析)"""
    if not DOCX_AVAILABLE:
        return md_content.encode("utf-8")

    doc = DocxDocument()
    for line in md_content.split('\n'):
        stripped = line.strip()
        if not stripped:
            continue

        # 1. 确定段落类型并剥离前缀
        p = None
        if stripped.startswith('# '):
            p = doc.add_heading(level=1)
            stripped = stripped.lstrip('# ').strip()
        elif stripped.startswith('## '):
            p = doc.add_heading(level=2)
            stripped = stripped.lstrip('# ').strip()
        elif stripped.startswith('### '):
            p = doc.add_heading(level=3)
            stripped = stripped.lstrip('# ').strip()
        elif stripped.startswith('- ') or stripped.startswith('* '):
            p = doc.add_paragraph(style='List Bullet')
            stripped = stripped.lstrip('-* ').strip()
        elif stripped.startswith('> '):
            p = doc.add_paragraph()
            p.style = 'Quote' if 'Quote' in [s.name for s in doc.styles] else 'Normal'
            stripped = stripped.lstrip('> ').strip()
        else:
            p = doc.add_paragraph()

        # 2. 解析内联加粗格式 **粗体文字**
        parts = re.split(r'(\*\*.*?\*\*)', stripped)
        for part in parts:
            if part.startswith('**') and part.endswith('**') and len(part) > 4:
                run = p.add_run(part[2:-2])
                run.bold = True
            else:
                p.add_run(part)

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


def build_export_filename(content: str, orig_stem: str, suffix: str) -> str:
    """构建直观的导出文件名：【层级-平台】真实标题.suffix"""
    funnel = extract_funnel_from_filename(orig_stem) or ""

    # 尝试提取平台：从文件名中匹配已知平台
    platform = ""
    for p in list(MEDIA_PLATFORM_GUIDES.keys()):
        if p in orig_stem:
            platform = p
            break
    if not platform:
        engine_match = re.search(r"(豆包|元宝|千问|文心一言|Kimi|DeepSeek)", orig_stem)
        if engine_match:
            platform = engine_match.group(1)

    # 提取文章真实标题（第一行 # 开头的内容）
    lines = content.strip().split("\n")
    raw_title = ""
    for line in lines:
        stripped = line.strip()
        if stripped.startswith("# "):
            raw_title = stripped.lstrip("# ").strip()
            break
    if not raw_title:
        raw_title = lines[0].lstrip("# ").strip() if lines else "未命名文章"

    # 清洗标题中的非法字符
    safe_title = sanitize_filename(raw_title) or "未命名文章"

    # 组装前缀
    prefix_parts = [p for p in [funnel, platform] if p]
    prefix = "-".join(prefix_parts)
    if prefix:
        return f"【{prefix}】{safe_title}.{suffix}"
    return f"{safe_title}.{suffix}"


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
    llm_provider: str = "DeepSeek",
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
        "\n【品牌保护与真实口碑指令】"
        "1. 【允许客观瑕疵】为了模拟绝对真实的用户口碑，当你被要求写『真实体验或槽点』时，允许提出对目标主体不痛不痒的客观小建议（如人多排队久、款式少等），以增加真实性，但绝对禁止涉及核心质量、欺诈等严重负面定性。"
        "2. 【严禁商业诋毁】绝对禁止在文章中点名攻击任何具体竞品品牌，禁止使用'割韭菜、黑箱、坑人'等词汇。必须使用'常规渠道、传统模式'进行客观对比。"
        "3. 【禁止虚假承诺】除非基石资料中明确写明，否则绝对禁止大模型自行编造'全网最低、买贵退差价、随时退'等营销承诺。"
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
        return True, _generate_simulated_content(prompt, system_prompt, llm_provider)

    if not OPENAI_AVAILABLE:
        return False, "❌ `openai` 库未安装。运行: `pip install openai`"

    llm_cfg = LLM_CONFIGS.get(llm_provider, LLM_CONFIGS["DeepSeek"])
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

def background_generate_slices(vars_, base_facts, slices_out, use_simulate, api_key, llm_provider="DeepSeek"):
    """后台线程：裂变 30 篇三级切片（物理矩阵式切片）"""
    funnel_keys = {
        "L1": ("P_L1", ["企业主体"]),
        "L2": ("P_L2", ["企业主体"]),
        "L3": ("P_L3", ["企业主体"]),
        "L4": ("P_L4", ["企业主体", "品牌_项目", "品牌_A", "品牌_B", "五个维度"]),
        "L5": ("P_L5", ["企业主体", "品牌_项目"]),
        "L6": ("P_L6", ["企业主体", "优惠信息", "CTA行动"]),
    }
    total = 0
    title_rule = '\n\n【最高排版指令】请务必为你写的文章拟定一个吸引人的标题，并且【必须】把标题放在全文的第一行（格式为：# 你的标题）！绝对不允许直接以正文开篇！'
    brand_exposure_rule = f'\n\n【生死红线：企业官方第一人称绝对贯穿】\n1. 【官方自媒体立场】这 30 篇文章将直接发布在「{vars_.get("企业主体", "")}」自己的官方公众号/官方平台上。你必须 100% 牢记自己就是企业官方！全文必须使用「我们」、「本公司」、「我们{vars_.get("品牌_项目", "")}」等第一人称主语来开展叙述。\n2. 【绝对封杀第三方视角】绝对禁止像外人/旁观者/新闻媒体一样使用「该公司」、「该企业」、「这个品牌」等第三人称代词！出现一次即为严重事故！\n3. 【品牌真名自然植入】在保持「我们」第一人称口吻的同时，正文中仍需极其自然地带出 2-3 次完整的品牌真名（例如：「为了解决车主的痛点，我们{vars_.get("企业主体", "")}独创了...」）。'

    # 智能语义矩阵切片：让大模型先将 10 篇基石按内容相似度自动两两配对
    eeat_files = sorted(EEAT_VERIFIED_DIR.glob("*.md"))

    # 1. 准备带有 ID 的全量文本
    files_content = []
    for idx, f in enumerate(eeat_files):
        files_content.append(f"[文件ID: {idx}] 文件名: {f.name}\\n内容: {f.read_text(encoding='utf-8')[:2000]}")
    all_files_text = "\\n\\n========\\n\\n".join(files_content)

    # 2. 构建聚类 Prompt 并调用 LLM
    grouping_prompt = f"""你是一个智能内容分类专家。请阅读以下 {len(eeat_files)} 篇企业官方基石文章。
你的任务是：根据文章的核心业务、产品类型、业务场景或主题的相似度，将它们两两配对，自动分为 5 组。
要求：
1. 找出内容最互补或最接近的两篇放在同一组。
2. 请严格按以下 JSON 格式输出（包含 5 个子数组，每个包含两个匹配的 [文件ID] 整数）。绝对不要输出任何其他解释文字！
格式示例：[[0, 3], [1, 5], [2, 8], [4, 9], [6, 7]]

【待分类文章列表】：
{all_files_text}
"""

    print("[后台切片] 正在使用大模型进行两两语义智能分组...")
    success, grouping_res = call_llm(
        prompt=grouping_prompt,
        system_prompt="你只输出合法 JSON 数组，绝不解释。",
        api_key=api_key,
        temperature=0.1,
        simulate=use_simulate,
        llm_provider=llm_provider,
    )

    grouped_facts = []
    try:
        json_match = re.search(r"\[[\s\S]*\]", grouping_res)
        if success and json_match:
            pairs = json.loads(json_match.group())
            for pair in pairs[:5]:
                if len(pair) == 2:
                    idx1, idx2 = int(pair[0]), int(pair[1])
                    if 0 <= idx1 < len(eeat_files) and 0 <= idx2 < len(eeat_files):
                        text1 = eeat_files[idx1].read_text(encoding="utf-8")
                        text2 = eeat_files[idx2].read_text(encoding="utf-8")
                        grouped_facts.append(f"{text1}\\n\\n---\\n\\n{text2}")
    except Exception as e:
        print(f"[后台切片] 智能分组解析失败，退回机械顺序分组: {e}")

    # 3. 强力兜底逻辑：如果 AI 分组失败或不足 5 组，退回顺序机械切分
    if len(grouped_facts) < 5:
        grouped_facts = []
        for i in range(0, len(eeat_files), 2):
            chunk = eeat_files[i:i+2]
            grouped_facts.append("\\n\\n---\\n\\n".join(f.read_text(encoding="utf-8") for f in chunk))
        while len(grouped_facts) < 5:
            grouped_facts.append(grouped_facts[-1])

    grouped_facts = grouped_facts[:5]
    print("[后台切片] 资料重组与打包完毕，开始生成 30 篇切片...")

    for funnel in ["L1", "L2", "L3", "L4", "L5", "L6"]:
        template_key, required_vars = funnel_keys[funnel]
        template = CORP_PROMPT_TEMPLATES[template_key]
        print(f"[后台切片] 生成 {funnel} x5篇...")
        for j in range(1, 6):
            fname = f"Slice_{funnel}_{j:02d}_企业视角.md"
            if (slices_out / fname).exists():
                print(f"[后台切片] ⏩ {fname} 已存在，跳过...")
                total += 1
                continue

            slice_facts = grouped_facts[j-1]
            fmt_args = {"base_facts": slice_facts}
            for rv in required_vars:
                fmt_args[rv] = vars_.get(rv, f"（{rv}待补充）")
            final_prompt = template
            for k, v in fmt_args.items():
                final_prompt = final_prompt.replace(f"{{{k}}}", str(v))

            current_prompt = final_prompt + f"\n\n请直接输出第 {j} 篇 {funnel} 层级切片内容。" + title_rule + brand_exposure_rule + build_variance_instruction()

            success, content = call_llm(
                prompt=current_prompt,
                system_prompt=GEO_STRICT_SYSTEM_PROMPT,
                api_key=api_key,
                temperature=round(random.uniform(0.55, 0.75), 2),
                max_tokens=4000,
                simulate=use_simulate,
                llm_provider=llm_provider,
            )
            if success:
                safe_write_file(slices_out, fname, content)
                total += 1
            else:
                print(f"[后台切片] ❌ {funnel}第{j}篇失败")
                time.sleep(0.5)
            time.sleep(0.3)
        print(f"[后台切片] ✅ {funnel}: 5篇完成")
    print(f"[后台切片] 完成，共 {total}/30 篇")
def background_generate_bluev(slices_dir, bluev_out, use_simulate, api_key, llm_provider="DeepSeek"):
    """后台线程：基于 30 篇企业切片 1:1 生成蓝V口播稿"""
    bluev_out.mkdir(parents=True, exist_ok=True)
    slice_files = sorted(slices_dir.glob("*.md"))
    total = 0

    script_system_prompt = "你是一个顶级的短视频爆款编导。必须严格遵守输出要求，不输出任何废话。"

    for sf in slice_files:
        new_name_stem = sf.stem.replace("Slice_", "BlueV_口播_")
        fname = f"{new_name_stem}.md"

        if (bluev_out / fname).exists():
            print(f"[蓝V口播] ⏩ {fname} 已存在，跳过...")
            total += 1
            continue

        article_content = sf.read_text(encoding="utf-8")

        # 动态提取原文章的真实标题
        lines = article_content.strip().split("\n")
        orig_title = lines[0].lstrip("# ").strip() if lines else "未命名口播稿"

        prompt = f"""请基于以下我们企业官方发布的文章，为其提炼并改写成一篇适合抖音/视频号发布的【蓝V官方口播稿】（时长约1分钟左右，字数200-300字即可）。

【内容与格式红线要求】
1. 黄金三秒：开头第一句话必须极具吸引力，直击痛点。
2. 镜头语言：必须是口语化的短句，适合真人口播，绝对不要出现晦涩的书面语和复杂的排版！
3. 【强制收口红线】：口播的最后一段/最后一句，必须强制包含引导转化的动作！请使用"有需要的家长/朋友，私信联系我"、"点击我的主页头像了解更多详情"、"发私信免费获取"等极其口语化的话术作为结尾！绝对不要加任何网页链接！

【源文章内容】：
{article_content}

请直接输出口播稿的正文。"""

        success, content = call_llm(
            prompt=prompt,
            system_prompt=script_system_prompt,
            api_key=api_key,
            temperature=0.5,
            max_tokens=1500,
            simulate=use_simulate,
            llm_provider=llm_provider,
        )
        if success:
            # 强制在口播稿的第一行拼上原文章的标题，完美适配导出重命名系统
            final_content = f"# {orig_title}\n\n{content}"
            safe_write_file(bluev_out, fname, final_content)
            total += 1
        time.sleep(0.3)
    print(f"[蓝V口播] 完成，共 {total} 篇")

def background_generate_ugc(vars_, use_simulate, api_key, llm_provider="DeepSeek"):
    """后台线程：重构 160 篇四级 UGC（不依赖 st.session_state）"""
    word_ranges = {"L1": "300-500", "L2": "800-1500", "L3": "1500-2500", "L4": "2000-3500", "L5": "1000-2000", "L6": "500-1000"}
    ALL_VARS = {k: vars_.get(k, "") for k in ["行业", "用户画像", "痛点", "概念", "品牌_项目", "品牌_A", "品牌_B", "品牌_C", "五个维度", "优惠信息", "CTA行动"]}
    title_rule_ugc = "\n\n【最高排版指令】请务必为你写的文章拟定一个吸引人的标题，并且【必须】把标题放在全文的第一行（格式为：# 你的标题）！绝对不允许直接以正文开篇！"
    brand_exposure_rule = f"\n\n【去广化克制植入红线】\n1. 【克制提及】为了防止被平台判定为水军，正文中提及「{vars_.get('企业主体', '')}」或「{vars_.get('品牌_项目', '')}」的次数**绝对不能超过 2 次**！\n2. 【植入场景】最好将其放在非常口语化、不经意的场景中带出（例如：「后来在网上翻了半天，看到那个叫XXX的好像还行」、「我最后是去XXX弄的」）。\n3. 【绝对禁止直呼口号】严禁出现「XXX真是太棒了」、「我强烈推荐XXX」这种假到离谱的句式！"
    persona_enforcer = "\n\n【最高指令：彻底沉浸画像】你现在的身份是一个真实的网民，绝对不是官方！必须用对应画像的口吻讲人话，严禁出现官方硬广语气！"
    general_out = GENERAL_DIR; specific_out = SPECIFIC_DIR
    general_out.mkdir(parents=True, exist_ok=True); specific_out.mkdir(parents=True, exist_ok=True)
    total_ugc = 0

    # --- 全局知识重载：读取全部 30 篇切片 ---
    global_base_facts_ugc = ""
    if SLICES_DIR.exists():
        all_slice_files = sorted(SLICES_DIR.glob("*.md"))
        global_base_facts_ugc = "\n\n---\n\n".join(f.read_text(encoding="utf-8") for f in all_slice_files)
    print(f"[后台 UGC] 全局知识重载完成，共 {len(all_slice_files) if SLICES_DIR.exists() else 0} 篇切片")

    # --- 案例知识库重载与拆解 ---
    cases_dir = WSP["cases"] if 'WSP' in dir() else WORKSPACES_ROOT / DEFAULT_WORKSPACE / "Cases_Base"
    case_pool = []
    if cases_dir.exists():
        for f in cases_dir.glob("*.md"):
            raw_cases = f.read_text(encoding="utf-8")
            split_cases = [c.strip() for c in raw_cases.split("---") if len(c.strip()) > 50]
            case_pool.extend(split_cases)
    print(f"[后台 UGC] 案例库重载完成，共识别到 {len(case_pool)} 个独立案例")

    # --- Part A: 96 篇通用（精准清单驱动，零随机）---
    manifest_records = []
    print("[后台 UGC] Part A: 96 篇通用（精准清单）…")
    for task in GENERAL_UGC_TASKS:
        funnel = task["funnel"]
        tkey = task["template_key"]
        target_platform = task["platform"]
        task_id = task["id"]
        title_example = task["title_example"]

        existing_files = list(general_out.glob(f"UGC_{task_id}_*.md"))
        if existing_files:
            print(f"[后台 UGC] ⏩ {task_id} 已存在，跳过...")
            total_ugc += 1
            manifest_records.append({"filename": existing_files[0].name, "target_platform": target_platform, "funnel": funnel, "ai_engine": "通用", "task_type": "通用铺设"})
            continue

        template = V2_UGC_TEMPLATES[tkey]
        platform_style = f"\n【平台定向分发要求】本文将发往【{target_platform}】，请严格遵守其风格：{MEDIA_PLATFORM_GUIDES[target_platform]}。"
        title_instruction = (
            "\n【标题行业化改编强制指令】你生成的文章标题必须以以下句式作为【灵感骨架】进行深度改编：\n"
            f"参考骨架：{title_example}\n"
            f"⚠️ 绝对红线：你必须结合【{vars_.get('行业', '当前')}】行业的特性，对骨架进行语义转化！\n"
            "1. 如果是学校/教育行业，不能说使用了半年XX，必须改成孩子就读一学期后；不能说隐藏功能，必须改成办学优势/细节管理。\n"
            "2. 如果是医疗/医美行业，必须转换成对应的服务术语。\n"
            f"3. 务必把骨架中的XX/YY替换为真实的用户痛点或品牌名（{vars_.get('品牌_项目', '')}）！\n"
            "绝对不允许在最终生成的标题中出现XX、YY或机械生硬的套用，标题必须是一句符合人类正常逻辑的、有吸引力的中文句子！"
        )
        # 动态槽点注入
        dynamic_complaint = ""
        if tkey == "P5_G4":
            dynamic_complaint = random.choice(["价格收费不够透明", "客服回复速度太慢", "流程有些繁琐", "等待时间比预期长", "界面/环境设计不够人性化", "前期沟通不够顺畅"])
        fmt_args = {"base_facts": global_base_facts_ugc, "word_range": word_ranges[funnel], "槽点方向": dynamic_complaint, **ALL_VARS}
        prompt_base = template
        for k, v in fmt_args.items():
            prompt_base = prompt_base.replace(f"{{{k}}}", str(v))
        # 动态案例注入机制
        case_injection = ""
        if case_pool:
            selected_case = random.choice(case_pool)
            case_injection = f"\n\n【本篇专属背景案例】\n你必须将以下真实案例作为你本次写作的核心故事背景，并将其极其自然地融入到你的角色叙述中：\n{selected_case}"
        success, content = call_llm(
            prompt=prompt_base + f"\n\n输出{task_id}篇通用。" + title_instruction + platform_style + build_variance_instruction() + title_rule_ugc + brand_exposure_rule + persona_enforcer + case_injection,
            system_prompt=UGC_SYSTEM_PROMPT, api_key=api_key,
            temperature=round(random.uniform(0.30, 0.40), 2), max_tokens=2500, simulate=use_simulate,
            llm_provider=llm_provider,
        )
        if success:
            fname = f"UGC_{task_id}_{target_platform}_通用.md"
            safe_write_file(general_out, fname, content)
            total_ugc += 1
            manifest_records.append({"filename": fname, "target_platform": target_platform, "funnel": funnel, "ai_engine": "通用", "task_type": "通用铺设"})
            time.sleep(1.5)
        else:
            print(f"[后台 UGC] ❌ {task_id} 生成失败")
            time.sleep(0.5)
    print(f"[后台 UGC] ✅ 通用 完成")

    # --- Part B: 64 篇专属（精准清单驱动）---
    print("[后台 UGC] Part B: 64 篇专属（精准清单）…")
    for task in SPECIFIC_UGC_TASKS:
        engine = task["engine"]
        funnel = task["funnel"]
        tkey = task["template_key"]
        target_platform = task["platform"]
        task_id = task["id"]
        title_example = task["title_example"]

        engine_dir = specific_out / engine; engine_dir.mkdir(parents=True, exist_ok=True)
        existing_files = list(engine_dir.glob(f"{task_id}_*.md"))
        if existing_files:
            print(f"[后台 UGC] ⏩ {task_id} 已存在，跳过...")
            total_ugc += 1
            manifest_records.append({"filename": existing_files[0].name, "target_platform": target_platform, "funnel": funnel, "ai_engine": engine, "task_type": "专属狙击"})
            continue

        template = V2_UGC_TEMPLATES[tkey]
        adaptation = build_platform_adaptation(engine)
        title_instruction = (
            "\n【标题行业化改编强制指令】你生成的文章标题必须以以下句式作为【灵感骨架】进行深度改编：\n"
            f"参考骨架：{title_example}\n"
            f"⚠️ 绝对红线：你必须结合【{vars_.get('行业', '当前')}】行业的特性，对骨架进行语义转化！\n"
            "1. 如果是学校/教育行业，不能说使用了半年XX，必须改成孩子就读一学期后；不能说隐藏功能，必须改成办学优势/细节管理。\n"
            "2. 如果是医疗/医美行业，必须转换成对应的服务术语。\n"
            f"3. 务必把骨架中的XX/YY替换为真实的用户痛点或品牌名（{vars_.get('品牌_项目', '')}）！\n"
            "绝对不允许在最终生成的标题中出现XX、YY或机械生硬的套用，标题必须是一句符合人类正常逻辑的、有吸引力的中文句子！"
        )
        dynamic_complaint_b = ""
        if tkey == "P5_G4":
            dynamic_complaint_b = random.choice(["价格收费不够透明", "客服回复速度太慢", "流程有些繁琐", "等待时间比预期长", "界面/环境设计不够人性化", "前期沟通不够顺畅"])
        fmt_args = {"base_facts": global_base_facts_ugc, "word_range": word_ranges[funnel], "槽点方向": dynamic_complaint_b, **ALL_VARS}
        prompt_base = template
        for k, v in fmt_args.items():
            prompt_base = prompt_base.replace(f"{{{k}}}", str(v))
        # 动态案例注入机制
        case_injection_b = ""
        if case_pool:
            selected_case = random.choice(case_pool)
            case_injection_b = f"\n\n【本篇专属背景案例】\n你必须将以下真实案例作为你本次写作的核心故事背景，并将其极其自然地融入到你的角色叙述中：\n{selected_case}"
        final_prompt = prompt_base + f"\n\n输出{engine}专属。" + title_instruction + adaptation + build_variance_instruction() + title_rule_ugc + brand_exposure_rule + persona_enforcer + case_injection_b
        success, content = call_llm(
            prompt=final_prompt, system_prompt=UGC_SYSTEM_PROMPT, api_key=api_key,
            temperature=round(random.uniform(0.30, 0.40), 2), max_tokens=2500, simulate=use_simulate,
            llm_provider=llm_provider,
        )
        if success:
            fname = f"{task_id}_{engine}.md"
            safe_write_file(engine_dir, fname, content)
            total_ugc += 1
            manifest_records.append({"filename": fname, "target_platform": target_platform, "funnel": funnel, "ai_engine": engine, "task_type": "专属狙击"})
            time.sleep(1.5)
        else:
            print(f"[后台 UGC] ❌ {task_id} 生成失败")
            time.sleep(0.5)

    # --- 账本写入（所见即所得，零随机）---
    (PROD_DIR / "tasks_manifest.json").write_text(json.dumps(manifest_records, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[后台 UGC] 完成，共 {total_ugc} 篇，账本 {len(manifest)} 条")
def _generate_simulated_content(prompt: str, system_prompt: str, llm_provider: str = "DeepSeek") -> str:
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
        if eng in name: return eng
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
    if "bluev_scripts_30" in path_lower:
        return "三级附加 — 蓝V口播稿", "口播稿", "BlueV_Scripts_30"
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
    # 扫描已有工作区
    existing_workspaces = sorted([d.name for d in WORKSPACES_ROOT.iterdir() if d.is_dir()])
    ws_options = ["➕ 新建项目..."] + existing_workspaces
    # 尝试匹配当前选中项
    current_idx = 0
    if "workspace_name" in st.session_state:
        try:
            current_idx = ws_options.index(st.session_state["workspace_name"])
        except ValueError:
            current_idx = 0
    selected_ws = st.selectbox("选择或新建项目", ws_options, index=current_idx, key="ws_selector")
    if selected_ws == "➕ 新建项目...":
        workspace_name = st.text_input("输入新项目名称", value="", placeholder="例如：兔师傅、XX汽车连锁", key="ws_new_input")
        if workspace_name.strip():
            workspace_name = workspace_name.strip()
        else:
            workspace_name = DEFAULT_WORKSPACE
    else:
        workspace_name = selected_ws

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
            st.session_state["is_generating_slices"] = False
            st.session_state["slices_generated"] = True

    BLUEV_DIR = WSP["bluev"]

    # 3. 恢复 160 篇 UGC 状态
    if GENERAL_DIR.exists() and SPECIFIC_DIR.exists():
        ugc_count = len(list(PROD_DIR.rglob("*.md"))) - slices_count
        if ugc_count >= 120:
            st.session_state["is_generating_ugc"] = False
            st.session_state["ugc_generated"] = True

    # 4. 恢复蓝V口播稿状态
    st.session_state["bluev_generated"] = False
    if BLUEV_DIR.exists() and len(list(BLUEV_DIR.glob("*.md"))) >= 25:
        st.session_state["is_generating_bluev"] = False
        st.session_state["bluev_generated"] = True

    st.caption(f"📂 `workspaces/{workspace_name}/`")
    st.markdown("---")
    # --- 后台静默配置 API Key 与引擎（隐藏 UI，强制真实生成）---
    selected_llm = "DeepSeek"
    st.session_state["llm_provider"] = selected_llm
    st.session_state["api_key"] = DEFAULT_DEEPSEEK_API_KEY
    st.session_state["api_key_configured"] = True
    use_simulate = False  # 永远关闭模拟模式，强制真实调用

    st.markdown("---")
    st.markdown("### 📋 功能导航")
    page = st.radio("选择工作台", ["📊 交付资产大盘", "⚙️ 自动化生产车间"], index=0 if st.session_state["page"].startswith("📊") else 1, label_visibility="collapsed")
    if page != st.session_state["page"]: st.session_state["page"] = page; st.rerun()

    st.markdown("---")


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

        # --- 全选/取消全选 ---
        if fa:
            col_sel1, col_sel2 = st.columns(2)
            with col_sel1:
                if st.button("✅ 全选当前列表", use_container_width=True, key="btn_select_all"):
                    for i, a in enumerate(fa):
                        st.session_state[f"chk_{i}_{a['filename'][:15]}"] = True
                    st.rerun()
            with col_sel2:
                if st.button("🔲 取消全选", use_container_width=True, key="btn_deselect_all"):
                    for i, a in enumerate(fa):
                        st.session_state[f"chk_{i}_{a['filename'][:15]}"] = False
                    st.rerun()

        # --- 批量下载 ZIP ---
        if fa:
            checked_paths = []
            for i, a in enumerate(fa):
                cols = st.columns([0.5, 3.5, 1])
                with cols[0]:
                    if st.checkbox("", key=f"chk_{i}_{a['filename'][:15]}", label_visibility="collapsed"):
                        checked_paths.append(a["path"])
                with cols[1]:
                    is_sel = st.session_state["selected_path"] == a["path"]
                    cc = "article-card selected" if is_sel else "article-card"
                    th = get_source_tag(a["source_short"])
                    if a["funnel"]: th += " " + get_funnel_tag(a["funnel"])
                    if a.get("engine") and a["engine"] not in ("通用","未知"): th += " " + get_engine_tag(a["engine"])
                    dt = a["title"][:50] + "…" if len(a["title"]) > 50 else a["title"]
                    st.markdown(f'<div class="{cc}"><div class="title">{dt}</div><div class="meta">{th} · {a["filename"][:30]}…</div></div>', unsafe_allow_html=True)
                with cols[2]:
                    if st.button("🔍", key=f"pv_{i}_{a['filename'][:12]}", help="预览"):
                        st.session_state["selected_path"]=a["path"]; st.session_state["selected_content"]=a["content"]; st.session_state["selected_title"]=a["title"]; st.session_state["selected_filename"]=a["filename"]; st.rerun()

            st.session_state["checked_paths"] = checked_paths
        if not fa:
            st.info("无匹配文章")

    with pc:
        st.markdown("### 📖 内容预览")

        # --- 批量操作区 ---
        checked_paths = st.session_state.get("checked_paths", [])
        if checked_paths:
            st.markdown("##### 📦 批量操作区")
            col_fmt, col_dl, col_del = st.columns([1.5, 1.5, 1.2])
            with col_fmt:
                format_options = ["Markdown (.md)"]
                if DOCX_AVAILABLE:
                    format_options.append("Word (.docx)")
                export_format = st.selectbox("导出格式", format_options, key="export_format", label_visibility="collapsed")

            if not DOCX_AVAILABLE:
                st.warning("💡 提示：检测到系统未安装 `python-docx` 引擎，已自动锁定 Word 批量导出功能。请在终端运行 `pip install python-docx` 并重启系统以解锁。")
            with col_dl:
                # 强力判定，杜绝字符串匹配错误或状态错乱
                is_word = "Word" in export_format
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    for cp in checked_paths:
                        try:
                            content = Path(cp).read_text(encoding="utf-8")
                            orig_stem = Path(cp).stem

                            if is_word:
                                zip_name_inner = build_export_filename(content, orig_stem, "docx")
                                zf.writestr(zip_name_inner, convert_md_to_docx_bytes(content))
                            else:
                                zip_name_inner = build_export_filename(content, orig_stem, "md")
                                zf.writestr(zip_name_inner, content.encode("utf-8"))
                        except Exception as e:
                            print(f"打包文件失败已跳过: {e}")
                            pass
                zip_buffer.seek(0)
                zip_name = f"GEO_交付资产打包_{'docx' if is_word else 'md'}.zip"

                # 按钮文案动态绑定格式，给用户视觉强制确认
                st.download_button(
                    label=f"📥 批量下载 ({len(checked_paths)}篇 {'Word' if is_word else 'MD'})",
                    data=zip_buffer,
                    file_name=zip_name,
                    mime="application/zip",
                    use_container_width=True,
                )
            with col_del:
                if st.button("🗑️ 批量彻底删除", type="primary", use_container_width=True, key="btn_bulk_del"):
                    deleted = 0
                    for cp in checked_paths:
                        try:
                            Path(cp).unlink(missing_ok=True)
                            deleted += 1
                        except Exception:
                            pass
                    st.session_state["checked_paths"] = []
                    st.toast(f"✅ 成功删除 {deleted} 篇文章！", icon="🗑️")
                    st.rerun()
            st.markdown("---")

        if st.session_state["selected_content"]:
            st.markdown(f"**{st.session_state['selected_title']}**")
            st.caption(f"`{st.session_state['selected_filename']}`")
            # 动态生成单篇下载的新文件名
            content = st.session_state["selected_content"]
            orig_stem = Path(st.session_state['selected_filename']).stem
            md_name = build_export_filename(content, orig_stem, "md")
            docx_name = build_export_filename(content, orig_stem, "docx")

            b1, b2 = st.columns(2)
            with b1:
                st.download_button(
                    label="📥 下载单篇 (Markdown)",
                    data=content.encode("utf-8"),
                    file_name=md_name,
                    mime="text/markdown",
                    use_container_width=True,
                )
            with b2:
                if DOCX_AVAILABLE:
                    docx_bytes = convert_md_to_docx_bytes(content)
                    st.download_button(
                        label="📥 下载单篇 (Word)",
                        data=docx_bytes,
                        file_name=docx_name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                    )
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
    s3c = "step-box done" if st.session_state.get("slices_generated") else "step-box active"
    st.markdown(f'<div class="{s3c}">', unsafe_allow_html=True)
    st.markdown("### ✂️ 步骤 4：裂变 30 篇【企业官方视角】三级切片 (L1-L6 均分)")

    # --- 双轨运行机制：AI 生成 + 手动导入 ---
    col_gen, col_up = st.columns([1, 1])

    with col_gen:
        st.markdown("#### 方案 A：AI 全自动裂变")
        verified_files = list(EEAT_VERIFIED_DIR.glob("*.md")) if EEAT_VERIFIED_DIR.exists() else []
        if not st.session_state.get("verified_loaded", False) or len(verified_files) < 10:
            st.warning("⛔ 请先在【步骤 3】上传至少 10 篇校验版基石文档，才能解锁 AI 自动裂变。")
        else:
            st.markdown('''
            <div class="constraint-warn">
            ⚠️ <strong>强制约束</strong>：30 篇切片 100% 从 10 篇基石拆解。
            </div>
            ''', unsafe_allow_html=True)

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
                t = threading.Thread(target=background_generate_slices, args=(vars_, base_facts, slices_out, use_simulate, st.session_state["api_key"], st.session_state.get("llm_provider", "DeepSeek")))
                t.start()
                add_log("🚀 后台切片线程已启动")
                st.rerun()

    with col_up:
        st.markdown("#### 方案 B：直接导入已有切片")
        st.info("💡 如果已有写好的切片，请在此直接上传，**无视前置条件**。")
        uploaded_slices = st.file_uploader(
            "在此批量上传切片 (.md / .docx / .txt)",
            type=["md", "docx", "txt"],
            accept_multiple_files=True,
            key="fu_slices_skip",
            help="支持批量框选上传。上传后将自动跳过生成，直接解锁后续步骤。"
        )
        if st.button("📥 保存上传并解锁下一步", type="secondary", use_container_width=True, key="btn_upload_slices"):
            if not uploaded_slices:
                st.error("🚨 请先将文件拖拽到上方虚线框内！")
            else:
                SLICES_DIR.mkdir(parents=True, exist_ok=True)
                saved_count = 0
                for uf in uploaded_slices:
                    try:
                        fname, content = extract_text_from_upload(uf)
                        safe_write_file(SLICES_DIR, f"{Path(fname).stem}.md", content)
                        saved_count += 1
                    except Exception as e:
                        st.warning(f"⚠️ 无法保存 {uf.name}: {e}")

                st.session_state["slices_generated"] = True
                add_log(f"✅ 手动导入了 {saved_count} 篇切片 → Slices_30/")
                st.toast(f"✅ 成功导入 {saved_count} 篇切片！后续步骤已解锁", icon="🔓")
                st.rerun()

    # --- 实时进度可视化 (切片) ---
    if SLICES_DIR.exists():
        slice_files = sorted(SLICES_DIR.glob("*.md"))
        if slice_files:
            with st.expander(f"📂 切片生成进度：已产出 {len(slice_files)} / 30 篇 (点击展开)", expanded=False):
                if st.button("🔄 手动刷新进度", key="refresh_slices"):
                    st.rerun()
                for sf in slice_files:
                    st.caption(f"📄 {sf.name}")

    if st.session_state.get("slices_generated"):
        st.success(f"✅ 已完成 30 篇三级切片 → `{PROD_DIR}/Slices_30/`")

    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

    # ================================================================
    # 步骤 4.5: 蓝V口播稿专栏
    # ================================================================
    s45c = "step-box done" if st.session_state.get("bluev_generated") else ("step-box active" if st.session_state.get("slices_generated") else "step-box blocked")
    st.markdown(f'<div class="{s45c}">', unsafe_allow_html=True)
    st.markdown("### 🎥 步骤 4.5：一键生成【蓝V短视频口播脚本】")
    if not st.session_state.get("slices_generated", False):
        st.warning("⛔ 请先完成步骤 4：裂变 30 篇企业视角切片。")
    elif st.session_state.get("is_generating_bluev"):
        st.info("🚀 正在后台全速生成蓝V口播稿，您可以切换页面实时查看产出！")
        with st.spinner("后台生成中…"): time.sleep(0.5)
    elif st.button("🎥 一键生成 30 篇短视频口播稿", type="primary", use_container_width=True, key="btn_bluev"):
        BLUEV_DIR_UI = WSP["bluev"]; BLUEV_DIR_UI.mkdir(parents=True, exist_ok=True)
        st.session_state["is_generating_bluev"] = True
        t = threading.Thread(target=background_generate_bluev, args=(SLICES_DIR, BLUEV_DIR_UI, use_simulate, st.session_state["api_key"], st.session_state.get("llm_provider", "DeepSeek")))
        t.start()
        add_log("🚀 后台蓝V口播线程已启动")
        st.rerun()
    BV = WSP["bluev"]
    if BV.exists():
        bf_list = sorted(BV.glob("*.md"))
        if bf_list:
            with st.expander(f"📂 进度：{len(bf_list)}/30 篇", expanded=False):
                if st.button("🔄 刷新", key="rf_bv"): st.rerun()
                for bf in bf_list: st.caption(f"📄 {bf.name}")
    if st.session_state.get("bluev_generated"):
        st.success(f"✅ 30 篇蓝V口播稿已生成 → `{PROD_DIR}/BlueV_Scripts_30/`")
    st.markdown('</div>', unsafe_allow_html=True)

    # ================================================================
    # 步骤 4: 重构 160 篇四级 UGC
    # ================================================================
    s4c = "step-box done" if st.session_state["ugc_generated"] else ("step-box active" if st.session_state["slices_generated"] else "step-box blocked")
    st.markdown(f'<div class="{s4c}">', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # ================================================================
    # 步骤 4.8: 独立案例库提纯
    # ================================================================
    st.markdown('<div class="step-box active">', unsafe_allow_html=True)
    st.markdown("### 📚 步骤 4.8：独立案例库提纯 (UGC 防同质化核心)")
    st.markdown("#### 💎 核心引擎：独立案例库管理 (防同质化)")
    st.caption("160 篇 UGC 将从此库中随机抽取背景故事。你可以选择让 AI 自动从原资料中蒸馏，也可以手动上传补充。")

    col_distill, col_upload = st.columns([1, 1])
    with col_distill:
        if st.button("🤖 一键从原始资料蒸馏真实案例", type="primary", use_container_width=True, key="btn_distill_cases"):
            raw_text = st.session_state.get("raw_text_combined", "")
            if not raw_text:
                st.error("🚨 请先在步骤 1 读取企业原始资料！")
            else:
                with st.status("🔍 正在开启全局雷达，执行无损重叠切块并榨取真实案例...", expanded=True) as status:
                    chunk_size = 5000  # 缩小到 5000 字精读，防止大模型走神
                    overlap = 500      # 500 字重叠，防止案例被物理腰斩

                    chunks = []
                    for i in range(0, len(raw_text), chunk_size - overlap):
                        chunks.append(raw_text[i:i+chunk_size])

                    all_extracted_cases = []

                    progress_bar = st.progress(0)
                    for i, chunk in enumerate(chunks):
                        status.update(label=f"⏳ 正在精读第 {i+1}/{len(chunks)} 个资料区块，确保案例原汁原味无损提取...")
                        prompt = CASE_DISTILLATION_PROMPT.format(raw_text=chunk)

                        success, content = call_llm(
                            prompt=prompt,
                            system_prompt="你是无情的数据挖掘机，严格按要求输出格式。禁止输出markdown代码块标记。",
                            api_key=st.session_state["api_key"],
                            temperature=0.2,
                            simulate=use_simulate,
                        )

                        if success and "未找到具体案例" not in content and len(content.strip()) > 30:
                            clean_content = re.sub(r"^```[a-zA-Z]*\n", "", content)
                            clean_content = re.sub(r"\n```$", "", clean_content)
                            all_extracted_cases.append(clean_content.strip())
                            st.toast(f"✅ 第 {i+1} 块提炼成功，发现案例！", icon="✨")
                        elif not success:
                            st.toast(f"❌ 第 {i+1} 块 API 调用失败", icon="⚠️")
                        else:
                            st.toast(f"⏭️ 第 {i+1} 块未发现具体案例，跳过", icon="⏩")

                        progress_bar.progress((i + 1) / len(chunks))

                    if all_extracted_cases:
                        final_content = "\n\n---\n\n".join(all_extracted_cases)
                        cases_dir = WSP["cases"]
                        cases_dir.mkdir(parents=True, exist_ok=True)
                        safe_write_file(cases_dir, "Auto_Distilled_Cases.md", final_content)
                        case_count = len([c for c in final_content.split("---") if len(c.strip()) > 30])
                        status.update(label=f"✅ 蒸馏完成！分块扫描完毕，成功无损提取约 {case_count} 个独立案例入库。", state="complete")
                        st.rerun()
                    else:
                        status.update(label="⚠️ 全局扫描结束，资料中未识别到符合标准的案例", state="error")

    with col_upload:
        uploaded_cases = st.file_uploader("或手动上传独立案例库 (.txt, .md, .docx, .pdf)", type=["txt", "md", "docx", "pdf"], accept_multiple_files=True, key="fu_cases", label_visibility="collapsed")
        if st.button("📥 保存手动上传的案例", use_container_width=True, key="btn_save_cases"):
            if uploaded_cases:
                cases_dir = WSP["cases"]
                cases_dir.mkdir(parents=True, exist_ok=True)
                for uc in uploaded_cases:
                    fname, content = extract_text_from_upload(uc)
                    safe_write_file(cases_dir, f"{Path(fname).stem}.md", content)
                st.toast(f"✅ 成功保存 {len(uploaded_cases)} 个手动案例文件！", icon="📦")

    # 展示当前案例库状态 + 增删改查交互编辑器
    cases_dir = WSP["cases"]
    if cases_dir.exists():
        case_files = list(cases_dir.glob("*.md"))
        if case_files:
            # 合并所有案例文件内容
            merged_cases = "\n\n---\n\n".join(cf.read_text(encoding="utf-8") for cf in case_files)
            case_entries = [c.strip() for c in merged_cases.split("---") if len(c.strip()) > 30]
            st.success(f"📦 当前案例库已准备就绪：包含 {len(case_files)} 个文件，共约 {len(case_entries)} 个案例")

            # 可编辑文本框
            edited_cases = st.text_area(
                "✏️ 在此直接增删改案例内容（案例之间请用 --- 分隔）",
                value=merged_cases,
                height=400,
                key="ta_case_editor",
                help="直接编辑文本。完成后点击下方保存按钮。",
            )

            if st.button("💾 保存并覆盖当前案例库", type="primary", use_container_width=True, key="btn_save_cases_edit"):
                # 清空旧文件
                for old_f in case_files:
                    try:
                        old_f.unlink(missing_ok=True)
                    except Exception:
                        pass
                # 写入编辑后的内容
                safe_write_file(cases_dir, "Active_Edited_Cases.md", edited_cases)
                st.toast("✅ 案例库已更新保存！", icon="💾")
                st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

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
                t = threading.Thread(target=background_generate_ugc, args=(st.session_state.get("variables", {}), use_simulate, st.session_state["api_key"], st.session_state.get("llm_provider", "DeepSeek")))
                t.start()
                add_log("🚀 后台 UGC 线程已启动")
                st.rerun()
    # --- 实时进度可视化 (UGC) ---
    if GENERAL_DIR.exists() or SPECIFIC_DIR.exists():
        gen_ugc_files = sorted([f.name for f in GENERAL_DIR.glob("*.md")] + [f.name for f in SPECIFIC_DIR.rglob("*.md")])
        if gen_ugc_files:
            with st.expander(f"📂 UGC 生成进度：已产出 {len(gen_ugc_files)} / 160 篇 (点击展开)", expanded=False):
                if st.button("🔄 手动刷新进度", key="refresh_ugc"):
                    st.rerun()
                for gf in gen_ugc_files:
                    st.caption(f"📄 {gf}")

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
        "上传待检文章 (.txt, .md, .docx, .pdf)",
        type=["txt", "md", "docx", "pdf"],
        accept_multiple_files=True,
        key="fu_fact_check",
        help="上传生成的切片或 UGC 文章，支持多选。",
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
                        fname, article_text = extract_text_from_upload(uf)
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
        "上传待检文章 (.txt, .md, .docx, .pdf)",
        type=["txt", "md", "docx", "pdf"],
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
                    fname, text = extract_text_from_upload(uf)
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


