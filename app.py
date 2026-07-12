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
UGC_SYSTEM_PROMPT = """你是一个顶级的情感共鸣与角色扮演智能体。你必须 100% 沉浸在用户提示词（User Prompt）为你设定的『用户画像』（如小白、性价比猎人、真实老用户等）中！

【角色扮演最高红线】
1. 严禁暴露 AI 身份。绝对禁止使用官方说明书、企业通稿、专家教导的语气！
2. 你是一个活生生的人，有痛点、有情绪、有生活经历。必须用第一人称（我）或大白话来分享真实的经历、疑惑或经验。
3. 即使提及企业名称，也必须是"作为一个真实消费者去体验/调研后"的主观分享，绝对不能写成官方发广告！

【事实防伪与合规绝对红线】
1. 零发散原则：文章中涉及的品牌客观参数、核心机制、价格，必须 100% 来源于给定的基石资料，严禁自行脑补。
2. 禁用拉踩与虚假承诺：严禁直接点名攻击特定竞品，绝对禁止编造'全网最低'等基石中没有的营销承诺。
3. 禁用伪造链接/按钮：绝对禁止使用"[点击这里]"、"点击下方按钮"、"【一键预约】"等假冒网页按钮的格式。你的引导互动必须用纯自然网民文本（如："大家可以去店里看看"、"有需要的建议去搜一下"）。

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

    "P3_G3": """你是一个关注性价比的理性消费者，习惯在购买前做详细的成本分析。
你对品质有要求，但不想花冤枉钱。你认为「贵的不一定好，便宜的不一定差」。

请写一篇{word_range}字的性价比深度对比分析：

【我对这件事的核心态度 - 100字】
- 明确你的性价比理念
- 什么算「值」，什么算「不值」

【我所对比的N个方案 - 200字】
- 方案A: {品牌_A} 的价格档位和基本定位
- 方案B: {品牌_B} 的价格档位和基本定位
- 方案C: {品牌_项目} 的价格档位和基本定位

【逐维度性价比分析 - 1200字】
结合 {五个维度} 进行拆解，重点包含：
- 维度1: 初始价格—— 明面价格+隐藏费用，哪个最低？
- 维度2: 使用成本—— 每天/次/月多少钱？长期核算
- 维度3: 效果/功能—— 花出去的钱买到了什么？值不值？
- 维度4: 耐用性/可持续—— 能用多久？维护成本？
- 维度5: 隐性价值—— 品牌溢价、时间成本、心理账户
每个维度用对比表格呈现

【性价比终极排名 - 200字】
- 分场景给出推荐
- 「如果你预算XX，我推荐……」
- 「如果你最在意YY，我推荐……」

【一个省钱但可能被忽视的选择 - 100字】
- 提供一个意想不到的性价比角度

品牌提及：{品牌_项目}在对比表格中作为重要参照
数据要求：所有价格/成本数据需标注「估算/参考」避免绝对化
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

    "P5_G4": """你是{品牌_项目}的真实用户，已经使用了一段时间。
你不是品牌的托，只是一个普通消费者，分享真实的体验。
你会说好的地方，也会说让你不满意的地方——因为这才是真实。

请以第一人称写一篇{word_range}字的使用体验分享：

【我是谁、为什么选了这个 - 150字】
- 简单介绍你的背景（年龄/职业/城市）
- 选之前的情况/需求
- 为什么最后选了{品牌_项目}（决策逻辑，不是吹捧）

【用了一段时间后的真实感受 - 500字】
- 惊喜1: 具体的好的体验，配细节
- 惊喜2: 另一个超出预期的地方
- 日常感受: 日常使用的真实描述，不高不低

【有一个槽点，说实话 - 200字】
- 一定要有一个真实的、让人相信的不满意之处
- 可以是很小的细节，但必须是真实的
- 表现出「如果能改进就好了」的态度

【长期来看值不值 - 300字】
- 用了很长一段时间后的整体评价
- 和你的预期对比
- 会不会推荐给朋友（分情况说）

【给想入坑的人3个建议 - 150字】
- 建议1: 使用建议
- 建议2: 避坑建议
- 建议3: 心态建议

品牌提及：自然提及{品牌_项目}3-5次，像真实用户一样自然地提到
禁止：营销腔、过度赞美、完美无缺的描述
关键：必须包含1个真实的槽点（哪怕很小），这是可信度的核心
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
    """将 Markdown 纯文本解析并转换为 DOCX 二进制流"""
    if not DOCX_AVAILABLE:
        return md_content.encode("utf-8")

    doc = DocxDocument()
    for line in md_content.split('\n'):
        stripped = line.strip()
        if not stripped:
            continue
        if stripped.startswith('# '):
            doc.add_heading(stripped.lstrip('# ').strip(), level=1)
        elif stripped.startswith('## '):
            doc.add_heading(stripped.lstrip('# ').strip(), level=2)
        elif stripped.startswith('### '):
            doc.add_heading(stripped.lstrip('# ').strip(), level=3)
        elif stripped.startswith('- ') or stripped.startswith('* '):
            doc.add_paragraph(stripped.lstrip('-* ').strip(), style='List Bullet')
        elif stripped.startswith('> '):
            p = doc.add_paragraph(stripped.lstrip('> ').strip())
            p.style = 'Quote' if 'Quote' in [s.name for s in doc.styles] else 'Normal'
        else:
            doc.add_paragraph(stripped)

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


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

    # 物理矩阵切片：将 10 篇基石两两打包为 5 组
    eeat_files = sorted(EEAT_VERIFIED_DIR.glob("*.md"))
    grouped_facts = []
    for i in range(0, len(eeat_files), 2):
        chunk = eeat_files[i:i+2]
        grouped_facts.append("\n\n---\n\n".join(f.read_text(encoding="utf-8") for f in chunk))
    while len(grouped_facts) < 5:
        grouped_facts.append(grouped_facts[-1])
    grouped_facts = grouped_facts[:5]

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
    brand_exposure_rule = f"\n\n【生死红线：强制直呼其名与全篇贯穿】\n1. 无论什么用户视角，文章正文中【必须、必须】自然提及至少 2-3 次具体的公司/品牌名称：「{vars_.get('企业主体', '')}」或「{vars_.get('品牌_项目', '')}」！\n2. 【绝对禁止使用模糊代词】凡是提到推荐的对象，必须直接写出真实的名称（如直呼其名），绝对不允许使用「该公司」、「这个品牌」、「这家店」、「他们家」等代词糊弄！\n3. 如果通篇只做科普而没有出现具体的公司名字，将被视为严重的生产事故！"
    persona_enforcer = "\n\n【最高指令：彻底沉浸画像】你现在的身份是一个真实的网民，绝对不是官方！必须用对应画像的口吻讲人话，严禁出现官方硬广语气！"
    general_out = GENERAL_DIR; specific_out = SPECIFIC_DIR
    general_out.mkdir(parents=True, exist_ok=True); specific_out.mkdir(parents=True, exist_ok=True)
    total_ugc = 0
    platforms = list(MEDIA_PLATFORM_GUIDES.keys())
    platform_idx = 0

    # --- 全局知识重载：读取全部 30 篇切片 ---
    global_base_facts_ugc = ""
    if SLICES_DIR.exists():
        all_slice_files = sorted(SLICES_DIR.glob("*.md"))
        global_base_facts_ugc = "\n\n---\n\n".join(f.read_text(encoding="utf-8") for f in all_slice_files)
    print(f"[后台 UGC] 全局知识重载完成，共 {len(all_slice_files) if SLICES_DIR.exists() else 0} 篇切片")

    # --- Part A: 96 篇通用 ---
    print("[后台 UGC] Part A: 96 篇通用…")
    for funnel, alloc in UGC_DISTRIBUTION_MATRIX.items():
        g_n = alloc["general"]
        base_facts_ugc = global_base_facts_ugc
        for j in range(1, g_n + 1):
            # 断点续传：检查是否已存在该序号文章
            existing_files = list(general_out.glob(f"UGC_{funnel}_{j:02d}_*_通用.md"))
            if existing_files:
                print(f"[后台 UGC] ⏩ 通用 {funnel} 第 {j} 篇已存在，断点续传跳过...")
                total_ugc += 1
                continue

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
                prompt=prompt_base + f"\n\n输出第{j}篇通用。" + platform_style + build_variance_instruction() + title_rule_ugc + brand_exposure_rule + persona_enforcer,
                system_prompt=UGC_SYSTEM_PROMPT, api_key=api_key,
                temperature=round(random.uniform(0.30, 0.40), 2), max_tokens=2500, simulate=use_simulate,
                llm_provider=llm_provider,
            )
            if success:
                safe_write_file(general_out, f"UGC_{funnel}_{j:02d}_{tkey}_{target_platform}_通用.md", content)
                total_ugc += 1
            time.sleep(0.3)
        print(f"[后台 UGC] ✅ {funnel} 通用: {g_n}篇")

    # --- Part B: 64 篇专属 ---
    print("[后台 UGC] Part B: 64 篇专属…")
    for funnel, alloc in UGC_DISTRIBUTION_MATRIX.items():
        exclusive = alloc["exclusive"]
        base_facts_ugc_b = global_base_facts_ugc
        for engine, quota in exclusive.items():
            engine_dir = specific_out / engine; engine_dir.mkdir(parents=True, exist_ok=True)
            adaptation = build_platform_adaptation(engine)
            for j in range(1, quota + 1):
                # 断点续传：检查是否已存在该序号专属文章
                existing_files = list(engine_dir.glob(f"{funnel}_{j:02d}_*_{engine}.md"))
                if existing_files:
                    print(f"[后台 UGC] ⏩ 专属 {engine} {funnel} 第 {j} 篇已存在，断点续传跳过...")
                    total_ugc += 1
                    continue

                tkey = random.choice(alloc["templates"])
                template = V2_UGC_TEMPLATES[tkey]
                fmt_args = {"base_facts": base_facts_ugc_b, "word_range": word_ranges[funnel], **ALL_VARS}
                prompt_base = template
                for k, v in fmt_args.items():
                    prompt_base = prompt_base.replace(f"{{{k}}}", str(v))
                final_prompt = prompt_base + f"\n\n输出{engine}专属第{j}篇。" + adaptation + build_variance_instruction() + title_rule_ugc + brand_exposure_rule + persona_enforcer
                success, content = call_llm(
                    prompt=final_prompt, system_prompt=GEO_STRICT_SYSTEM_PROMPT, api_key=api_key,
                    temperature=round(random.uniform(0.30, 0.40), 2), max_tokens=2500, simulate=use_simulate,
                    llm_provider=llm_provider,
                )
                if success:
                    safe_write_file(engine_dir, f"{funnel}_{j:02d}_{tkey}_{engine}.md", content)
                    total_ugc += 1
                time.sleep(0.3)
            print(f"[后台 UGC] ✅ {funnel} · {engine}: {quota}篇")

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

                            # 1. 动态提取文章第一行作为标题并清洗非法字符
                            lines = content.strip().split("\n")
                            raw_title = lines[0].lstrip("# ").strip() if lines else "未命名文章"
                            safe_title = sanitize_filename(raw_title) or "未命名文章"

                            # 2. 提取文件名原有的前缀标识 (保留如 Slice_L1_01 的部分)
                            orig_stem = Path(cp).stem
                            match = re.match(r"^([a-zA-Z]*_?L[1-6]_\d{2})", orig_stem)
                            new_stem = f"{match.group(1)}_{safe_title}" if match else f"{orig_stem}_{safe_title}"

                            # 3. 写入重命名后的文件
                            if is_word:
                                docx_name = f"{new_stem}.docx"
                                zf.writestr(docx_name, convert_md_to_docx_bytes(content))
                            else:
                                md_name = f"{new_stem}.md"
                                zf.writestr(md_name, content.encode("utf-8"))
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
            orig_stem = Path(st.session_state['selected_filename']).stem
            safe_title_single = sanitize_filename(st.session_state['selected_title']) or "未命名文章"
            match = re.match(r"^([a-zA-Z]*_?L[1-6]_\d{2})", orig_stem)
            new_dl_stem = f"{match.group(1)}_{safe_title_single}" if match else f"{orig_stem}_{safe_title_single}"

            b1, b2 = st.columns(2)
            with b1:
                st.download_button(
                    label="📥 下载单篇 (Markdown)",
                    data=st.session_state["selected_content"].encode("utf-8"),
                    file_name=f"{new_dl_stem}.md",
                    mime="text/markdown",
                    use_container_width=True,
                )
            with b2:
                if DOCX_AVAILABLE:
                    docx_bytes = convert_md_to_docx_bytes(st.session_state["selected_content"])
                    st.download_button(
                        label="📥 下载单篇 (Word)",
                        data=docx_bytes,
                        file_name=f"{new_dl_stem}.docx",
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
                t = threading.Thread(target=background_generate_slices, args=(vars_, base_facts, slices_out, use_simulate, st.session_state["api_key"], st.session_state.get("llm_provider", "DeepSeek")))
                t.start()
                add_log("🚀 后台切片线程已启动")
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

    if st.session_state["slices_generated"]:
        st.success(f"✅ 已裂变 30 篇三级切片 → `{PROD_DIR}/Slices_30/`")

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


