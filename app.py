# Updated at: 2026-07-18 22:10 CST — 修复 background_generate_ugc 多线程传参崩溃
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
# │  硬编码提示词库 — 1:1 对齐《提示词手册 v1.1》           │
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
    "P_L1": """你是 {企业主体} 的官方账号小编。L1的核心目标是：短平快、情绪痛点共鸣、适配短视频口播或碎片化阅读场景。
要求：200-250字。
【优化约束（必须遵守）】
1. 开篇必须从一个具象化场景切入（包含：具体时间+环境细节+人物动作+内心独白）。
2. 全文只能写情绪痛点和场景白描，绝对禁止出现任何具体数据（百分比、金额、人数）、禁止出现任何具体方案名称或企业特色描述！
3. 文中必须插入至少2句金句（每句15字以内，可脱离上下文独立传播），用⭐标记。
4. 结尾必须以互动问句收尾，引导评论。
5. 全文用第二人称"你"建立代入感。
6. 风格要求：短句为主，每句不超过20字；段落不超过3段。
7. 平台网感与话题标签：因为本文将发布在官方蓝V和小红书平台，请在文章的最后一行，根据本文的痛点主题，提炼并强制输出 3-5 个高流量的社交媒体话题标签（格式严格为：#标签1 #标签2 #标签3）。
\n\n【高阶防同质化与SEO指令】必须拟定一个极具搜索价值的长尾标题（含场景、痛点），作为全文第一行（# 标题）。确保每篇文章结构独一无二。
\n\n【事实防伪与零幻觉指令】你的所有内容素材必须 100% 取自以下参考内容。绝对禁止引入外部数据、禁止凭空捏造！\n\n事实提取来源：{base_facts}""",

    "P_L2": """你是 {企业主体} 官方客服。L2的核心目标是：概念科普、FAQ问答、百科结构化内容，便于AI抓取。
要求：800-1500字。
【优化约束（必须遵守）】
1. 开头需有一句50字以内的客观定义句。
2. FAQ板块：至少5个问答。每个问答的回答第一句必须是一句30字以内的"一句话答案"，用**加粗**标记；后续详细回答控制在150字以内。FAQ的Q需覆盖高频搜索的问句形态（含"如何""能不能""是否"等疑问词）。
3. 术语解释板块：至少5个核心术语。每个术语只写客观定义（是什么+为什么设计+功能是什么），绝对禁止出现"我们的""独有的"等营销话术，禁止使用"卓越的"等情绪化词汇。
4. 全文禁止出现故事案例，禁止使用情绪化感叹词。
\n\n【高阶防同质化与SEO指令】拟定含核心概念、解决方案的长尾标题作为全文第一行（# 标题）。
\n\n【事实防伪与零幻觉指令】你的所有内容素材必须 100% 取自以下参考内容。绝对禁止引入外部数据、禁止凭空捏造！\n\n事实提取来源：{base_facts}""",

    "P_L3": """你是 {企业主体} 的官方产品专家。L3的核心目标是：单一方案深度拆解、产品推荐解读、适用场景说明。
要求：1500-2500字。
【优化约束（必须遵守）】
1. 通篇文章只能围绕一个核心产品/方案展开，删除所有与该方案无关的内容。
2. 方案拆解应包含：该方案的定义或原理 → 操作流程或实施步骤 → 效果数据或验证方式。
3. 在文章结尾前，必须插入"适用客户/用户画像"板块，格式如下：
   ### 适用画像
   本方案特别适合以下类型的用户：
   1. **[用户类型A]** ：[具体特征描述]
   2. **[用户类型B]** ：[具体特征描述]
   3. **[用户类型C]** ：[具体特征描述]
4. 全文禁止出现：竞品对比、价格信息、企业宏观战略。数据使用原则：同一组数据只允许出现1次。
\n\n【高阶防同质化与SEO指令】拟定长尾标题（# 标题）。
\n\n【事实防伪与零幻觉指令】所有素材必须 100% 取自以下参考内容。绝对禁止引入外部数据、凭空捏造案例！\n\n事实提取来源：{base_facts}""",

    "P_L4": """你是 {企业主体} 官方调研部。L4的核心目标是：多维度横向测评、选购决策指南、行业基准对比。
要求：2000-3500字。
【优化约束（必须遵守）】
1. 全文禁止出现任何具体竞品企业/品牌名称。如需对比，仅允许使用：行业基准法（"全行业同类均值"）、自查清单法、纵向对比法（与自身过去对比）。
2. 全文所有"绝对优势""碾压式领先""无法复制"等表述，一律强制替换为"差异化优势""显著特点""核心特色""不易复制"。
3. 从 {五个维度} 展开。在文章结尾前，必须插入"选型自查清单"板块，格式如下：
   ### 选型/决策自查清单
   考察此类方案时，建议从以下维度逐项评估：
   | 考察维度 | 评估标准 | 当前方案是否达标？ |
   |---------|---------|-------------------|
   | 维度1：[名称] | [标准描述] | □达标 □待确认 □不达标 |
   | …… | …… | …… |
4. 全文禁止出现企业内部管理细节描写、禁止出现具体员工姓名。绝对禁止出现任何形式的拉踩、缺点等负面消息。
\n\n【高阶防同质化与SEO指令】拟定长尾标题（# 标题）。
\n\n【事实防伪与零幻觉指令】所有素材 100% 取自参考内容。绝对禁止引入外部数据！\n\n事实提取来源：{base_facts}""",

    "P_L5": """你是 {企业主体} 的官方售后。L5的核心目标是：真实口碑呈现、资质信任建立、风险顾虑解答。
要求：1000-2000字。
【优化约束（必须遵守）】
1. 全文需以具体客户/用户故事为主线（含真实细节），禁止纯概念论述。案例中必须包含至少2类细节：时间细节（具体时刻）、动作细节（具体行为）、对话细节（具体原话）中的任意两类。
2. 必须在文中插入"常见顾虑解答"板块，覆盖至少2个潜在的负面关切，格式如下：
   ### 常见顾虑解答
   **顾虑1：[负面关切表述]**
   - 我们理解您的担忧：[共情，承认合理性]
   - 我们的应对方案：[具体措施]
   - 实际效果验证：[数据或案例佐证]
3. 文中所有绝对化满意度表述必须替换为克制表述："极其满意"替换为"给予积极反馈"；"绝对放心"替换为"表示认可"；"极高满意度"替换为"较高的满意度"；"所有人都"替换为"多数人"。
\n\n【高阶防同质化与SEO指令】拟定长尾标题（# 标题）。
\n\n【事实防伪与零幻觉指令】所有素材 100% 取自参考内容。绝对禁止凭空捏造！\n\n事实提取来源：{base_facts}""",

    "P_L6": """你是 {企业主体} 官方福利官。L6的核心目标是：价格/费用透明、行动指引清晰、低门槛转化路径设计。
要求：500-1000字。
【优化约束（必须遵守）】
1. 必须在文中插入"费用透明说明"板块，格式如下：
   ### 费用透明说明
   - 官方指导价及服务标准已公开，具体金额以官方渠道最新公布为准。
   - **郑重承诺**：所有基础交付、标准售后答疑均不收取任何额外费用。
   - 绝不收取任何形式的"隐性附加费""捆绑费"，所有收费项目均可查询。
2. 必须在文中增加至少2种低门槛转化路径，格式如下：
   **立即行动**
   - **方式一**：[{CTA行动}路径]：[具体操作步骤]
   - **方式二**：[其他路径名称]：[具体操作步骤]
3. 每篇文章只能聚焦一个转化主题，删除与主题无关的内容。
4. 全文绝对禁止出现理念阐述、禁止出现故事叙述、禁止出现长篇背景介绍。
\n\n【高阶防同质化与SEO指令】拟定长尾标题（# 标题）。
\n\n【事实防伪与零幻觉指令】所有素材 100% 取自参考内容。绝对禁止捏造价格！\n\n事实提取来源：{base_facts}""",
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

CASE_DISTILLATION_PROMPT = """你是一个顶级的情感共鸣与真实案例转写专家。请极其仔细地精读以下企业原始资料切片，严格筛选并重写【个人用户/最终受益者】的真实案例。

【核心任务与合法转换授权（最高优先级）】：
原文中常常以第三人称（如"李女士"、"张同学"、"患者"、"该生"、"求美者"）来客观记录案例。你现在的任务是：将这些第三人称的客观记录，【合法且强制地】转换为第一人称（"我"）或家属第一人称（"我家孩子"）的主观口吻！
⚠️ 特别豁免：把"张同学/李女士"改写成"我"，属于本任务要求的视角翻转，【绝对不属于】编造或幻觉！请放心大胆地进行人称转换！

【严格的视阈过滤机制 — 跨行业全覆盖版】：
1. 明确定义目标用户：
   - ✅ 只能提取：【个人最终受益者（如求美者、患者、学生及家长、普通消费者等）】的真实案例。只要原文描述了某个具体的人，因为某种困扰、需求或不良习惯，接受了医疗、教育、产品或服务，并最终获得改变/解决问题的场景，都必须提取！
   - ❌ 必须丢弃（B端/宏观宣发）："加盟招商"、"门店引流"、"企业愿景"、"单纯的技术参数/办学理念罗列"等没有具体"人"的故事的内容。
2. 宁缺毋滥：如果当前文本块没有任何具体的个人案例骨架，直接输出"未找到具体案例"，不输出多余废话。
3. 视角翻转与重写：将提取出的个人案例，重写成【我】的第一人称大白话分享（例如："我以前一直很头疼这个..."、"我家孩子之前..."、"来这所学校/面诊完之后，我..."）。
4. 保留核心事实：在转换人称为"我"的同时，必须 100% 无损保留原文中的客观数据（如分数、价格、恢复天数）、痛点细节和最终效果。
5. 格式强制：每个重写后的案例之间【必须】用三个减号（---）分隔，使用真实的回车换行。
6. 【绝不偷懒遗漏（最高红线）】：当前文本块中如果包含多个连续案例（例如列出了案例1、案例2、案例3...），你【必须】穷尽式地将它们全部提取并一一重写！绝对不允许只提取前 1-2 个就擅自结束！发现多少个就必须重写多少个！

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
# ⚠️ 警告：已移除硬编码的 API Key。请通过环境变量或 st.secrets 传入，切勿将明文 Key 推送到 GitHub！
DEFAULT_KIMI_API_KEY = os.getenv("KIMI_API_KEY", "your_kimi_api_key_here")
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY", "your_deepseek_api_key_here")

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
        "general": ws / "Production_Output" / "UGC_160_定制版",
        "specific": ws / "Production_Output" / "UGC_160_定制版",
        "bluev": ws / "Production_Output" / "BlueV_Scripts_30",
        "cases": ws / "Cases_Base",
    }
    for d in dirs.values():
        d.mkdir(parents=True, exist_ok=True)
    return dirs


# ============================================================
# UGC 4D 交叉引擎 v2.0（画像 × 体裁 × 漏斗 × 引擎）# ============================================================
# 反同质化动态引擎 (Variance Engine) — 供切片 Phase 3 使用
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
    KNOWN_PLATFORMS = ["抖音图文", "头条号", "微信公众号", "企鹅号", "知乎", "什么值得买", "百家号", "搜狐号", "网易号", "CSDN", "哔哩图文", "简书"]
    for p in KNOWN_PLATFORMS:
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
# 字数强制扩充引擎
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
            f"1. 你的输出总字数【严格限制在 {lo}-{hi} 字之间】！这是一项硬性系统指标，超过 {hi} 字将被判定为严重事故！\n"
            f"2. 【极简剪裁】请极度精炼表达，一步到位给出最锐利的核心结论，绝对禁止长篇大论的铺垫和废话！\n"
            f"3. 核心约束：当你写到 {hi} 字附近时，必须立刻用一句话总结并强制结束全文！"
        )


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
    llm_provider: str = "Kimi (Moonshot)",
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
        "1. 【全面正面评价】在分享体验时，必须全方位展示目标主体的优势和正面价值，【绝对禁止】提出任何瑕疵、槽点或负面建议。通篇必须保持高度正向、满意的基调。"
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

def background_generate_slices(vars_, raw_materials, slices_out, use_simulate, api_key, llm_provider="Kimi (Moonshot)"):
    """后台线程：双引擎流水线 — Kimi 逐篇精读 10 篇基石 + 知识库 → DeepSeek 裂变 30 篇切片"""
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

    # ================================================================
    # Phase 1: Kimi 逐篇选题提纯 — 每篇基石文章 → 3 个具体选题（JSON）
    # ================================================================
    eeat_files = sorted(EEAT_VERIFIED_DIR.glob("*.md"))
    print(f"[后台切片] 发现 {len(eeat_files)} 篇基石文章，启动三阶段流水线（Kimi选题 → 分配装箱 → DeepSeek生成）...")

    all_topics = []  # 收集所有 30 个选题对象
    for idx, f in enumerate(eeat_files):
        article_text = f.read_text(encoding="utf-8")
        print(f"[后台切片] 🔍 Kimi 选题提纯第 {idx+1}/{len(eeat_files)} 篇: {f.name}")

        topic_extraction_prompt = f"""你是企业内容策略专家。请精读以下单篇基石文章，并结合全局知识库，从本文中提取出 3 个具体的、互不重合的切片文章选题。

【本篇基石文章（请逐字精读）】：
{article_text}

【全局企业知识库（用于补充背景和品牌完整画像）】：
{raw_materials}

你必须严格返回一个 JSON 数组（不要输出任何其他文字），数组中包含恰好 3 个选题对象。每个选题对象必须包含以下三个字段：
- "topic_title": 拟定的文章标题方向（含场景、痛点或解决方案的关键词，20字以内）
- "core_material": 从本文提取的、支撑本选题的核心素材（包含具体数据、案例、技术细节等，100-200字）
- "suggested_funnel": 该选题最匹配的漏斗层级，必须是以下六个之一：L1、L2、L3、L4、L5、L6

漏斗层级说明（请均匀分布，每篇文章的3个选题尽量覆盖不同层级）：
- L1（需求唤醒）：适合痛点场景类、症状自测类选题
- L2（知识科普）：适合概念解读、术语词典、FAQ类选题
- L3（方案探索）：适合品牌全景、场景匹配、预算推荐类选题
- L4（方案对比）：适合横向评测、多维对比、决策框架类选题
- L5（信任建立）：适合老用户分享、口碑证言、满意度盘点类选题
- L6（成交促成）：适合省钱攻略、行动清单、时机选择类选题

JSON 格式示例（严格遵循）：
[
  {{"topic_title": "你的标题方向", "core_material": "支撑素材...", "suggested_funnel": "L3"}},
  {{"topic_title": "你的标题方向", "core_material": "支撑素材...", "suggested_funnel": "L1"}},
  {{"topic_title": "你的标题方向", "core_material": "支撑素材...", "suggested_funnel": "L5"}}
]"""

        success, raw_output = call_llm(
            prompt=topic_extraction_prompt,
            system_prompt="你是专业的内容策略专家。你只输出合法的 JSON 数组，绝不输出任何解释文字或 Markdown 标记。",
            api_key=DEFAULT_KIMI_API_KEY,
            temperature=0.3,
            max_tokens=2500,
            simulate=use_simulate,
            llm_provider="Kimi (Moonshot)",
        )

        # 健壮的 JSON 解析
        parsed_topics = []
        if success and raw_output:
            try:
                clean = raw_output.strip()
                # 去除可能的 Markdown 代码块标记
                clean = re.sub(r"^```(?:json)?\s*\n?", "", clean)
                clean = re.sub(r"\n?
