# geo-saas-mvp
<img width="1570" height="933" alt="image" src="https://github.com/user-attachments/assets/cbc62a17-f389-473d-bcf9-351c726de90a" />
自动裂变文章 拥有详细的提示词生成GEO文章
适用于GEO文章产出，GEO文章权重的裂变 

## 本地运行与验证

本轮验证环境：Python 3.13.3；依赖版本见 `requirements.txt`。

```sh
python3 -m pip install -r requirements.txt
python3 -m streamlit run app.py
```

页面默认勾选**模拟模式**，生成的是明确标注的占位内容。取消模拟模式后，真实生成需要自行配置对应引擎的环境变量 `KIMI_API_KEY` / `DEEPSEEK_API_KEY`；缺失凭据会返回错误，不会假装成功。不要提交凭据或客户资料。

无需凭据的回归测试：

```sh
python3 -m py_compile app.py tests/test_app.py
sh scripts/test_offline.sh
git diff --check
```

测试在临时工作区使用合成资料，清空继承的环境变量与用户 HOME，拦截网络和真实 OpenAI 客户端。只对请求参数的测试使用内存模拟客户端；不启动 Streamlit 服务。

## 恢复来源与范围

`24f682cb3791d115bab653a89f4b2aa59a58f9b0` 的 `app.py` 在第 793 行截断，无法解析。缺失的后半段来自其直接父提交 `dd06cd005cad47baa43341e25be6e73610af0e49`，保留了最新提交中移除硬编码凭据的修改。恢复了资产大盘、生产车间、后台函数和模拟生成函数，并补充路径隔离、线程路径传参、离线默认值和基本 CI 配置。

仓库**不包含** `ugc_160_prompts.txt`（或编号变体），160 篇 UGC 生成会提示缺失并禁用按钮，需要用户提供已经确认的原始模板。模拟变量提取没有真实语义能力，可在离线模式下手动填写确认。真实模型质量、费用、并发长任务和端到端 200 篇生产没有在本轮验证。

完整证据与限制见 [LOCAL_REPAIR.md](LOCAL_REPAIR.md) 和 [evidence/](evidence/)。
