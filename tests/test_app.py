"""Offline regressions: real Streamlit execution, synthetic files, mocked transport."""
import contextlib
import io
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import Mock, patch

from streamlit.testing.v1 import AppTest


SOURCE = Path(__file__).resolve().parents[1] / "app.py"
FUNCTIONS = (
    "get_workspace_dir", "get_workspace_paths", "safe_write_file",
    "call_llm", "background_generate_slices", "build_export_filename",
    "extract_text_from_upload", "convert_md_to_docx_bytes", "scan_all_files",
)


class AppTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        app = self.root / "app.py"
        # AppTest executes the actual app in a disposable directory. Export the
        # actual function objects for direct regressions without importing it twice.
        app.write_text(SOURCE.read_text() + "\n\nst.session_state['_test_functions'] = "
                       + "{name: globals()[name] for name in " + repr(FUNCTIONS) + "}\n")
        self.addCleanup(patch.stopall)
        self.connect = patch("socket.socket.connect", side_effect=AssertionError("Network forbidden")).start()
        self.create_connection = patch("socket.create_connection", side_effect=AssertionError("Network forbidden")).start()
        self.provider = patch("openai.OpenAI", side_effect=AssertionError("Live provider forbidden")).start()
        self.at = AppTest.from_file(str(app), default_timeout=15).run()
        self.assertFalse(self.at.exception, [e.message for e in self.at.exception])
        self.api = self.at.session_state["_test_functions"]

    def tearDown(self):
        self.connect.assert_not_called()
        self.create_connection.assert_not_called()
        self.provider.assert_not_called()

    def production(self):
        self.at.sidebar.radio[0].set_value("⚙️ 自动化生产车间").run()
        self.assertFalse(self.at.exception, [e.message for e in self.at.exception])

    def test_both_pages_render_and_default_to_offline(self):
        self.assertTrue(self.at.checkbox(key="simulate_mode").value)
        self.assertGreater(len(self.at.button), 0)
        self.production()
        self.assertGreaterEqual(len(self.at.get("file_uploader")), 5)
        self.assertGreater(len(self.at.button), 5)
        self.assertTrue(any("缺少 ugc_160_prompts" in w.value for w in self.at.warning))

    def test_rendered_page_has_no_remote_resource_defaults(self):
        self.assertEqual(len(self.at.get("imgs")), 0)
        self.assertFalse(any("@import" in m.value for m in self.at.markdown))

    def test_offline_manual_variables_can_be_confirmed(self):
        raw = self.root / "workspaces" / "默认项目" / "Raw_Materials" / "synthetic.txt"
        raw.write_text("Synthetic company facts for offline testing only.")
        self.production()
        self.assertTrue(self.at.session_state["raw_files_loaded"])
        self.assertEqual(len([t for t in self.at.text_input if t.key and t.key.startswith("var_")]), 12)
        self.at.button(key="btn_confirm_vars").click().run()
        self.assertFalse(self.at.exception, [e.message for e in self.at.exception])
        self.assertTrue(self.at.session_state["variables_confirmed"])

    def test_missing_ugc_templates_disable_generation(self):
        paths = self.api["get_workspace_paths"]("默认项目")
        for i in range(30):
            (paths["slices"] / f"synthetic_{i}.md").write_text("# synthetic slice")
        self.production()
        self.assertTrue(self.at.button(key="btn_ugc").disabled)

    def test_invalid_workspace_names_cannot_escape_or_alias_root(self):
        for name in ("", "  ", ".", "..", "../outside", "a/b", "a\\b", "bad\x00name"):
            with self.subTest(name=name), self.assertRaises(ValueError):
                self.api["get_workspace_dir"](name)
        self.assertFalse((self.root / "outside").exists())
        valid = self.api["get_workspace_paths"](" 合成测试项目 ")
        self.assertEqual(valid["ws"].name, "合成测试项目")
        self.assertTrue(all(p.is_dir() and p.is_relative_to(valid["ws"]) for p in valid.values()))

    def test_workspace_and_nested_symlinks_are_rejected(self):
        outside = self.root / "outside"
        outside.mkdir()
        (self.root / "workspaces" / "linked").symlink_to(outside, target_is_directory=True)
        with self.assertRaises(ValueError):
            self.api["get_workspace_dir"]("linked")
        ws = self.api["get_workspace_dir"]("nested")
        (ws / "Production_Output").symlink_to(outside, target_is_directory=True)
        with self.assertRaises(ValueError):
            self.api["get_workspace_paths"]("nested")
        self.assertEqual(list(outside.iterdir()), [])

    def test_output_symlink_cannot_overwrite_another_file(self):
        outside = self.root / "outside.md"
        outside.write_text("original")
        output = self.root / "output"
        output.mkdir()
        (output / "article.md").symlink_to(outside)
        with self.assertRaises(ValueError):
            self.api["safe_write_file"](output, "article.md", "overwrite")
        self.assertEqual(outside.read_text(), "original")
        result = self.api["safe_write_file"](output, "normal.md", "synthetic text")
        self.assertEqual(result.read_text(), "synthetic text")

    def test_workspace_root_symlink_is_rejected(self):
        fn = self.api["get_workspace_dir"]
        outside = self.root / "outside"
        outside.mkdir()
        link = self.root / "linked_root"
        link.symlink_to(outside, target_is_directory=True)
        with patch.dict(fn.__globals__, {"WORKSPACES_ROOT": link}), self.assertRaises(ValueError):
            fn("synthetic project")
        self.assertEqual(list(outside.iterdir()), [])

    def test_invalid_workspace_is_reported_without_ui_exception(self):
        self.at.text_input(key="ws_new_input").set_value("..").run()
        self.assertFalse(self.at.exception)
        self.assertTrue(self.at.error)
        self.assertIn("项目名称", self.at.error[0].value)

    def test_switching_workspace_clears_previous_material_and_variables(self):
        self.production()
        self.at.session_state["raw_text_combined"] = "SYNTHETIC PROJECT A"
        self.at.session_state["raw_files_loaded"] = True
        self.at.session_state["variables"] = {"企业主体": "SYNTHETIC A"}
        self.at.session_state["variables_confirmed"] = True
        self.at.session_state["selected_content"] = "SYNTHETIC OLD ARTICLE"
        self.at.text_input(key="ws_new_input").set_value("SYNTHETIC B").run()
        self.assertFalse(self.at.exception, [e.message for e in self.at.exception])
        for key in ("raw_text_combined", "selected_content"):
            self.assertEqual(self.at.session_state[key], "")
        self.assertEqual(self.at.session_state["variables"], {})
        self.assertFalse(self.at.session_state["variables_confirmed"])
        self.assertFalse(self.at.session_state["raw_files_loaded"])

    def test_real_widget_edits_do_not_follow_workspace_selector(self):
        for name in ("SYNTHETIC A", "SYNTHETIC B"):
            paths = self.api["get_workspace_paths"](name)
            (paths["raw"] / "synthetic.txt").write_text(f"Synthetic facts for {name}")
        self.at.run()
        self.at.selectbox(key="ws_selector").set_value("SYNTHETIC A").run()
        self.production()
        company = next(t for t in self.at.text_input if t.label == "**企业主体**")
        company_a_key = company.key
        company.set_value("SYNTHETIC_ORG_A_ONLY").run()
        self.at.button(key="btn_confirm_vars").click().run()
        self.assertTrue(self.at.session_state["variables_confirmed"])
        self.at.selectbox(key="ws_selector").set_value("SYNTHETIC B").run()
        self.assertFalse(self.at.exception, [e.message for e in self.at.exception])
        company_b = next(t for t in self.at.text_input if t.label == "**企业主体**")
        # The browser can replay an old widget value even after server-side
        # session deletion. Project-specific controls need distinct identities.
        self.assertNotEqual(company_a_key, company_b.key)
        self.assertNotEqual(company_b.value, "SYNTHETIC_ORG_A_ONLY")
        self.assertFalse(self.at.session_state["variables_confirmed"])
        self.assertNotIn("SYNTHETIC A", self.at.session_state["raw_text_combined"])
        self.assertIn("SYNTHETIC B", self.at.session_state["raw_text_combined"])

    def test_simulation_succeeds_without_client_or_credentials(self):
        ok, text = self.api["call_llm"]("测试内容，要求 200-250 字", simulate=True, llm_provider="Kimi (Moonshot)")
        self.assertTrue(ok)
        self.assertIn("模拟生成", text)
        self.assertIn("moonshot-v1-128k", text)

    def test_real_mode_without_key_fails_instead_of_faking_success(self):
        ok, text = self.api["call_llm"]("synthetic prompt", simulate=False, api_key="")
        self.assertFalse(ok)
        self.assertIn("API Key", text)

    def test_token_limit_reaches_mock_transport(self):
        fn = self.api["call_llm"]
        client = Mock()
        client.chat.completions.create.return_value.choices = [
            Mock(finish_reason="stop", message=Mock(content="synthetic response"))
        ]
        constructor = Mock(return_value=client)
        with patch.dict(fn.__globals__, {"OpenAI": constructor}):
            ok, text = fn("synthetic prompt", simulate=False, api_key="unit-test-placeholder", max_tokens=1234)
        self.assertTrue(ok)
        self.assertEqual(text, "synthetic response")
        self.assertEqual(client.chat.completions.create.call_args.kwargs["max_tokens"], 1234)

    def test_slice_worker_uses_explicit_workspace_source(self):
        a = self.api["get_workspace_paths"]("SYNTHETIC A")
        b = self.api["get_workspace_paths"]("SYNTHETIC B")
        for i in range(10):
            (a["eeat_verified"] / f"article_{i}.md").write_text(f"SYNTHETIC_A_FACT_{i}")
        (b["eeat_verified"] / "other.md").write_text("SYNTHETIC_B_FACT")
        calls = []

        def fake_llm(**kwargs):
            calls.append(kwargs)
            if kwargs["max_tokens"] == 2500:
                return True, json.dumps([{"topic_title": f"topic_{j}", "core_material": "synthetic fact", "suggested_funnel": f"L{j+1}"} for j in range(3)])
            return True, "# synthetic output\n\nUnit test placeholder."

        fn = self.api["background_generate_slices"]
        with patch.dict(fn.__globals__, {"EEAT_VERIFIED_DIR": b["eeat_verified"], "call_llm": fake_llm}), patch("time.sleep"), contextlib.redirect_stdout(io.StringIO()):
            fn({}, "synthetic knowledge", a["slices"], True, "", eeat_verified_dir=a["eeat_verified"])
        self.assertEqual(len(list(a["slices"].glob("*.md"))), 30)
        self.assertEqual(list(b["slices"].glob("*.md")), [])
        extraction = [c["prompt"] for c in calls if c["max_tokens"] == 2500]
        self.assertEqual(len(extraction), 10)
        self.assertTrue(all("SYNTHETIC_A_FACT" in p and "SYNTHETIC_B_FACT" not in p for p in extraction))

    def test_text_upload_export_and_workspace_scan(self):
        upload = io.BytesIO("# 合成标题\n\n**合成正文**".encode())
        upload.name = "synthetic.md"
        name, text = self.api["extract_text_from_upload"](upload)
        self.assertEqual(name, "synthetic.md")
        self.assertIn("合成正文", text)
        self.assertEqual(self.api["build_export_filename"](text, "Slice_L2_知乎", "md"), "【L2-知乎】合成标题.md")
        docx = self.api["convert_md_to_docx_bytes"](text)
        self.assertTrue(docx.startswith(b"PK"))
        a = self.api["get_workspace_paths"]("scan A")
        b = self.api["get_workspace_paths"]("scan B")
        self.api["safe_write_file"](a["slices"], "Slice_L2_知乎.md", text)
        self.api["safe_write_file"](b["slices"], "other.md", "private synthetic B")
        (a["slices"] / "linked.md").symlink_to(b["slices"] / "other.md")
        articles = self.api["scan_all_files"](str(a["ws"]))
        self.assertEqual(len(articles), 1)
        self.assertEqual(articles[0]["title"], "合成标题")
        self.assertEqual(articles[0]["funnel"], "L2")


if __name__ == "__main__":
    unittest.main()
