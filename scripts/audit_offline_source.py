"""Audit tracked application code; output counts and paths, never candidate secrets."""
import ast
import json
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "app.py"
text = SOURCE.read_text()
tree = ast.parse(text)
name_pattern = re.compile(r"api_?key|token|secret|password", re.I)


def nonempty_string(node):
    return isinstance(node, ast.Constant) and isinstance(node.value, str) and bool(node.value.strip())


credential_count = 0
for node in ast.walk(tree):
    if isinstance(node, (ast.Assign, ast.AnnAssign)):
        targets = node.targets if isinstance(node, ast.Assign) else [node.target]
        if any(isinstance(t, ast.Name) and name_pattern.search(t.id) for t in targets) and nonempty_string(node.value):
            credential_count += 1
    elif isinstance(node, ast.Dict):
        for key, value in zip(node.keys, node.values):
            if isinstance(key, ast.Constant) and isinstance(key.value, str) and name_pattern.search(key.value) and nonempty_string(value):
                credential_count += 1
    elif isinstance(node, ast.keyword):
        if node.arg and name_pattern.search(node.arg) and nonempty_string(node.value):
            credential_count += 1

provider_patterns = sum(len(re.findall(p, text)) for p in (
    r"sk-[A-Za-z0-9_-]{12,}", r"(?i)Bearer\s+[A-Za-z0-9_.-]{16,}",
))
resource_count = len(re.findall(r"@import\s+url", text))
for node in ast.walk(tree):
    if (isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute)
            and isinstance(node.func.value, ast.Name) and node.func.value.id == "st"
            and node.func.attr in {"image", "audio", "video"} and node.args
            and isinstance(node.args[0], ast.Constant) and isinstance(node.args[0].value, str)
            and node.args[0].value.startswith(("http://", "https://"))):
        resource_count += 1

result = {
    "paths": ["app.py", ".streamlit/config.toml"],
    "hardcoded_credential_candidates": credential_count,
    "credential_pattern_matches": provider_patterns,
    "default_remote_resource_candidates": resource_count,
}
print(json.dumps(result, indent=2))
raise SystemExit(bool(credential_count or provider_patterns or resource_count))
