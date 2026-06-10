import importlib.util
import io
import os
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest.mock import patch

from sharepoint_excel_sync import CONFIG_ENV_VAR


REPO_ROOT = Path(__file__).resolve().parents[1]


def read_repo_file(relative_path: str) -> str:
    return (REPO_ROOT / relative_path).read_text(encoding="utf-8")


def load_probe_module():
    path = REPO_ROOT / "scripts" / "sharepoint_workbook_probe.py"
    spec = importlib.util.spec_from_file_location("sharepoint_workbook_probe_for_tests", path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


class ToolingAndPackagingTests(unittest.TestCase):
    def test_macos_requirements_include_sharepoint_runtime_dependencies(self):
        requirements = read_repo_file("macOS/requirements-macos.txt").splitlines()

        self.assertIn("msal", requirements)
        self.assertIn("requests", requirements)

    def test_windows_and_macos_builds_include_lazy_graph_imports(self):
        windows_build = read_repo_file("build_exe.ps1")
        macos_workflow = read_repo_file(".github/workflows/build-macos-app.yml")

        for content in (windows_build, macos_workflow):
            self.assertIn("--hidden-import msal", content)
            self.assertIn("--hidden-import requests", content)

    def test_gitignore_protects_sharepoint_config_and_token_cache(self):
        gitignore = read_repo_file(".gitignore")

        self.assertIn("sharepoint_config.json", gitignore)
        self.assertIn(".sharepoint_token_cache.json", gitignore)
        self.assertIn("Reference Data/SharePoint/", gitignore)

    def test_docs_and_example_config_explain_required_runtime_contract(self):
        docs = read_repo_file("docs/sharepoint-excel-sync.md")
        example_config = read_repo_file("sharepoint_config.example.json")

        self.assertIn("ConstructionProjects", docs)
        self.assertIn("OilGasProjects", docs)
        self.assertIn("CompletedProjects", docs)
        self.assertIn("Record ID", docs)
        self.assertIn("Files.ReadWrite", docs)
        self.assertIn('"enabled": false', example_config)
        self.assertIn('"sharing_url"', example_config)

    def test_probe_script_returns_clear_failure_when_sharepoint_config_is_disabled(self):
        probe = load_probe_module()
        with tempfile.TemporaryDirectory() as tmp:
            missing_config = Path(tmp) / "missing-sharepoint.json"
            output = io.StringIO()
            with patch.dict(os.environ, {CONFIG_ENV_VAR: str(missing_config)}, clear=False):
                with patch.object(sys, "argv", ["sharepoint_workbook_probe.py"]):
                    with redirect_stdout(output):
                        result = probe.main()

        self.assertEqual(result, 2)
        self.assertIn("SharePoint sync is not enabled", output.getvalue())


if __name__ == "__main__":
    unittest.main()
