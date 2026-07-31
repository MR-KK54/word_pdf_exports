import builtins
import importlib
import sys


def test_web_app_imports_without_windows_com_modules(monkeypatch):
    real_import = builtins.__import__

    def fake_import(name, *args, **kwargs):
        if name in {"win32com", "win32com.client", "pythoncom"}:
            raise ImportError("simulated missing Windows COM module")
        return real_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, "__import__", fake_import)

    for name in ["word_exporter_pro.core.com_engine", "word_exporter_pro.web.app"]:
        sys.modules.pop(name, None)

    module = importlib.import_module("word_exporter_pro.web.app")

    assert hasattr(module, "app")
    assert callable(module.main)
