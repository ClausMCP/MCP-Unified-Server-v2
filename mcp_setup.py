#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MCP Setup Helper v6.9 – гибридная установка с автоматическим переключением зеркал,
увеличенными таймаутами и поддержкой обновления pip/setuptools через зеркала.
"""

import os
import sys
import ast
import json
import shutil
import subprocess
import argparse
import time
from pathlib import Path
from datetime import datetime

ROOT = Path(__file__).resolve().parent
VENV = ROOT / ".venv"
PY_EXE = str(VENV / "Scripts" / "python.exe") if sys.platform == "win32" else str(VENV / "bin" / "python3")
PIP_CMD = [PY_EXE, "-m", "pip", "--no-input"]
DEPS_DIR = ROOT / "python_deps"

BASE_DEPS = {
    "watchdog", "psutil", "requests", "xxhash", "cryptography", "keyring",
    "beautifulsoup4", "feedparser", "icalendar", "openpyxl", "python-docx",
    "python-pptx", "pytesseract", "Pillow", "mutagen", "duckdb", "pyodbc",
    "patool", "py7zr", "rarfile", "playwright", "pandas",
    "sentence-transformers", "chromadb", "pypdf", "pdfplumber", "ebooklib",
    "trafilatura", "readability-lxml", "html-table-takeout",
    "mempalace"
}

# Список зеркал PyPI (в порядке приоритета)
PIP_MIRRORS = [
    "",   # официальный PyPI (без --index-url)
    "https://pypi.tuna.tsinghua.edu.cn/simple",
    "https://mirrors.aliyun.com/pypi/simple/",
    "https://mirrors.cloud.tencent.com/pypi/simple",
]

def find_plugin_deps():
    deps = set()
    for search_dir in [ROOT, ROOT / "mcp_plugins"]:
        if not search_dir.is_dir():
            continue
        for py_file in search_dir.glob("*.py"):
            if py_file.name.startswith("_"):
                continue
            try:
                tree = ast.parse(py_file.read_text(encoding="utf-8"))
                for node in ast.walk(tree):
                    if isinstance(node, ast.Assign):
                        for target in node.targets:
                            if isinstance(target, ast.Name) and target.id == "__mcp_plugin__":
                                if isinstance(node.value, ast.Dict):
                                    keys = [k.value for k in node.value.keys if isinstance(k, ast.Constant)]
                                    if "dependencies" in keys:
                                        idx = keys.index("dependencies")
                                        val = node.value.values[idx]
                                        if isinstance(val, ast.List):
                                            for elem in val.elts:
                                                raw = getattr(elem, 'value', getattr(elem, 's', ''))
                                                if isinstance(raw, str):
                                                    pkg = raw.split("^")[0].split("~")[0].split("=")[0].strip()
                                                    deps.add(pkg)
            except Exception as e:
                print(f"⚠️ Предупреждение: {py_file.name}: {e}")
    return deps

def get_full_deps():
    return sorted(BASE_DEPS | find_plugin_deps())

def run(cmd, check=False, env=None, timeout=120):
    print(f">>> {' '.join(cmd)}")
    proc = subprocess.run(cmd, capture_output=True, text=True, env=env, timeout=timeout)
    if proc.stdout:
        print(proc.stdout)
    if proc.stderr:
        print(proc.stderr, file=sys.stderr)
    if check and proc.returncode != 0:
        sys.exit(proc.returncode)
    return proc

def ensure_venv():
    if not VENV.exists():
        print("📦 Создаю виртуальное окружение...")
        run([sys.executable, "-m", "venv", str(VENV)], check=True)
    if not Path(PY_EXE).exists():
        print(f"❌ Ошибка: {PY_EXE} не найден")
        sys.exit(1)

def check_pip():
    result = subprocess.run([PY_EXE, "-m", "pip", "--version"], capture_output=True)
    if result.returncode != 0:
        print("❌ Pip не найден в виртуальном окружении.")
        print("   Запустите опцию 1 (Recreate virtual environment) в setup.bat")
        sys.exit(1)
    return True

def check_import(package: str) -> bool:
    import_name = package.split('[')[0].replace('-', '_')
    mapping = {
        "beautifulsoup4": "bs4",
        "Pillow": "PIL",
        "python_docx": "docx",
        "python_pptx": "pptx",
        "patool": None,
        "readability_lxml": "readability",
        "html_table_takeout": "html_table_takeout",
        "mempalace": "mempalace"
    }
    if import_name in mapping:
        if mapping[import_name] is None:
            return True
        import_name = mapping[import_name]

    if not import_name.isidentifier():
        print(f"⚠️ Пропущен небезопасный импорт: {import_name}")
        return False

    try:
        subprocess.run(
            [PY_EXE, "-c", f"import {import_name}"],
            capture_output=True, check=True, text=True, timeout=30
        )
        return True
    except subprocess.CalledProcessError:
        return False

def install_with_mirrors(packages, upgrade=False, timeout=120):
    if not packages:
        return True
    cmd_base = PIP_CMD + (["install", "--upgrade"] if upgrade else ["install"])
    cmd_base += ["--timeout", str(timeout)]
    for mirror in PIP_MIRRORS:
        cmd = cmd_base.copy()
        if mirror:
            cmd += ["--index-url", mirror]
        cmd += packages
        print(f"🌐 Пробую зеркало: {mirror or 'официальный PyPI'}")
        proc = run(cmd, timeout=timeout)
        if proc.returncode == 0:
            return True
        print(f"⚠️ Зеркало {mirror or 'официальный'} недоступно, пробую следующее...")
        time.sleep(2)
    return False

def download_with_mirrors(packages, timeout=120):
    if not packages:
        return True
    DEPS_DIR.mkdir(exist_ok=True)
    for mirror in PIP_MIRRORS:
        cmd = PIP_CMD + ["download", "-d", str(DEPS_DIR), "--prefer-binary", "--timeout", str(timeout)]
        if mirror:
            cmd += ["--index-url", mirror]
        cmd += packages
        print(f"🌐 Скачиваю через зеркало: {mirror or 'официальный PyPI'}")
        proc = run(cmd, timeout=timeout)
        if proc.returncode == 0:
            return True
        print(f"⚠️ Зеркало {mirror or 'официальный'} недоступно, пробую следующее...")
        time.sleep(2)
    return False

def ensure_playwright_browsers():
    try:
        subprocess.run([PY_EXE, "-m", "playwright", "install", "chromium"],
                       capture_output=True, check=True, timeout=300)
        print("✅ Браузеры Playwright установлены.")
    except subprocess.CalledProcessError:
        print("⚠️ Не удалось автоматически установить браузеры Playwright.")
        print("   Запустите позже: %PY_EXE% -m playwright install chromium")
    except FileNotFoundError:
        pass

def check_and_install_missing():
    ensure_venv()
    check_pip()
    deps = get_full_deps()
    missing = [dep for dep in deps if not check_import(dep)]
    if not missing:
        print("✅ Все зависимости уже установлены.")
        ensure_playwright_browsers()
        return
    print(f"⚠️ Отсутствуют {len(missing)} пакетов: {', '.join(missing)}")
    print("🔧 Устанавливаю недостающие пакеты...")

    local_whls = set()
    if DEPS_DIR.exists():
        for whl in DEPS_DIR.glob("*.whl"):
            pkg_name = whl.stem.split('-')[0].lower()
            local_whls.add(pkg_name)

    online_pkgs = []
    offline_pkgs = []
    for pkg in missing:
        if pkg.lower() in local_whls:
            offline_pkgs.append(pkg)
        else:
            online_pkgs.append(pkg)

    if offline_pkgs:
        print(f"📦 Устанавливаю из локальной папки: {', '.join(offline_pkgs)}")
        cmd = PIP_CMD + ["install", "--no-index", "--find-links", str(DEPS_DIR),
                         "--no-build-isolation"] + offline_pkgs
        if run(cmd).returncode != 0:
            print("❌ Ошибка при установке из локальной папки.")
            sys.exit(1)

    if online_pkgs:
        print(f"🌐 Устанавливаю из интернета: {', '.join(online_pkgs)}")
        if not install_with_mirrors(online_pkgs, timeout=120):
            print("❌ Ошибка при установке из интернета (все зеркала недоступны).")
            sys.exit(1)

    ensure_playwright_browsers()
    print("✅ Все зависимости успешно установлены.")

def online_mode():
    ensure_venv()
    check_pip()
    deps = get_full_deps()
    print(f"🌐 Скачиваю {len(deps)} пакетов...")
    # Обновляем pip и setuptools через зеркала
    if not install_with_mirrors(["pip", "setuptools", "wheel"], upgrade=True, timeout=120):
        print("❌ Не удалось обновить pip/setuptools/wheel (все зеркала недоступны)")
        sys.exit(1)
    if not download_with_mirrors(deps):
        print("❌ Ошибка при скачивании пакетов")
        sys.exit(1)
    # Дополнительно скачиваем pip/setuptools/wheel
    download_with_mirrors(["pip", "setuptools", "wheel"])
    print(f"✅ Пакеты скачаны в {DEPS_DIR}")
    print("Запустите mcp_setup.py --offline для установки.")

def offline_mode():
    if not DEPS_DIR.exists() or not any(DEPS_DIR.glob("*.whl")):
        print("❌ Папка python_deps пуста. Сначала --online")
        sys.exit(1)
    ensure_venv()
    check_pip()
    deps = get_full_deps()
    print(f"📦 Устанавливаю {len(deps)} пакетов из {DEPS_DIR}...")
    upgrade_cmd = PIP_CMD + ["install", "--no-index", "--find-links", str(DEPS_DIR),
                             "--upgrade", "pip", "setuptools", "wheel"]
    if run(upgrade_cmd).returncode != 0:
        print("❌ Ошибка при обновлении pip")
        sys.exit(1)
    install_cmd = PIP_CMD + ["install", "--no-index", "--find-links", str(DEPS_DIR),
                             "--no-build-isolation"] + deps
    if run(install_cmd).returncode != 0:
        print("❌ Ошибка при установке пакетов")
        sys.exit(1)
    ensure_playwright_browsers()
    print("✅ Зависимости установлены.")

def fix_config(config_path: str, python_exe: str):
    config_file = Path(config_path)
    if not config_file.exists():
        print(f"❌ Файл не найден: {config_path}")
        return 1

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = config_file.with_name(f"{config_file.name}.backup_{ts}")
    try:
        shutil.copy2(config_file, backup)
        print(f"💾 Бэкап: {backup}")
    except Exception as e:
        print(f"❌ Не удалось создать бэкап: {e}")
        return 1

    try:
        data = json.loads(config_file.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        print(f"❌ Ошибка парсинга JSON: {e}")
        return 1
    except Exception as e:
        print(f"❌ Ошибка чтения: {e}")
        return 1

    counter = {"n": 0}

    def walk(obj):
        if isinstance(obj, dict):
            for k, v in list(obj.items()):
                if k == "command" and isinstance(v, str) and "python" in v.lower():
                    obj[k] = python_exe
                    counter["n"] += 1
                else:
                    walk(v)
        elif isinstance(obj, list):
            for item in obj:
                walk(item)

    walk(data)

    if counter["n"] == 0:
        print("ℹ️ В конфиге не найдено полей 'command' со словом 'python'.")
    else:
        try:
            config_file.write_text(
                json.dumps(data, indent=2, ensure_ascii=False),
                encoding="utf-8"
            )
            print(f"✅ Обновлено записей: {counter['n']}")
            print(f"✅ Конфиг сохранён: {config_path}")
        except Exception as e:
            print(f"❌ Ошибка записи: {e}")
            return 1
    return 0

def main():
    parser = argparse.ArgumentParser(description="MCP Setup Helper")
    group = parser.add_mutually_exclusive_group(required=False)
    group.add_argument("--online", action="store_true", help="Скачать пакеты в python_deps")
    group.add_argument("--offline", action="store_true", help="Установить из python_deps")
    group.add_argument("--check", action="store_true", help="Проверить и доустановить недостающее")
    parser.add_argument("--fix-config", nargs=2, metavar=("CONFIG", "PYTHON_EXE"),
                        help="Заменить пути к python в JSON-конфиге")

    if len(sys.argv) == 1:
        parser.print_help()
        input("\nНажмите Enter для выхода...")
        sys.exit(0)

    args = parser.parse_args()

    if args.fix_config:
        cfg_path, py_exe = args.fix_config
        sys.exit(fix_config(cfg_path, py_exe))
    elif args.online:
        online_mode()
    elif args.offline:
        offline_mode()
    elif args.check:
        check_and_install_missing()
    else:
        parser.print_help()

if __name__ == "__main__":
    main()
