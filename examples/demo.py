"""Run the actual comparison scripts on fictional products in an isolated workspace.

No model call, network connection, private catalog, or API key is required.
The deterministic picks below are specific to this example, not agent decisions.
"""
import argparse
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")


def run(output):
    output = Path(output).resolve()
    output.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="selection-demo-") as tmp:
        workspace = Path(tmp)
        shutil.copytree(ROOT / "scripts", workspace / "scripts", ignore=shutil.ignore_patterns("__pycache__"))
        refs = workspace / "references"
        refs.mkdir()
        shutil.copyfile(ROOT / "examples" / "catalog.json", refs / "category_products_full_with_internal_model.json")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "需求清单"
        ws.append(["序号", "设备名称", "技术参数", "数量"])
        ws.append([1, "会议摄像机", "4K，12倍光学变焦", 1])
        ws.append([2, "门禁控制器", "门禁系统", 1])
        source = workspace / "demo.xlsx"
        wb.save(source)
        wb.close()
        env = dict(os.environ, PYTHONUTF8="1", PYTHONIOENCODING="utf-8")

        def command(script, *args):
            result = subprocess.run([sys.executable, str(workspace / "scripts" / script), *map(str, args)],
                                    cwd=workspace, env=env, capture_output=True, text=True,
                                    encoding="utf-8", timeout=60)
            print(result.stdout, end="")
            if result.returncode:
                raise RuntimeError(result.stderr or f"Command failed: {script} {args}")

        command("bidan.py", "draft", source)
        draft = workspace / "方案表格" / "demo_底稿.json"
        rows = json.loads(draft.read_text(encoding="utf-8"))["rows"]
        if len(rows) != 2:
            raise RuntimeError("Expected two fictional input rows")
        for row in rows:
            pick = "HT-DEMO-CAM4K" if row["name"] == "会议摄像机" else "空:示例目录不提供门禁产品"
            command("bidan.py", "pick", draft, f"{row['id']}={pick}")
        command("bidan.py", "export", draft)
        generated = draft.with_name("demo_比单表.xlsx")
        target = output / generated.name
        shutil.copyfile(generated, target)
        print(f"演示文件：{target}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", default=str(ROOT / "demo-output"))
    run(parser.parse_args().output)
