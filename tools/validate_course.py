from __future__ import annotations

import json
import subprocess
from html.parser import HTMLParser
from pathlib import Path
from urllib.parse import urlparse

ROOT = Path(__file__).resolve().parents[1]
NODE = Path("/Users/anshuman/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node")


class AuditParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.title_parts: list[str] = []
        self.in_title = False
        self.stage_count = 0
        self.choice_count = 0
        self.links: list[str] = []
        self.scripts: list[str] = []
        self.in_script = False
        self.script_parts: list[str] = []
        self.text_parts: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        values = dict(attrs)
        classes = set((values.get("class") or "").split())
        if tag == "title":
            self.in_title = True
        if tag == "section" and "stage" in classes:
            self.stage_count += 1
        if tag == "button" and "choice" in classes:
            self.choice_count += 1
        if tag == "a" and values.get("href"):
            self.links.append(values["href"] or "")
        if tag == "script" and not values.get("src"):
            self.in_script = True
            self.script_parts = []

    def handle_endtag(self, tag: str) -> None:
        if tag == "title":
            self.in_title = False
        if tag == "script" and self.in_script:
            self.scripts.append("".join(self.script_parts))
            self.in_script = False
            self.script_parts = []

    def handle_data(self, data: str) -> None:
        if self.in_title:
            self.title_parts.append(data)
        if self.in_script:
            self.script_parts.append(data)
        else:
            self.text_parts.append(data)

    @property
    def title(self) -> str:
        return "".join(self.title_parts).strip()


def fail(errors: list[str], message: str) -> None:
    errors.append(message)


def main() -> None:
    errors: list[str] = []
    modules = sorted(ROOT.glob("module-*.html"))
    if len(modules) != 69:
        fail(errors, f"Expected 69 course modules, found {len(modules)}")

    expected = [f"module-{i:02d}.html" for i in range(1, 70)]
    names = [path.name for path in modules]
    if names != expected:
        fail(errors, "Module filenames do not form the complete 01 through 69 sequence")

    html_files = [ROOT / "index.html", *modules]
    for path in html_files:
        text = path.read_text(encoding="utf-8")
        if "—" in text:
            fail(errors, f"Em dash found in {path.relative_to(ROOT)}")
        parser = AuditParser()
        parser.feed(text)
        if not parser.title:
            fail(errors, f"Missing title in {path.relative_to(ROOT)}")
        if path.name.startswith("module-"):
            module_id = int(path.stem.split("-")[1])
            if f"Module {module_id:02d}" not in parser.title:
                fail(errors, f"Wrong title number in {path.name}")
            if parser.stage_count != 5:
                fail(errors, f"Expected five stages in {path.name}")
            if parser.choice_count < 6:
                fail(errors, f"Expected at least six Core and Lab choices in {path.name}")
            if "Source check: statutory anchors" not in "".join(parser.text_parts):
                fail(errors, f"Missing source-check statement in {path.name}")
            if "STRATEGY LAB" not in "".join(parser.text_parts):
                fail(errors, f"Missing strategic Lab layer in {path.name}")
            if "index.html" not in parser.links:
                fail(errors, f"Missing course-map link in {path.name}")

        for href in parser.links:
            parsed = urlparse(href)
            if parsed.scheme or href.startswith("#") or href.startswith("mailto:"):
                continue
            target = (path.parent / parsed.path).resolve()
            if not target.exists():
                fail(errors, f"Broken internal link in {path.relative_to(ROOT)}: {href}")

        for index, script in enumerate(parser.scripts, 1):
            checked = subprocess.run(
                [str(NODE), "--check", "-"],
                input=script,
                text=True,
                capture_output=True,
                check=False,
            )
            if checked.returncode:
                fail(errors, f"JavaScript syntax error in {path.relative_to(ROOT)} script {index}: {checked.stderr.strip()}")

    manifest_path = ROOT / "legal-source-manifest.json"
    if not manifest_path.exists():
        fail(errors, "Missing legal-source-manifest.json")
    else:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        ids = [item["id"] for item in manifest.get("modules", [])]
        if ids != list(range(1, 70)):
            fail(errors, "Legal source manifest does not cover Modules 1 through 69")

    if errors:
        print("VALIDATION FAILED")
        for error in errors:
            print(f"- {error}")
        raise SystemExit(1)
    print("VALIDATION PASSED")
    print(f"- {len(modules)} generated modules")
    print(f"- {len(html_files)} HTML files checked")
    print("- internal links, visible em-dash ban, source notes, structure and JavaScript syntax checked")


if __name__ == "__main__":
    main()
