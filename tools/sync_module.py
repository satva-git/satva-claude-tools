"""Copy the helper module out of SKILL.md into satva_brand.py, so both stay identical.
Run from the repo root after editing SKILL.md:  python tools/sync_module.py"""
import pathlib
d = pathlib.Path("plugins/satva-brand-format/skills/satva-brand-format")
s = (d / "SKILL.md").read_text(encoding="utf-8")
code = s.split("```python\n", 1)[1].rsplit("```", 1)[0]
(d / "satva_brand.py").write_text(code, encoding="utf-8")
print("satva_brand.py updated from SKILL.md")
