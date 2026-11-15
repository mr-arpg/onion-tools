"""Patch customUI.xml into a PowerPoint .ppam or .pptm (Office Open XML zip)."""
import re
import zipfile
import shutil
import os
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.dirname(SCRIPT_DIR)
XML_SRC = os.path.join(REPO_ROOT, "src", "ribbon", "customUI.xml")
XML_DEST = "customUI/customUI14.xml"
# customUI14 (2009 schema) must use the 2007 relationship and the standard
# application/xml content type (via <Default Extension="xml" .../>). Adding a
# custom ContentType <Override> makes PowerPoint silently drop the ribbon tab.
REL_TYPE = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"

# Default target — override with: python patch_ribbon.py "C:\path\to\OnionPPT.ppam"
DEFAULT_PPTM = os.path.join(
    os.path.expandvars(r"%APPDATA%"),
    "Microsoft",
    "AddIns",
    "OnionPPT.ppam",
)


def patch(pptm_path: str) -> None:
    if not os.path.isfile(pptm_path):
        raise FileNotFoundError(f"Add-in not found: {pptm_path}")
    if not os.path.isfile(XML_SRC):
        raise FileNotFoundError(f"Ribbon XML not found: {XML_SRC}")

    tmp = pptm_path + ".tmp"
    shutil.copy2(pptm_path, tmp)

    with zipfile.ZipFile(tmp, "r") as zin:
        data = {n: zin.read(n) for n in zin.namelist()}

    with open(XML_SRC, "rb") as f:
        data[XML_DEST] = f.read()

    ct = data["[Content_Types].xml"].decode("utf-8")
    ct = re.sub(r'\s*<Override PartName="/customUI/customUI(?:14)?\.xml"[^>]*/>', "", ct)
    if not re.search(r'<Default[^>]*Extension="xml"[^>]*/>', ct):
        ct = ct.replace(
            "</Types>",
            '<Default Extension="xml" ContentType="application/xml"/></Types>',
        )
    data["[Content_Types].xml"] = ct.encode("utf-8")

    rel = data["_rels/.rels"].decode("utf-8")
    rel_entry = (
        '<Relationship Id="rIdUI" '
        f'Type="{REL_TYPE}" '
        f'Target="{XML_DEST}"/>'
    )
    rel_pattern = r'<Relationship [^>]*relationships/ui/extensibility[^>]*/>'
    if re.search(rel_pattern, rel):
        rel = re.sub(rel_pattern, rel_entry, rel, count=1)
    else:
        rel = rel.replace("</Relationships>", f"{rel_entry}</Relationships>")
    data["_rels/.rels"] = rel.encode("utf-8")

    with zipfile.ZipFile(pptm_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, content in data.items():
            zout.writestr(name, content)

    os.remove(tmp)
    print(f"Ribbon updated successfully: {pptm_path}")


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_PPTM
    patch(target)
