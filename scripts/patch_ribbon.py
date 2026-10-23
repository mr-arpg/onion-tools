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
# customUI14 (2009 schema) must use the 2007 relationship. The 2006
# relationship plus an ActiveX content type makes PowerPoint ignore the tab.
REL_TYPE = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
CONTENT_TYPE = "application/vnd.ms-office.customui+xml"

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

    ct = data["[Content_Types].xml"].decode()
    ct_entry = (
        '<Override PartName="/customUI/customUI14.xml" '
        f'ContentType="{CONTENT_TYPE}"/>'
    )
    ct_pattern = r'<Override PartName="/customUI/customUI14.xml"[^>]*/>'
    if re.search(ct_pattern, ct):
        ct = re.sub(ct_pattern, ct_entry, ct, count=1)
    else:
        ct = ct.replace("</Types>", f"  {ct_entry}\n</Types>")
    data["[Content_Types].xml"] = ct.encode()

    rel = data["_rels/.rels"].decode()
    rel_entry = (
        '<Relationship Id="rIdUI" '
        f'Type="{REL_TYPE}" '
        f'Target="{XML_DEST}"/>'
    )
    rel_pattern = r'<Relationship [^>]*relationships/ui/extensibility[^>]*/>'
    if re.search(rel_pattern, rel):
        rel = re.sub(rel_pattern, rel_entry, rel, count=1)
    else:
        rel = rel.replace("</Relationships>", f"  {rel_entry}\n</Relationships>")
    data["_rels/.rels"] = rel.encode()

    with zipfile.ZipFile(pptm_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, content in data.items():
            zout.writestr(name, content)

    os.remove(tmp)
    print(f"Ribbon updated successfully: {pptm_path}")


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_PPTM
    patch(target)
