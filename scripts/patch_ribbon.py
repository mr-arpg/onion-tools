"""Patch customUI.xml into a PowerPoint .ppam or .pptm (Office Open XML zip)."""
import zipfile
import shutil
import os
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.dirname(SCRIPT_DIR)
XML_SRC = os.path.join(REPO_ROOT, "src", "ribbon", "customUI.xml")
XML_DEST = "customUI/customUI14.xml"

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
    entry = (
        '<Override PartName="/customUI/customUI14.xml" '
        'ContentType="application/vnd.ms-office.activeX+xml"/>'
    )
    if "customUI14.xml" not in ct:
        ct = ct.replace("</Types>", f"  {entry}\n</Types>")
        data["[Content_Types].xml"] = ct.encode()

    rel = data["_rels/.rels"].decode()
    rel_entry = (
        '<Relationship Id="rIdUI" '
        'Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" '
        'Target="/customUI/customUI14.xml"/>'
    )
    if "extensibility" not in rel:
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
