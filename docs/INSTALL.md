# Onion Tools — Installation

## For end users

You only need:

- `OnionPPT.ppam`

### Install in PowerPoint

1. Open PowerPoint.
2. Go to **File → Options → Add-ins**.
3. At the bottom, under **Manage**, choose **PowerPoint Add-ins** and click **Go…**.
4. Click **Add New…** (or **Add…**, depending on your version).
5. Select `OnionPPT.ppam`.
6. Make sure the add-in is checked in the list.
7. Close and reopen PowerPoint if the **Onion Tools** tab does not appear.

### Uninstall

1. **File → Options → Add-ins → Go…**
2. Uncheck **OnionPPT** to disable it.
3. Optionally remove it from the list and delete the `.ppam` file.

### Optional: install script (Windows)

From the repo root:

```powershell
.\scripts\install-onionppt.ps1 -AddInPath ".\path\to\OnionPPT.ppam"
```

Uninstall:

```powershell
.\scripts\uninstall-onionppt.ps1
```

---

## For developers (build the add-in)

1. Open your PowerPoint macro-enabled project (`.pptm`) or existing add-in.
2. In the VBA Editor (**Alt+F11**), import:
   - `src/vba/OnionSkinTools.bas`
   - `src/vba/OnionSkinEventBridge.cls`
3. Embed the Ribbon UI from `src/ribbon/customUI.xml` into the package (as `customUI/customUI14.xml` inside the Office Open XML zip). You can use `scripts/patch_ribbon.py` to patch an existing `.ppam` / `.pptm`.
4. Save as **PowerPoint Add-In** (`*.ppam`) named `OnionPPT.ppam`.

### Patch Ribbon into an existing file

Edit `PPTM` in `scripts/patch_ribbon.py` to point at your `.ppam` or `.pptm`, then run from the repo root:

```bash
python scripts/patch_ribbon.py
```

The script reads `src/ribbon/customUI.xml` automatically.
