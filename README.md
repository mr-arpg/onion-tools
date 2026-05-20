# 🧅 Onion Tools

> *Peel back the layers. Animate with confidence.*

**Onion Tools** is a PowerPoint add-in for frame-by-frame animation — onion-skin guides, smart backgrounds, and one-click GIF export, right inside your deck.

---

## ✨ What it does

| Feature | What you get |
|--------|----------------|
| 🫥 **Onion-skin guide** | Ghost of the previous slide, semi-transparent, so you can line up the next frame |
| 🎬 **Animation modes** | **Classic** (guide only) or **Duplicate** (full slide copy + guide) on every new slide |
| 🖼️ **Background shapes** | Lock scene elements behind your drawings; they carry forward when you press Enter |
| 📤 **Export** | GIF or MP4, with or without guides — straight to your output folder |

---

## 🚀 Quick start

### For animators (just use it)

1. Get `OnionPPT.ppam` (from a release or whoever built it for you).
2. In PowerPoint: **File → Options → Add-ins → Go… → Add…** and pick the `.ppam`.
3. Look for the **Onion Tools** tab. You're in.

More detail → [docs/INSTALL.md](docs/INSTALL.md)

### For builders (from this repo)

1. Import the VBA modules from `src/vba/` into your `.pptm` / add-in project.
2. Embed `src/ribbon/customUI.xml` as the Ribbon UI (see [docs/INSTALL.md](docs/INSTALL.md)).
3. Save as **PowerPoint Add-In** (`OnionPPT.ppam`).

Optional helper scripts live in `scripts/`.

---

## 📁 Project layout

```
onion-tools/
├── README.md                 ← you are here
├── docs/
│   └── INSTALL.md            ← install & build notes
├── src/
│   ├── vba/                  ← VBA source
│   │   ├── OnionSkinTools.bas
│   │   └── OnionSkinEventBridge.cls
│   └── ribbon/
│       └── customUI.xml      ← Ribbon tab definition
└── scripts/
    ├── install-onionppt.ps1
    ├── uninstall-onionppt.ps1
    └── patch_ribbon.py       ← patch customUI into an existing .ppam/.pptm
```

---

## 🧅 Tips

- Turn on an **animation mode** (Classic or Duplicate) before using **Background Shapes**.
- **Set Background** keeps your selected figures editable — it adds locked copies behind them.
- Use **Export → Set output folder…** so GIFs land somewhere sensible.

---

## 🛠️ Requirements

- Microsoft PowerPoint (Windows)
- Macros enabled for the add-in

---

## 📜 License

Use it, share it, animate cool stuff. *(Add a license file here if you want something formal.)*

---

<p align="center">
  <sub>Made with layers 🧅 — because good animation is never just one slide deep.</sub>
</p>
