# Salt and Sanctuary — Translation Toolkit

A set of Python tools for extracting, editing, and reimporting all localizable text from *Salt and Sanctuary* game files into Excel (.xlsx), and rebuilding the patched binary files for use in mods.

> **Tested on game version 1.0.2.2.** Compatibility with older versions is unknown.

---

## Supported File Formats

| Game File | Location | Contents |
|---|---|---|
| `dialog.zdx` | `dialog/data/classic/` and `dialog/data/enhanced/` | NPC dialog lines, localized names, lore descriptions |
| `loot.zlx` | `loot/data/classic/` and `loot/data/enhanced/` | Item display names and flavor text |
| `monsters.zox` | `monsters/data/classic/` and `monsters/data/enhanced/` | Bestiary creature names and lore text |
| `skilltree.zsx` | `skilltree/data/classic/` and `skilltree/data/enhanced/` | Skill names, descriptions, and effect text |
| `strings.ztx` | `dialog/data/` | UI strings shown in menus and HUD (shared, not version-specific) |

> **Classic vs Enhanced:** Classic preserves the original 2016 launch experience. Enhanced (introduced in Patch 10) is an official community-made rebalance — it buffs mediocre gear, removes stat requirements on elemental weapons, improves heavy armor/equip loads, and balances boss resistances. The two modes store their files separately, so they need to be patched independently.

All formats support 13 languages: **EN · FR · IT · DE · ES · PT · JA · ZH · KO · RU · PL · ZH(TW) · TR**

---

## Requirements

- Python 3.10+
- [openpyxl](https://openpyxl.readthedocs.io/)

Install dependencies:

```bash
pip install -r requirements.txt
```

`requirements.txt`:

```
openpyxl>=3.1.0
```

---

## Usage

Each format follows the same three-step workflow. The scripts can be run from anywhere — just make sure the paths to the input files and output files are correct.

> **Important:** The output file **must use the exact same filename** as the original (e.g. `dialog.zdx`, not `dialog_patched.zdx`) so the game engine can find it.

### Step 1 — Export to Excel

```bash
python zdx_export.py "PATH/dialog/data/classic/dialog.zdx" dialog_classic.xlsx
python zdx_export.py "PATH/dialog/data/enhanced/dialog.zdx" dialog_enhanced.xlsx

python zlx_export.py "PATH/loot/data/classic/loot.zlx" loot_classic.xlsx
python zlx_export.py "PATH/loot/data/enhanced/loot.zlx" loot_enhanced.xlsx

python zox_export.py "PATH/monsters/data/classic/monsters.zox" monsters_classic.xlsx
python zox_export.py "PATH/monsters/data/enhanced/monsters.zox" monsters_enhanced.xlsx

python zsx_export.py "PATH/skilltree/data/classic/skilltree.zsx" skilltree_classic.xlsx
python zsx_export.py "PATH/skilltree/data/enhanced/skilltree.zsx" skilltree_enhanced.xlsx

python ztx_export.py "PATH/dialog/data/strings.ztx" strings.xlsx
```

### Step 2 — Edit the Excel file

Open the generated `.xlsx` in Excel or LibreOffice Calc and edit any language column. Each file includes an **Info** sheet explaining the structure.

**Rules:**
- Do **not** change the `Key`, `NPC`, `Action`, `Index`, or `_RowID` columns — these are internal identifiers used to match rows back to their binary locations.
- Do **not** reorder or delete rows.
- Leave a cell blank or unchanged to keep the original text.

### Step 3 — Import back to binary

The output filename must match the original exactly.

```bash
python zdx_import.py "PATH/dialog/data/classic/dialog.zdx" dialog_classic.xlsx "PATH/dialog/data/classic/dialog.zdx"
python zdx_import.py "PATH/dialog/data/enhanced/dialog.zdx" dialog_enhanced.xlsx "PATH/dialog/data/enhanced/dialog.zdx"

python zlx_import.py "PATH/loot/data/classic/loot.zlx" loot_classic.xlsx "PATH/loot/data/classic/loot.zlx"
python zlx_import.py "PATH/loot/data/enhanced/loot.zlx" loot_enhanced.xlsx "PATH/loot/data/enhanced/loot.zlx"

python zox_import.py "PATH/monsters/data/classic/monsters.zox" monsters_classic.xlsx "PATH/monsters/data/classic/monsters.zox"
python zox_import.py "PATH/monsters/data/enhanced/monsters.zox" monsters_enhanced.xlsx "PATH/monsters/data/enhanced/monsters.zox"

python zsx_import.py "PATH/skilltree/data/classic/skilltree.zsx" skilltree_classic.xlsx "PATH/skilltree/data/classic/skilltree.zsx"
python zsx_import.py "PATH/skilltree/data/enhanced/skilltree.zsx" skilltree_enhanced.xlsx "PATH/skilltree/data/enhanced/skilltree.zsx"

python ztx_import.py "PATH/dialog/data/strings.ztx" strings.xlsx "PATH/dialog/data/strings.ztx"
```

> **Tip:** It's recommended to back up the original files before overwriting them.

---

## File Structure

```
(anywhere you like)/
├── zdx_parser.py       # Shared binary parser for dialog.zdx
├── zdx_export.py       # Export dialog.zdx → Excel
├── zdx_import.py       # Rebuild dialog.zdx from Excel
│
├── zlx_parser.py       # Shared binary parser for loot.zlx
├── zlx_export.py       # Export loot.zlx → Excel
├── zlx_import.py       # Rebuild loot.zlx from Excel
│
├── zox_parser.py       # Shared binary parser for monsters.zox
├── zox_export.py       # Export monsters.zox → Excel
├── zox_import.py       # Rebuild monsters.zox from Excel
│
├── zsx_parser.py       # Shared binary parser for skilltree.zsx
├── zsx_export.py       # Export skilltree.zsx → Excel
├── zsx_import.py       # Rebuild skilltree.zsx from Excel
│
├── ztx_parser.py       # Shared binary parser for strings.ztx
├── ztx_export.py       # Export strings.ztx → Excel
├── ztx_import.py       # Rebuild strings.ztx from Excel
│
└── requirements.txt

PATH/                          ← game data root
├── dialog/
│   └── data/
│       ├── strings.ztx        ← UI / HUD strings (shared)
│       ├── classic/
│       │   └── dialog.zdx
│       └── enhanced/
│           └── dialog.zdx
├── loot/
│   └── data/
│       ├── classic/
│       │   └── loot.zlx
│       └── enhanced/
│           └── loot.zlx
├── monsters/
│   └── data/
│       ├── classic/
│       │   └── monsters.zox
│       └── enhanced/
│           └── monsters.zox
└── skilltree/
    └── data/
        ├── classic/
        │   └── skilltree.zsx
        └── enhanced/
            └── skilltree.zsx
```

---

## Notes on Specific Formats

### `dialog.zdx`
The Dialog sheet uses `NPC` + `Action` + `_RowID` as a composite key for each line. `_RowID` disambiguates duplicate `NPC`/`Action` pairs that appear at more than one physical location in the file. Do not edit or delete `_RowID` values.

### `monsters.zox`
Records are split across four sheets: `Monster_Names`, `Monster_Descriptions`, `NPC_Names`, `NPC_Descriptions`. A matching `Key` + `_RowID` pair across a Names sheet and its Descriptions sheet represents one creature record.

### `skilltree.zsx`
Skill keys are **not unique** — the same key can appear at many nodes in the skill tree with different lore text per node. Rows are matched by their `Index` (position in the file). Do not sort or filter rows in Excel.

### `strings.ztx`
Duplicate keys (same string appearing in multiple sections) are highlighted in yellow. Each occurrence is a separate physical entry and must be edited independently.

---

## How It Works

Each parser reverse-engineers the game's binary format. Because the game files do not declare explicit blob-length fields between records, the parsers use a heuristic: they scan forward for a byte sequence that looks like a valid length-prefixed key followed by a valid UTF-8 string, and use that as the record boundary. This approach has been verified against all records in the original game files from version 1.0.2.2.

Imports are non-destructive: only bytes that differ from the original are patched. All binary stat/animation/blob data is copied through verbatim and never modified.

---

## License

This project is not affiliated with Ska Studios. *Salt and Sanctuary* is © Ska Studios. Use these tools at your own risk and only for personal modding purposes.
