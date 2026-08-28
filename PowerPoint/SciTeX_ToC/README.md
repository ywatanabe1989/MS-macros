# SciTeX ToC for PowerPoint

A table of contents that keeps itself correct.

Every index page in a deck is rebuilt from the slides themselves — numbering,
titles, links, and how the entries split across columns. Add, delete or reorder
slides, run it again, and the index follows.

- **Module** `SciTeX_ToC`
- **Macro** `RefreshToC` (Alt+F8)
- **Template** [`dist/SciTeX_ToC_v0.5.0.pptm`](dist/) — usage included in the file

## Quick start

1. Open `dist/SciTeX_ToC_v0.5.0.pptm`. It explains itself in ten slides.
2. Alt+F8 → `RefreshToC` → Run.
3. Copy your own slides in, or copy the module into your own deck.

A backup is written beside the file first, named
`<name>.before-toc-<timestamp>.pptm`, because PowerPoint's undo does not
survive a macro.

## What a deck must have

| Name | Where | What it is |
|---|---|---|
| `SCITEX_TITLE` | every slide | the slide's title |
| `SCITEX_TOC_BODY_*` | each index page | one index column |
| `SCITEX_STATUS` | anywhere, once | where the macro reports what it did |
| `SCITEX_CFG_*` | the config page | the settings |

Slide tags mark the roles: `SCITEX_COVER` on the cover, `SCITEX_TOC` on each
index page, `SCITEX_CONFIG` on the settings page. Rename a shape in the
Selection Pane (Home → Select → Selection Pane).

The macro checks all of this before it changes anything, and reports every
problem at once rather than stopping at the first.

## Columns

Any shape whose name starts with `SCITEX_TOC_BODY_` is a column, and they are
ordered by their left edge. Two columns or three — add another box and the
macro uses it. Nothing counts to two.

Entries fill the first column from the top and carry on into the next. The
split is measured, not written down, so it moves when the slides do. More
columns lets the type stay larger:

| 54 entries | 2 columns | 3 columns |
|---|---|---|
| split | 28 / 26 | 18 / 18 / 18 |
| type | 12pt (the floor) | 15pt |

## Configuration

The last page. Each setting shows what it accepts beside the box.

| Setting | Accepts |
|---|---|
| `SCITEX_CFG_FONT_LATIN` | any installed font |
| `SCITEX_CFG_FONT_CJK` | Yu Gothic / Meiryo / MS Gothic |
| `SCITEX_CFG_FONT_MIN` | points; the floor when it will not fit |
| `SCITEX_CFG_FONT_MAX` | points; the size it starts from |
| `SCITEX_CFG_HIDE_HIDDEN` | Yes / No |
| `SCITEX_CFG_VERSION` | written by the macro; do not edit |

`FONT_MIN` is the one worth understanding. The macro starts at `FONT_MAX` and
steps down until everything fits. If it reaches `FONT_MIN` and the index still
does not fit, it stops there and names the offending slides on the status line
rather than letting entries fall off the page. A `FONT_MIN` set too high is
therefore not an error — it is a floor you chose — but it is the usual reason
an index overflows.

## When it will not fit

Small type is the last resort, not the first. If the index needs a size you do
not want, the deck is carrying more than an index page can show: cut entries,
or add a column.

## Setup — AccessVBOM (only for scripted runs)

Importing the module by hand and pressing the button needs **nothing**. This is
required only when a script writes VBA, as `tools/apply-toc.ps1` does.

File → Options → Trust Center → Trust Center Settings → Macro Settings →
tick *Trust access to the VBA project object model* → restart PowerPoint.

The restart is part of the step: a process that has already started keeps the
setting it read at startup.

Registry equivalent:

```
HKCU\Software\Microsoft\Office\16.0\PowerPoint\Security
  AccessVBOM (DWORD) = 1
```

**What this allows.** Macros can rewrite macros. If you enable it permanently,
pair that with not opening macro-enabled files of unknown origin.
`tools/apply-toc.ps1` instead turns it on, runs, and puts it back — from a
`finally` block, so a crash still closes it.

## Tools

| Script | What it does |
|---|---|
| `tools/apply-toc.ps1` | run the macro on a deck unattended |
| `tools/build-template.ps1` | build the distributable template |
| `tools/check-structure.py` | read an exported `.bas` for unbalanced blocks, undefined calls, and members PowerPoint does not have |
| `tools/build-and-test.ps1` | build a sandbox deck and assert against it |
| `tools/restore-access-vbom.ps1` | put AccessVBOM back |

### Running these from WSL

Call `powershell.exe` through `cmd.exe`. Invoking it directly across the WSL
interop socket fails intermittently, and when it does it prints only

```
WSL ERROR: UtilAcceptVsock:273: accept4 failed 110
```

— so a run that never started looks exactly like one that produced no output.

```sh
cd /mnt/c/Users/<user>          # cmd cannot start in a \\wsl.localhost path
cmd.exe /c "powershell -NoProfile -ExecutionPolicy Bypass -File C:\...\apply-toc.ps1 ..."
```

## Compatibility

`RunSciTeXNavigation` still works; it forwards to `RefreshToC` so an existing
button keeps working. It is removed in the next version.

Shape names and slide tags keep the `SCITEX_` prefix and are unchanged by the
rename — they live inside existing decks, and renaming them would break every
file already using this.
