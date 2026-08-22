# SciTeX Navigation for PowerPoint

`SciTeXNavigation` is a tested VBA module for repeatable slide numbering and full-presentation tables of contents.

Version: **0.1.0**

## What it does

- Keeps every TOC slide synchronized with the complete presentation outline.
- Highlights the current section and dims every other section.
- Indents child entries such as `1a`, `1b`, and `2a` below their major section.
- Adds clickable internal links to TOC entries.
- Re-runs safely without duplicating number prefixes.
- Optionally omits hidden slides from the TOC.
- Applies configurable Latin/CJK fonts and minimum/maximum sizes only to managed navigation shapes.
- Exposes one public macro, `RunSciTeXNavigation`; implementation helpers stay private.

## Try the tested sandbox

Open [`dist/SciTeXNavigationSandbox_v0.1.0.pptm`](./dist/SciTeXNavigationSandbox_v0.1.0.pptm), enable macros, then run `RunSciTeXNavigation` from `Alt+F8`. The last slide is the hidden English configuration page.

## Configuration

The configuration slide carries the `SCITEX_CONFIG=1` slide tag and uses these named text boxes:

| Shape | Value |
| --- | --- |
| `SCITEX_CFG_FONT_LATIN` | Latin font name |
| `SCITEX_CFG_FONT_CJK` | CJK font name |
| `SCITEX_CFG_FONT_MIN` | Minimum navigation font size in points |
| `SCITEX_CFG_FONT_MAX` | Maximum navigation font size in points |
| `SCITEX_CFG_HIDE_HIDDEN` | `Yes` or `No` |
| `SCITEX_CFG_VERSION` | Displayed release version |

The configuration slide is hidden and is never included in the TOC.

## Navigation model

The generic mode uses these slide tags and shape names:

- `SCITEX_COVER=1`: exclude the cover.
- `SCITEX_TOC=1`: make a slide a major section/TOC page.
- `SCITEX_SECTION_TITLE`: stable unnumbered section title.
- `SCITEX_TITLE`: title shape on each navigated slide.
- `SCITEX_TOC_BODY`: full TOC body in a one-column layout.

For decks whose numbering does not follow slide order, explicit mode adds:

- `SCITEX_NAV_CODE`: fixed code such as `3`, `3a`, or `4f`.
- `SCITEX_CURRENT_SECTION`: major section emphasized on a TOC slide.
- `SCITEX_TOC_SPLIT_AFTER`: last major section placed in the left column.
- `SCITEX_TOC_BODY_LEFT` and `SCITEX_TOC_BODY_RIGHT`: two-column TOC shapes.

## Source and validation

- [`src/SciTeXNavigation.bas`](./src/SciTeXNavigation.bas): canonical VBA source.
- [`tools/build-and-test.ps1`](./tools/build-and-test.ps1): creates and validates the isolated sandbox.
- [`tools/reopen-test.ps1`](./tools/reopen-test.ps1): reopens the generated file in a fresh PowerPoint process and validates repeated runs, links, indentation, dimming, configuration, and hidden-slide behavior.
- [`tools/prepare-aichi-v10.ps1`](./tools/prepare-aichi-v10.ps1): deterministic v10 migration tool; it keeps the existing slide master, layout, theme, and slide order.
- [`tools/validate-aichi-v10.ps1`](./tools/validate-aichi-v10.ps1): fresh-reopen validation for the migrated v10 deck.

The PowerShell tools require desktop PowerPoint on Windows. When VBA project access is needed, the scripts temporarily enable `AccessVBOM` and restore its previous registry state during cleanup.

## Release policy

The repository contains the reusable module and generic sandbox. Business-contest decks remain outside this public repository.
