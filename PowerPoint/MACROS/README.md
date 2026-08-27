# MACROS.pptm

`../MACROS.pptm` is the working macro container kept open alongside whatever
presentation is being edited. Macros run from it act on the *active*
presentation, so the deck being edited does not itself have to be a `.pptm`.

That matters in practice: a `.pptx` cannot store macros at all, and a `.pptm`
opened from a network location (including a Windows drive mapped to WSL via
`\\wsl.localhost`) has its macros disabled by Office. Keeping the macros in a
local `MACROS.pptm` sidesteps both.

## Modules

The `.bas` files here are the VBA source extracted from `MACROS.pptm`, so the
macro code is readable and diffable in git. The `.pptm` remains the artifact you
actually open; these are its contents in text form.

| module | contents |
| --- | --- |
| `Module1.bas` | `AddElements`, `RemElements`, `GetRGBColor`, `SetDefaultColors`, `MultipleCropping`, `CropWhiteSpace`, `FileExists` |
| `Module2.bas` | `InsertImagesOnePerSlide` |
| `Module3.bas` | `InsertImagesTiled` |
| `Module4.bas` | `ImportImagesFromFolder` |
| `Module5.bas` | `RunSciTeXNavigation` and its helpers — the navigation toolkit, `NAVIGATION_VERSION = 0.1.2` |

## Relationship to the other files in this directory

`Module5.bas` and [`../SciTeXNavigation/src/SciTeXNavigation.bas`](../SciTeXNavigation/src/SciTeXNavigation.bas)
are the same v0.1.2 navigation code. They differ only in the `Attribute VB_Name`
line and in `Err.number` / `Err.Number` capitalisation, which the VBA editor
rewrites on its own; VBA identifiers are case-insensitive, so the two behave
identically.

`Module1.bas` is **ahead of** [`../macros.vba`](../macros.vba): it holds
everything that file has plus `AddElements`, `RemElements`, `CropWhiteSpace` and
`FileExists`. `macros.vba` has not been refreshed from the live container.

## Note on `CropWhiteSpace`

`CropWhiteSpace` shells out to WSL and hardcodes two paths:

    /home/ywatanabe/.dotfiles/.bin/crop_whitespace.py
    /home/ywatanabe/.env/bin/python3

It will not work on a machine without those. Nothing else in these modules
depends on the local environment.

## Regenerating

    python3 -m oletools.olevba --code MACROS.pptm

Editing a `.bas` file here does **not** change `MACROS.pptm` — the `.pptm` is the
source of truth, and these files are exported from it.
