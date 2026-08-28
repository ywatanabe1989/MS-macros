# SciTeX Navigation for PowerPoint

`SciTeXNavigation` is a tested VBA module for repeatable slide numbering and full-presentation tables of contents.

Version: **0.1.2**

## What it does

- Keeps every TOC slide synchronized with the complete presentation outline.
- Highlights the current section and dims every other section.
- Treats each TOC slide as a top-level section and every following ordinary slide as its child until the next TOC.
- Underlines only top-level section entries; child links remain clickable without underlines.
- Adds clickable internal links to TOC entries.
- Re-runs safely without duplicating number prefixes.
- Recomputes the active section when a TOC slide is copied or moved.
- Optionally omits hidden slides from the TOC.
- Applies configurable Latin/CJK fonts and minimum/maximum sizes only to managed navigation shapes.
- Exposes one public macro, `RunSciTeXNavigation`; implementation helpers stay private.

## セットアップ — AccessVBOM（スクリプトから取り込む場合のみ）

手でモジュールを取り込んで実行するだけなら **不要** です。以下が要るのは
`tools/apply-navigation.ps1` のようにスクリプトが VBA を書き込む場合だけ。

**GUI で恒久的に有効化する**（運用者の手順、2026-08-28）:

```
ファイル → オプション → トラスト センター → トラスト センターの設定
  → マクロの設定
  → 「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」にチェック
  → PowerPoint を再起動
```

再起動まで含めて一手順です。チェックしただけで再起動していないプロセスは、
起動時に読んだ古い設定のまま動きます。

**レジストリ**（同じ設定の実体）:

```
HKCU\Software\Microsoft\Office\16.0\PowerPoint\Security
  AccessVBOM (DWORD) = 1
```

**何を許可しているのか。** これはマクロがマクロを書き換えられる状態です。
恒久的に有効にするなら、出所の分からないマクロ入りファイルを開かない運用と
セットで考えてください。一時的に開けて元に戻す形が要るなら
`tools/apply-navigation.ps1` がそれをやります（元の値を記録し、`finally` から
`tools/restore-access-vbom.ps1` を呼んで戻すので、途中で落ちても閉じます）。

## Try the tested sandbox

Open [`dist/SciTeXNavigationSandbox_v0.1.2.pptm`](./dist/SciTeXNavigationSandbox_v0.1.2.pptm), enable macros, then run `RunSciTeXNavigation` from `Alt+F8`. The last slide is the hidden English configuration page.

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

The TOC-driven model uses these slide tags and shape names:

- `SCITEX_COVER=1`: exclude the cover.
- `SCITEX_TOC=1`: make a slide a major section/TOC page.
- `SCITEX_SECTION_TITLE`: section name derived from the TOC heading. A heading such as `Contents: Company Overview` keeps `Company Overview` as the authoritative name.
- `SCITEX_TITLE`: title shape on each navigated slide.
- `SCITEX_TOC_BODY`: full TOC body in a one-column layout.
- `SCITEX_NAV_CODE`: automatically assigned code such as `3`, `3a`, or `4f`.
- `SCITEX_CURRENT_SECTION`: cached major section; the macro refreshes it from TOC order.
- `SCITEX_TOC_SPLIT_AFTER`: automatically balanced last section in the left column.
- `SCITEX_TOC_BODY_LEFT` and `SCITEX_TOC_BODY_RIGHT`: two-column TOC shapes.

## Source and validation

- [`src/SciTeXNavigation.bas`](./src/SciTeXNavigation.bas): canonical VBA source.
- [`tools/build-and-test.ps1`](./tools/build-and-test.ps1): creates and validates the isolated sandbox.
- [`tools/reopen-test.ps1`](./tools/reopen-test.ps1): reopens the generated file in a fresh PowerPoint process and validates repeated runs, links, indentation, dimming, configuration, and hidden-slide behavior.
- [`tools/upgrade-pptm.ps1`](./tools/upgrade-pptm.ps1): replaces only the SciTeX Navigation module in a copied PPTM, preserving its slides, master, layout, and theme.
- [`tools/validate-aichi-v10.ps1`](./tools/validate-aichi-v10.ps1): fresh-reopen validation for the TOC-driven AICHI v10 deck.

The PowerShell tools require desktop PowerPoint on Windows. When VBA project access is needed, the scripts temporarily enable `AccessVBOM` and restore its previous registry state during cleanup.

## Release policy

The repository contains the reusable module and generic sandbox. Business-contest decks remain outside this public repository.

## Preconditions — what a deck must already have

The macro edits shapes it finds by name. It does not create them. Give it a
deck that does not meet these and it raises rather than guessing:

| Requirement | Why | Error if missing |
|---|---|---|
| Every non-config slide has a shape named `SCITEX_TITLE` | it is read and rewritten as the slide's title | `2102` / `2103` |
| Slides to be indexed carry the tag `SCITEX_NAV_CODE` | untagged slides are skipped, silently and on purpose | — |
| Index slides carry `SCITEX_TOC`, and a body named `SCITEX_TOC_BODY` (or `_LEFT` + `_RIGHT` for two columns) | the index text is written into it | — |
| `SCITEX_TITLE` has usable width to the left of the logo | the title is shrunk to fit, and cannot fit in nothing | `2114` |
| `SCITEX_TOC_BODY*` has usable height | same, vertically | `2115` |
| The config slide is tagged `SCITEX_CONFIG` | optional; defaults apply without it | — |

Shapes named anything else are never touched. The managed set is exactly
`SCITEX_TITLE`, `SCITEX_TOC_BODY*`, `SCITEX_STATUS`, `SCITEX_RUN_BUTTON`.

## v0.2.0 — the three defects reported 2026-08-27

Measured on `AICHI-NEXT-UNICORN_SciTeX_v11_navigation_v0.1.1.pptm` with
`scitex-kk/08_slides/scripts/check_deck.py` in the `business` repo.

**1. `3a.` / `3b.` prefixes did not line up.** The entry was built as
`code & ". " & title` in one proportional run. `1a.` and `3i.` are not the same
width, so no amount of padding aligns the titles. Now the separator is a TAB
and the body carries one left tab stop at `SCITEX_CFG_TOC_PREFIX_TAB`
(default 34pt), with a hanging indent so a wrapped title lines up under itself.

**2. Titles did not match the slide master.** `ApplyTypography` wrote
`Font.Name = mFontLatin` onto `SCITEX_TITLE` on every run — the deck's config
said `Arial`, and that overwrote the theme font the master asks for. It also
clamped the title into the *body* size range, which is why every title came out
at the body maximum instead of the master's size. Titles now bind to the theme
(`+mj-lt` / `+mj-ea`), and take their size from `SCITEX_CFG_TITLE_SIZE` when it
is set (`0`, the default, leaves the authored size alone).

Note the shape is a plain text box, not the title placeholder, so it inherits
*nothing* from the master. The size has to be stated; there is no "just let it
inherit" available.

**3. The index ran off the slide.** The bodies are authored with `spAutoFit`,
which does not clip — it grows the shape. The overflow was never text leaving
its box, it was the box leaving the slide, which is why every stored rectangle
still looked correct. At 18pt with 15 entries there was about 154pt of headroom,
seven lines: a few wrapped titles on a machine with different font metrics are
enough. `FitTocBody` now turns autosize off and shrinks to fit, so the box
cannot leave the slide rather than being unlikely to.

`FitTocTitles` also now runs on **every** slide, not only index slides. Every
slide's `SCITEX_TITLE` had the same grow-sideways settings; the index pages
were just where it was noticed first.

### How to check the fix without opening the deck

```bash
# in the business repo, after running the macro and saving
/uvwork/venv-pptx/bin/python scitex-kk/08_slides/scripts/check_deck.py <deck> \
  --baseline <the same deck before the macro ran>
```

Accepted when `prefix-font` reports 0 and `text-fit` reports no grow-to-fit box
past a slide edge. Exit `0` clean, `10` findings, `20` unreadable.

### Not verified here

These changes are **not executed**. There is no PowerPoint on the machine that
wrote them, so what has been checked is the structure of the module
(Sub/Function, For/Next, With/End With, Do/Loop all balanced; every module
variable declared; every `GoTo` target present) and the diagnosis the changes
follow from. The behaviour itself is unproven until someone runs it. Run
`tools/build-and-test.ps1` on Windows before shipping a deck with it.

### 縮めても入らないとき

`SCITEX_CFG_FONT_MIN`（既定18pt）まで縮めても収まらない場合、**それ以上は
小さくしません**。代わりに `SCITEX_STATUS` に該当スライド番号を出します。

```
Navigation v0.2.0 updated - 4 sections. Too much content at 18pt on slide(s): 8,19
```

operator の判断（2026-08-27）:「あんまり文字が小さくなるようなら、そもそも
スライドにそんな入れるなって話」。読めない字で収めるのは解決ではないので、
その時点で内容を減らす合図として出します。黙って溢れさせないのは、それが
このバージョンで直した不具合と見分けがつかなくなるためです。

### 走らせる前に必ずバックアップが取られます

マクロが動くと PowerPoint の**元に戻す（Ctrl+Z）は効きません**。スライドの
採番と目次2列の書き換えが終わった後では、元の内容を持っているのは古い
ファイルだけです。

そこで `RunSciTeXNavigation` は、何かを変更する前に必ず別名でコピーを保存
します。

```
<元のファイル名>.before-navigation-20260828-031500.pptm
```

保存先は元のファイルと同じフォルダで、ファイル名はステータス欄にも出ます。
おかしくなったらこのファイルを開いてください。

**一度も保存していないファイルでは動きません**（エラー `2120`）。コピーの
置き場所が無く、戻る先も無いためです。

## v0.3.0 — レイアウトの数字を全部やめた

operator 2026-08-28:「レイアウトの調節は数えれば最適化できる問題ですよね」
「そこを手でやりたくないのでマクロでお願いしています」
「入らない場合文字が小さくなるのは仕方ないです」

そのとおりで、これまで書き込まれていた数字は全部スライド上にあるものだった。
測れるものを定数で持っていたので、外した。

| 消した数字 | 何だったか | 何に置き換えたか |
|---|---|---|
| `34pt` | 目次のタブ位置 | 一番幅の広い接頭辞を測り、その右に1字分 |
| `72pt` | ロゴのぶんの余白 | タイトルの右で縦に重なる図形の左端。無ければスライド端 |
| `24pt` / `28pt` | タイトルのサイズ | 全タイトルが収まる最大の1サイズを計算 |
| 左右バラバラ | 目次2カラムのサイズ | 小さいほうに揃える |
| `SCITEX_CFG_TITLE_SIZE` | タイトルサイズの設定 | 計算するので削除 |
| `SCITEX_CFG_TOC_PREFIX_TAB` | タブ位置の設定 | 計算するので削除 |

### 設定として残したもの、その理由

`SCITEX_CFG_FONT_MIN`（既定18pt）と `SCITEX_CFG_FONT_MAX`（既定32pt）は
残した。**これは測れない。** 「何ptより小さいと読めないか」はスライドの
性質ではなく読む人についての判断で、数えても出てこない。

下限まで縮めても入らないときは、それ以上小さくせず `SCITEX_STATUS` に
スライド番号を出す。operator 判断:「あんまり文字が小さくなるようなら、
そもそもスライドにそんな入れるなって話」。

### 副作用として直ったこと

`ApplyTypography` と `FitTocTitles` が両方タイトルのサイズを決めていた。
2つのルーチンが同じことを決めて、後から動くほうが黙って勝つ状態だった。
サイズの決定は `FitTocTitles` に一本化した。
