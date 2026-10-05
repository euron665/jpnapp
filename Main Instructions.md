# KOKU — Main Instructions

The single source of truth for the KOKU Japanese-learning app. Everything another session
needs: architecture, the update workflow, every data pipeline, every review mode, the full
colour system, spacing rules, and the traps that have already cost real damage.

---

## ⚠ MAINTENANCE RULE — THIS FILE MUST BE KEPT CURRENT

**Any change to the essence or framework of the app must be reflected in this document, in
the same session that makes the change.** This is not optional and does not wait for a
later tidy-up pass.

That means, at minimum:

- a new or removed **review or study mode** → Part 7
- a change to the **data pipeline**, spreadsheet format, or how a constant is derived → Parts 4–6
- a new or changed **colour, tag, or part of speech** → Parts 5 and 9.2
- a change to **navigation, screens, or layout/spacing** → Part 8
- a change to the **grammar module** structure, conventions, or furigana system → Part 9
- a change to the **update or deploy workflow** → Part 3
- any **new trap, bug class, or hard-won lesson** → Part 0

If a change makes a statement here wrong, fix the statement — do not leave it stale and do
not append a contradicting note elsewhere. A doc that lies is worse than no doc: every
future session trusts this file over the code.

Counts quoted throughout (cards, examples, groups, colour usages) are snapshots. When a
change moves one, update it.

---

# PART 0 — HARD RULES (read before touching anything)

### 0.1 Never bulk-write `index.html` through the shell

It is 3.7 MB, and shell/Python writes have **silently truncated and corrupted it before**.
Reading it in Python for *counting and inspection* is fine and done constantly — it is
**writing** that is forbidden.

Use the harness tools with targeted, unique string matches:
- `Grep` to locate, `Read` with `offset`/`limit` to read a slice, `Edit` to replace a unique string.
- For a large structural rewrite: build the new file in `/tmp`, verify exhaustively, then
  install with `rm` + `cp`, then verify the installed file again.

### 0.2 Never edit with computed offsets

On 2026-07-26 an edit built from `seg.index` / `rindex` plus manual slice arithmetic
**silently deleted 28 examples** and unbalanced the div tree. Exact-string replacement with
an expected count, always.

### 0.3 Never use a non-greedy regex to grab a block

`<div class="description">.*?</div>` is **wrong** — descriptions contain nested
`<div class="recap-rules">`, so the match stops at the first inner `</div>`. This made a
whole prose pass blind to ~5,700 characters across 16 cards. Depth-parse instead:

```python
def bal(raw, s):
    i, dep = s, 0
    while i < len(raw):
        if raw.startswith('<div', i):  dep += 1; i += 4
        elif raw.startswith('</div>', i):
            dep -= 1; i += 6
            if dep == 0: return i
        else: i += 1
```

### 0.4 Catastrophic regex on big files

`[^{}]*\{[^}]*var\(--x\)[^}]*\}` against 3.4 MB backtracks forever and will hang the shell.
Use plain string search, or anchor tightly.

### 0.5 Other traps already paid for

| Trap | What happened |
|---|---|
| **Curly quotes** | `&#x201C;…&#x201D;` and `"…"` are invisible to a regex matching `"`. Always match `["“”]` and `&#x201[CD];`. |
| **Case-folded patterns** | A pattern containing `English` tested against `pre.lower()` never matches. |
| **Series order** | When replacing several spans in one block, splice **right-to-left** or earlier edits invalidate later anchors. |
| **Tag-name overcounting** | `AFFIRMATIVE` also appears in example tags (20 hits in one card). Scope to the extracted substring, not the card. |
| **CSS rule order** | Same-specificity rules: **last one wins.** Adding an override *before* an existing rule does nothing. Always verify source position. |
| **Inline styles** | Markup uses inline `style="padding-top:56px"` in 17 places; CSS can only beat it with `!important`. |
| **Stray closers** | Old grammar files carried stray extra `</div>` tags. `index.html` is balanced at 0 — keep it that way; verify after every structural edit. |

### 0.6 Referring to grammar cards

Always by **coordinate** — `CH.8 · CR.13` — never by `card-NN` id. The id is an
implementation detail; the coordinate is what the user navigates by.

---

# PART 1 — WHAT KOKU IS

A self-contained Japanese-learning PWA (kanji, vocab, katakana, grammar, conjugation),
installed to the iPhone home screen. Everything ships in **one `index.html`** — HTML, CSS, JS
and all data inlined as JS constants. No backend, no runtime data fetches.

**Live:** `https://euron665.github.io/jpnapp/` — GitHub Pages, repo **`euron665/jpnapp`**,
served under `/jpnapp/`. Account owner: Seb (s.dybalski95@gmail.com).

**Deploy = upload two files** to the repo root: `index.html` and `sw.js`.

### Current scale

| Thing | Count |
|---|---|
| Kanji | 889 |
| SKGs (kanji sets) | 104 |
| Vocab words (POS-tagged) | 3,055 |
| Same-meaning groups | 271 |
| Katakana words | 260 |
| Grammar cards | 200 (N5 + N4) |
| Grammar examples | 2,198 |
| Furigana readings | 8,758 |
| index.html | ~3.7 MB |

---

# PART 2 — FILE MAP (`C:\Users\Seb\Desktop\app`)

| File | Role |
|---|---|
| `index.html` | **The live app.** Single-file HTML/CSS/JS + inlined data. Harness-edit only. |
| `sw.js` | Service worker. **Network-first.** `const CACHE='koku-vNN'` bumped every change. |
| `manifest.json`, `icon-192.png`, `icon-512.png` | PWA install assets. |
| `update_koku.py` | The pipeline: spreadsheets → JSON + inlined consts → SW bump. |
| `update_koku.bat` | One-line launcher: `python "%~dp0update_koku.py"`. Run on Windows. |
| `same_meaning_detector.py` | Finds new same-meaning candidate groups. Read-only on spreadsheets. |
| `FINAL VOCAB LIST v.1.8.xlsx` | Vocab source. Reading groups encoded by **cell fill colour**. |
| `FINAL KANJI & SKG LIST v1.2.xlsx` | Kanji + set (SKG) source. |
| `FINAL POS LIST v.1.3.xlsx` | Part-of-speech tags source. |
| `FINAL KATAKANA VOCAB v1.0.xlsx` | Katakana words source. |
| `FINAL SAME MEANING LIST v2.7.xlsx` | Same-meaning groups + nuance notes. |
| `*.json` | Pipeline byproducts (`vocab/kanji/skgs/pos_tags/katakana/same_meaning`). **The app does not read these at runtime** — it uses the inlined consts. They exist for inspection. |
| `Main Instructions.md` | This file. |

---

# PART 3 — THE UPDATE WORKFLOW (what the user actually does)

### 3.1 Normal update — adding kanji or vocab

1. **Edit the spreadsheet** in Excel (`FINAL VOCAB LIST`, `FINAL KANJI & SKG LIST`, etc.).
   Save it. Version bumps are optional — the pipeline globs `FINAL <TYPE>*.xlsx` and takes
   the **newest by modification time**, so `v1.8` edited today beats `v1.9` from last month.
2. **Double-click `update_koku.bat`.** It runs `update_koku.py`, which rewrites the inlined
   constants in `index.html` and bumps `sw.js`.
3. **Upload `index.html` + `sw.js`** to `euron665/jpnapp` via the GitHub web UI
   (Add file → Upload files → commit).
4. On the phone, open the app while online. The network-first SW pulls the new file.

### 3.2 When the `.bat` is NOT needed

If a change only touched `index.html` code (UI, a new review mode, CSS), the data consts are
already current — running the `.bat` only bumps the SW. In that case bump `sw.js` by hand
(or run the `.bat`, harmless) and upload.

**Check whether data is stale** by comparing a baked count against the spreadsheet, e.g.
count `"m":` occurrences inside the `SAME_MEANING` const vs rows in the sheet.

### 3.3 Deployment and the service worker

`sw.js` is **network-first**: online it always fetches the latest and refreshes the cache;
offline it falls back to cache. This was the fix that made updates reliable — **do not revert
it to cache-first.**

```js
const CACHE = 'koku-vNN';          // bumped on every change
const ASSETS = ['/jpnapp/', '/jpnapp/index.html', '/jpnapp/manifest.json',
                '/jpnapp/icon-192.png', '/jpnapp/icon-512.png'];
```

If the phone shows a stale build: delete the PWA from the home screen, open
`euron665.github.io/jpnapp/` in Safari, confirm the change, re-add to home screen. That
clears a stuck service worker.

**GitHub gotcha:** `raw.githubusercontent.com` and even the contents API can serve a **stale
cached copy** for a while. A "missing" upload has more than once turned out to be CDN lag,
not a failed commit. Check the repo page in a browser before concluding an upload failed.

---

# PART 4 — HOW DATA IS DERIVED FROM THE SPREADSHEETS

`update_koku.py` runs in this order:

| § | Step |
|---|---|
| helpers | `get_color`, `parse_rm`, `parse_vocab_cell`, `find_const_bounds`, `replace_const` |
| 1 | KANJI + SKGS |
| 2 | VOCAB (colour-based reading groups) |
| 3 | POS_TAGS |
| 3b | KATAKANA WORDS |
| 3c | SAME-MEANING |
| 4 | Write JSON files |
| 5 | Rebuild `index.html` inlined constants |
| 6 | Bump SW cache version |

### 4.1 Cell formats

A **vocab cell** is one cell holding three lines:

```
word
reading
meaning
```

→ `{'word':…, 'reading':…, 'meaning':…}`. Fewer than 3 lines ⇒ ignored.

A **kanji cell** is `reading\nmeaning` (`parse_rm`).

### 4.2 Reading groups come from CELL FILL COLOUR

This is the least obvious part of the whole system. In `FINAL VOCAB LIST`, each kanji row's
vocab cells are **background-coloured**, and each distinct colour is one *reading group*. The
group's reading (`rg_reading`) is the kanji's reading for that group.

`get_color(cell)` returns the explicit `fgColor` RGB or `None`. Vocab sharing a colour lands
in the same `rgs[]` entry. **If the user removes or changes a fill colour, reading groups
silently regroup.**

Resulting shape:

```json
{ "林": { "rgs": [ { "rg_reading": "はやし",
                     "vocab": [ {"word":"林","reading":"はやし","meaning":"Forest"} ] } ] } }
```

`rg_reading` is also what powers the **conjugation stem hint** and the **Kanji Readings
review mode** — it is the single most load-bearing derived field in the app.

### 4.3 How the constants get into `index.html`

`replace_const(html, NAME, value)` locates `const NAME = …` and replaces the whole value by
**bracket counting** (not regex). Rewritten consts: `KANJI`, `SKGS`, `VOCAB`, `POS_TAGS`,
`KATAKANA`, `SAME_MEANING`.

Hand-written code elsewhere in `index.html` survives a `.bat` run untouched — the pipeline
only replaces those named const blocks. **Never hand-edit the const blocks**; change the
spreadsheet and re-run.

---

# PART 5 — PART OF SPEECH (POS)

### 5.1 Source and alignment

`FINAL POS LIST` maps each vocab **word** → one or more POS tags. The pipeline writes
`POS_TAGS` (`{word: [tags]}`, 3,055 entries). A word with no entry gets no tag and is
**excluded from conjugation review** (`buildConjWordList` skips untagged words).

`tags[0]` is authoritative — it decides the conjugation class.

### 5.2 The tag set, with colours

These are the app's vocab-pill tags (`.vp-tag-*`), coloured by **background**:

| POS value | Label | Hex | Notes |
|---|---|---|---|
| `noun` | NOUN | `#5090d0` | blue |
| `i-adj` | い-ADJECTIVE | `#5db89a` | teal-green |
| `na-adj` | な-ADJECTIVE | `#a07ed4` | purple |
| `u-verb` | う-VERB | `#c9a84c` | gold |
| `u-verb-ru` | う-VERB (る, godan) | `#c9a84c` | gold — godan verb ending in る |
| `ru-verb` | る-VERB | `#c9a84c` | gold — ichidan |
| `suru-verb` | する-VERB | `#c9a84c` | gold |
| `adverb` | ADVERB | `#e8863a` | orange |
| `expression` | EXPRESSION | `#e8e8f0` on `#0c0c0e` | inverted (white chip, dark text) |

**All four verb classes share gold `#c9a84c`** — that is deliberate: colour means "verb",
the label disambiguates the class.

`POS_META` drives the Study/Review-by-POS picker, including the う-verb sub-split:
All · う・つ・る · む・ぶ・ぬ · く・ぐ · す · る (godan).

### 5.3 う-verb vs る-verb — why `u-verb-ru` exists

A verb ending in る can be either ichidan (`ru-verb`) or godan (`u-verb-ru`). The spreadsheet
must say which; it cannot be inferred. Conjugation correctness depends entirely on this tag.

### 5.4 Adverbs / expressions

These were extracted per WaniKani level by a helper script and merged into `FINAL POS LIST`.
They carry no conjugation, so they appear in study and POS review but not conjugation drills.

---

# PART 6 — SAME MEANING, DIFFERENT KANJI

### 6.1 What it is

Curated groups of words sharing an English gloss but written with different kanji
(計る・測る・量る), each with a **nuance note** explaining how they differ. 271 groups.

In the app it powers:
- the teal **⇄ mark** on any study-vocab pill whose word is in a group → tap opens a popup
  with the group and note;
- a **Same-Meaning Words review mode** (Review Vocab picker) that drills every such word and
  shows the note on reveal;
- a **Same-Meaning study view** (Study picker) listing every group as a card with its note inline.

### 6.2 Spreadsheet format — **6 word slots**

```
Meaning | Word 1 | Reading 1 | … | Word 6 | Reading 6 | Note
```

- Words in columns **2/4/6/8/10/12**, readings in **3/5/7/9/11/13**.
- **Note is column 14.** (It used to be column 10 with only 4 word slots; the format was
  widened when the "Condition" group needed a 5th word. Any pre-v2.1 file read with the
  current pipeline yields blank notes — do not reuse old versions.)
- Note cells are newline-separated, one line per word, wrap-text on.
- Word columns are tinted `DCE6F4` purely for readability.

### 6.3 Note style

One line per word: **`WORD = gloss — when to use it`**. Concise, practical, distinguishing
nuance only. Japanese inline where useful. No markdown inside cells.

```
近所 = neighborhood — the area right around your home; where the neighbours are (近所の人).
付近 = vicinity — the area near a given point or landmark (駅の付近, この付近).
```

### 6.4 The workflow for adding groups

The detector finds the **words**; a human/assistant writes the **notes**. Not automatable —
nuance is knowledge.

1. User adds vocab as normal, runs `update_koku.bat`.
2. User runs `same_meaning_detector.py` → writes `same_meaning_NEW_candidates.xlsx`
   (groups containing at least one word not yet covered).
3. User sends that file to the assistant.
4. Assistant **curates first**, then writes notes:
   - **Drop English homonyms.** The detector matches on English strings, so it produces false
     positives: ad → 広告/紀元後 (A.D.), ring → 鳴る/指輪/輪, fine → 結構 (okay) /罰金 (penalty),
     tie → 結ぶ (knot)/引き分け (draw).
   - **Drop same-word pairs.** 支援/支援する, 輸出/輸出する, 自動販売機/自販機 — a する-verb form or
     an abbreviation is not a nuance pair.
   - **Check the group doesn't already exist** before appending — duplicates have been created
     twice this way (Support, Neighborhood). Search the Meaning column first.
   - **Add words the detector missed.** It keys on exact gloss, so 閉まる ("to close") was
     missed for a group keyed "close sth", and 触る ("to touch sth") for "touch".
5. Assistant hands back an updated spreadsheet with the version bumped.
6. User runs the `.bat` and uploads.

### 6.5 The REJECTED set

`same_meaning_detector.py` carries a `REJECTED` set of meaning keys that would otherwise be
re-flagged forever (because their words are deliberately never added):

```
ad, be raised, character, close, compromise, conclusion, face, kind, mercury, nature,
order, ring, second, spring, state, superior, typical, support, tie, fine,
vending machine, export
```

Add a key here (lowercase, leading `a/an/the/to` stripped) to silence a recurring false positive.

---

# PART 7 — REVIEW & STUDY MODES

All reviews share one engine: `startReview()` → `loadItem()` → `handleSubmit()` →
`nextItem()` → `endReview()`, with `queue`, `currentItem`, `itemKey`, `itemFailed`,
`rPass`/`rFail`, `updateProgress()`, `updateScoreDisplay()`, `armNext()`, `skipItem()`.

A mode = *build a `queue` of items, then call `startReview()`*. `loadItem()` branches on
`currentItem.type`.

### 7.1 Item types

| `type` | Mode | Answer |
|---|---|---|
| `kanji` | Review Kanji | meaning, then reading (two steps) |
| *(vocab, default branch)* | Review Vocab | meaning, then reading |
| `kana` (`_type`) | Review Kana | romaji |
| `katakana` | Review Katakana Words | meaning |
| `conj` | Review Conjugation | the conjugated form, typed in hiragana |
| `kreading` | **Kanji Readings** | multiple-choice tiles |

### 7.2 Study modes

Study Kana · Study Katakana Words · Study Kanji & Vocab (swipe between kanji and vocab
views) · Study by Part of Speech · Study Conjugation · Study Grammar.

Plus two **Special Drills** in the Study picker: *1 Kanji – 1 Vocab* and *Same-Meaning Words*.

### 7.3 Special Drills (Review picker)

- **1 Kanji – 1 Vocab** — every kanji whose total vocab across all reading groups is exactly 1
  (currently 62). Live-computed by `getSingleVocabItems()` on every open, so it can never go
  stale. The rare kanji is highlighted **gold** in the word.
- **Same-Meaning Words** — every word in a same-meaning group that exists in the vocab; the
  note shows on reveal.
- **Kanji Readings** — see below.

Each opens a **session-size prompt** (blank = all) before starting.

### 7.4 Kanji Readings (multiple choice)

- **Correct answers** = the `rg_reading` values of that kanji (889 kanji have at least one).
- **Mechanic:** select *every* reading that belongs to the kanji, then Confirm.
- **Tile count:** `max(6, min(12, correct × 2))` → 1–3 readings = 6 tiles, 4 = 8, 5 = 10, 6+ = 12.
- **Distractors** are drawn from the pool of 773 distinct readings and **scored for
  similarity**: same length preferred, bonus for sharing the first or last kana. So 林 (はやし)
  gets はなし・はたら・むかし・ひがし・わたし, not random long kun-readings.
- **Reveal colours:** green = correct pick, red = wrong pick, **gold = a reading you missed**.
- Reading counts per kanji: 348 have 1, 363 have 2, 118 have 3, 44 have 4, 11 have 5, 2 have 6,
  one each at 7, 8 and 11. The 11 is **日** (に・にち・にっ・か・び・ひ・じつ・きょう・とい・て・た) —
  at the 12-tile cap it gets one distractor. Accepted.

### 7.5 Conjugation review

`conjugate(reading, pos, form, formality)` covers `present-aff / present-neg / past-aff /
past-neg` × `plain / polite`, plus `te`. Irregulars (する・くる・いく・ある・いる) are table-driven.
`conjugateAlts()` supplies accepted alternates (e.g. polite negative adjectives).

**Stem hint** — a blurred line between the word and its translation; tap to reveal, re-blurs
on every new item. The stem is **the reading of the word's kanji portion**, computed by
stripping the word's trailing okurigana off the reading:

```
楽しむ (たのしむ) → たの      心配する (しんぱいする) → しんぱい
飲む  (のむ)    → の        話す   (はなす)      → はな
```

Deriving it from the reading-group alone was wrong for multi-kanji words — 心配する is filed
under its *second* kanji 配, which gave ぱい. Verified across all 464 verbs, no failures.

### 7.6 Wrong-script guard

Every submission passes a script check before grading:
- kana-answer modes reject leftover Latin characters;
- text-answer modes reject stray Japanese.

It shows an amber "try again" nudge, **does not count a failure and does not advance**. This
stops a mistyped keystroke from burning an item. A genuinely wrong *Japanese* answer still fails.

### 7.7 Review UX rules that were explicitly requested

- Manual advance: Enter/Confirm turns into **Next →**; the app never auto-jumps.
- The keyboard must stay up through an item — the input is never disabled; buttons use
  `onmousedown="event.preventDefault()"` and `ontouchend` to avoid losing focus.
- Review order on the kanji screen: Meaning & Reading leftmost.
- ✕ exits to the picker.

---

# PART 8 — THE APP UI

### 8.1 Design tokens (`:root`)

```
--bg #0c0c0e   --surface #141416   --surface2 #1c1c20   --border #2a2a30
--text #e8e8f0 --text-dim #666680  --text-muted #333345
--accent #e8193c   --accent-glow rgba(232,25,60,0.4)   --gold #c9a84c
--pass #1a3a2a  --fail #3a1a1a  --pass-text #4af080  --fail-text #f04a4a
--color-noun #5090d0  --color-i-adj #5db89a  --color-na-adj #a07ed4
--furi-hit #ff9d3c
--font-jp 'Noto Sans JP'  --font-ui 'Noto Sans JP'  --font-mono 'DM Mono'  --r 4px
```

Fonts come from Google Fonts — **the only external dependency**. Teal `#3dd6c8` is used for
same-meaning marks, search flash and pill selection, and is **hardcoded**, not a token.

### 8.2 Screens and navigation

Screens (`<div id="…" class="screen">`): `pos-picker, study-pos, skg-picker, study-kanji,
study-vocab, study-single, study-sm, mode-select, vocab-kanji-picker, order-select,
study-katakana, katakana-picker, review, study-kana, study-verbs, kana-picker, done,
conj-picker, grammar`.

`goTo(id, mode)` sets `currentMode`, runs that screen's builder, then `showScreen(id)`.
`showScreen` toggles `.screen.active` and the body state classes
`grammar-active`, `review-active`, `home-active`, `gr-reading`.

`currentMode` ∈ `study-kanji | study-vocab | review-kanji | review-vocab | study-pos |
review-pos` and drives the shared SKG picker.

### 8.3 Floating navigation (no header bars)

The old full-width header bar is gone. Every screen shows **two floating 34×34 icons at
top-left**: a back arrow and a magnifier.

- Both are drawn as **stroked SVG** (`stroke-width:2`, round caps) — the back arrow via a CSS
  `mask` with `background-color:currentColor`, so one rule covers all 38 back buttons without
  touching markup, and hover colour still works.
- `.screen-header` is transparent, `pointer-events:none`, and no longer reserves space;
  `.screen-title` is hidden.
- Content padding dropped from **56px → 16px** desktop, **50px** mobile (mobile puts the two
  icons side by side instead of stacked, to keep the vertical footprint to one row).
- On Home there is no back button, so the magnifier takes the top slot.

### 8.4 Collapsible search

The search bar is hidden by default (`transform:translateY(-100%)`) and **overlays** on
demand — it does not push content. Toggled by the magnifier; closes on outside tap, Escape,
or screen change.

- `#search-results` sits at `top:88px` (the bar's height).
- When the bar opens both floating icons slide **down 88px** (back → 97px, glass → 138px),
  animated to match the bar's own `0.22s` easing.
- **Emptying the input must not close the bar.** `clearSearch()` deliberately does *not*
  remove `search-open` — that was a bug once.
- Search supports three input modes: EN / かな / カナ, with live romaji→kana conversion.
- The magnifier is **hidden in grammar and review** on purpose — grammar is for focused
  reading, and search would pull the user out of it.

### 8.5 Key component classes

| Area | Classes |
|---|---|
| Buttons / menu | `.btn-primary` `.btn-secondary` `.menu-title` |
| Set picker | `.skg-grid` `.skg-card` `.picker-label` `.picker-head` `.sort-cycle` `#quick-drills` |
| Study Kanji grid | `.study-kanji-list` `.study-kanji-card` `.skh-kanji/.skh-reading/.skh-meaning/.skh-count` |
| Study Vocab | `.svk-card` → `.svk-header` (`.svk-kanji/.svk-info/.svk-reading/.svk-meaning/.vki-count`) + `.svk-body` → `.rg-block` → `.rg-vocab-grid` → `.vocab-pill` |
| Vocab pill | `.vp-word/.vp-reading/.vp-meaning`, `.vp-tags/.vp-tag-*`, `.sm-mark`, `.vp-selected`, `.sr-flash` |
| Kanji→vocab popup | `.kv-modal-overlay/.kv-modal/.kv-modal-header/.kv-modal-kanji/.kv-modal-close` |
| Same-meaning popup | `.sm-modal-overlay/.sm-modal/.sm-meaning/.sm-words/.sm-word/.sm-w/.sm-r/.sm-note` |
| Review | `.review-header/.review-progress/.review-known-meaning/.review-input/.answer-reveal/.review-footer`, `.conj-stem-hint`, `.kr-tile` |
| Search | `.search-result-item/.sr-kanji/.sr-info/.sr-sub/.sr-main/.sr-meta` |
| Katakana | `.kata-cat-card` (+ `.cat-open`) · Kana: `.kana-script-btn/.kana-cell` |

### 8.6 Interactions worth knowing

- **Two-step tap on a vocab pill:** first tap selects (teal border), second opens a kanji
  breakdown popup listing the word's component kanji; tapping one jumps to that kanji's set.
- **1 Kanji – 1 Vocab:** tapping the vocab opens a confirm popup (kanji, reading, meaning, set)
  with **Stay / Go to set →** — it never jumps directly.
- **Search jumps** scroll the target card to the top of the screen (`block:'start'`, with
  `scroll-margin-top:68px` on `.svk-card` so the floating icons don't cover it) and flash it teal.
- `.screen` is `position:absolute; inset:0; overflow-y:auto` — **the screen is the scroll
  container, not the window.** Any `window.scrollTo` inside a screen silently does nothing;
  use `document.getElementById('<screen>').scrollTo(...)`.

---

# PART 9 — THE GRAMMAR MODULE

200 cards (N5 + N4), 20 chapters, 2,198 examples, 8,758 furigana readings. It lives inside
`index.html` as `<div id="grammar" class="screen">`. There is no separate grammar file — all
grammar work happens in `index.html`.

### 9.1 Structure

```
chapter-section
└─ card  (id="card-N", may carry "lv-n4" and "examples-hidden")
   ├─ title-block   → level-badge · card-coord · title · subtitle-romaji · subtitle-meaning
   ├─ formula-section → formula-box → formula-box-label (OUTSIDE the box)
   │                    formula-box-rows  ← overflow:hidden, clips furigana
   │                    └─ formula-row-group → formula-register (casual/formal headliner)
   │                                            formula-row → formula-var-group
   │                                              ├─ formula-var       (slot, word-type coloured)
   │                                              └─ formula-var-type  (subtitle, same colour)
   │                                              formula-plus "+" · formula-gram (red JP)
   │                                              formula-strike (struck kana being removed)
   ├─ description-section → section-label + description (prose)
   └─ examples → example-jp · example-en · example-note · example-tags
```

| Element | Count |
|---|---|
| cards | 200 (ids 1–200, contiguous) |
| chapters | 20 |
| examples | 2,198 (each has jp + en + note, none empty) |
| etags | 9,444 |
| formula boxes | 40 · rows 471 · var slots 641 · var-types 267 · strikes 44 |
| tables | 42 · rows 315 · bullets 138 |
| italics | 1,121 · N4-marked 412 |

### 9.2 Colours — THE SOURCE OF TRUTH

**Word-type colours are semantic. A colour *means* a part of speech. Never reuse one.**

| Colour | Meaning | Uses |
|---|---|---|
| `#c9a84c` | **VERB** (gold) | 2,536 |
| `#5db89a` | **い-ADJ** | 640 |
| `#a07ed4` | **な-ADJ** | 618 |
| `#5090d0` | **NOUN** | 1,458 |
| `#888899` | neutral / multi-type | 223 |
| `#8080f0` | **て-FORM** (indigo) | 87 |
| `#e8193c` | Japanese constructs (`--accent`) | everywhere |

**Example tags** — `<span class="etag" style="color:X; border-color:X;">LABEL</span>`
(colour duplicated in both properties):

| Tag | Hex | Count |
|---|---|---|
| AFFIRMATIVE | `#4af080` | 1,599 |
| PRESENT | `#4060e0` | 1,572 |
| VERB | `#c9a84c` | 1,268 |
| POLITE | `#a0a0a0` | 1,183 |
| PLAIN | `#c060c0` | 1,078 |
| PAST | `#e09080` | 673 |
| NEGATIVE | `#f04a4a` | 616 |
| NOUN | `#5090d0` | 521 |
| QUESTION | `#e07840` | 481 |
| い-ADJ | `#5db89a` | 225 |
| な-ADJ | `#a07ed4` | 199 |
| て-FORM | `#8080f0` | 4 |
| NEUTRAL | `#8898b8` | 1 |
| A / N | `#4af080` / `#f04a4a` | 12 each — **Chapter 19 only**, abbreviated polarity in wide tables |
| CASUAL / VERY CASUAL | `#c060c0` | share the PLAIN colour |

Fixed tag order:
`VERB · い-ADJ · な-ADJ · NOUN · PRESENT · PAST · AFFIRMATIVE · NEGATIVE · PLAIN · POLITE · QUESTION`

**Tags follow the final predicate of the sentence**, not every clause. This looks like a bug
in multi-clause examples. It isn't — one card states the rule explicitly.

**Derived-form colours (Chapter 19 / appearance family):** passive `#a07ed4` · causative
`#e8863a` · causative-passive `#3dd6c8` · potential `#9ad04a` · ば-form `#e0608c` ·
volitional `#7fd4f5` · たら `#f0a35c` · 時-clause `#b48ce8` · hearsay そうだ `#4ad6b0`.

**Level badges:** N5 `#3dd6c8` · N4 `#4ecb6e` · N3 `#a07ed4` · N2 `#e8863a` · N1 `#e8193c`.

### 9.3 Register scheme — 4 levels, deliberately

| Register | Count | Meaning |
|---|---|---|
| POLITE | 520 | です/ます |
| PLAIN | 293 | neutral plain — planning, firm statements, writing |
| CASUAL | 171 | plain form in conversational register — peer speech, casual offers |
| VERY CASUAL | 7 | contracted colloquial — なきゃ, ちゃ, んじゃ |

This granularity is **intentional and was reviewed** — do not collapse CASUAL into PLAIN.

### 9.4 Typography & spacing

| Element | Size | Line-height | Letter-spacing |
|---|---|---|---|
| `.title` | 58px (38px ≤480px) | 1 | — |
| `.subtitle-meaning` | 21px | — | — |
| `.subtitle-romaji` | 15px | — | 0.05em |
| `.description` | 14px | **1.8** | — |
| `.example-jp` | 19px | default | — |
| `.example-en` | 13px | — | — |
| `.example-note` | 11px | — | 0.04em |
| `.etag` | 9px | — | 0.1em |
| `.card-coord` | 10.5px | — | 0.12em |
| `.section-label` | 10px | — | **0.22em** |
| `.formula-gram` | 26px (22px ≤480px) | — | — |
| `.formula-var` | 14px | — | 0.12em |
| `.formula-var-type` | 9px | — | 0.18em |
| `.formula-register` | 11px | — | 0.22em |
| `.recap-rule` | 13px | 1.6 | — |
| `.recap-table` | 13px | td 1.5 | th 0.18em |
| `.guide-p` | 14px | 1.8 | — |
| `.fw::after` (furigana) | `clamp(9px, 0.56em, 13px)` | 1 | 0.02em |

- Mobile breakpoint: **`@media (max-width: 480px)`**.
- Content constrained to **`max-width: 640px`**, centred; chapter/guide sections
  `padding: 0 24px 80px` (16px on mobile).
- Weight **300** is the default for prose and Japanese; 400/500 for labels and headings.
- Tables are horizontally scrollable on mobile (`.recap-table` → `display:block; overflow-x:auto`)
  with an `.xscroll-track` indicator.

### 9.5 Separators — conventions, not decoration

| Glyph | Use | Count |
|---|---|---|
| `—` em dash | gloss separator, appositives, `Usage N — label` | 2,327 |
| `·` middle dot | coordinates `CH.4 · CR.7`, English alternatives | 1,137 |
| `→` arrow | conjugation chains | 270 |
| `・` JP middle dot | alternatives **inside Japanese** (これら・それら) | 246 |
| `〜` wave dash | placeholder in **grammar names** (〜すぎる) | 221 |
| `~` ASCII tilde | placeholder in **English** glosses (too ~) | 323 |
| `…` ellipsis | trailing-off speech | 59 |

Never swap `・` for `·` or `〜` for `~` — they mark Japanese vs English context.

### 9.6 Prose conventions

**Translations vs terms:**
- *Italics, no quotes* — the English is a direct translation of preceding Japanese:
  `遅れてすみません — *sorry for being late*`
- `"Plain double quotes"` — the English word is being named/taught as English:
  `the "Wanna?" of English`

A linking verb substitutes for the dash — **no dash after one**: `20歳以上 is *20 and up*`.

Current state: **1,121 italic runs**, **192 quoted runs** (386 quote chars — always 2×). Every
one was reviewed individually. **Do not re-run a blanket regex over them.**

Trailing punctuation: `,` and `;` go outside `<em>`. `.` `?` `!` go outside **only** when the
quote ends its sentence. `...` and `…` always stay inside.

**Enumerations → bullets** (`.recap-rules`/`.recap-rule`, 138 of them) only for genuinely
parallel lists. Do not multiply single-sentence paragraphs.

**Cross-references** always `covered in card CH.n · CR.m` — never "next card" or a topic name.
50 refs, all resolve.

**Register headliners** (`casual`/`formal`) go **above** the formula box, never inside it.

### 9.7 Formula box semantics

- **Brackets hold attaching kana only** — `[て]`, `[る]`.
- **Deletion is shown by `formula-strike`**, not by brackets (44 uses).
- Word-type qualifiers (`plain`, `stem`, `material`) are **subtitles** (`formula-var-type`), never brackets.
- A bare `Verb` with no subtitle means **any form, tense or polarity**.
- `Verb stem` means the ます-stem; **ます is removed**. Never write "stem" twice or mention ます in the slot.

### 9.8 Furigana system

```html
<span class="fw" data-r="よ">読</span>む
```

Only the **kanji run** is wrapped; okurigana stays outside. Readings are **word-level and
context-aware**, never per-character — 一人 = ひとり, 今日 = きょう, 上手 = じょうず.

Rendering is **absolutely positioned**, not `<ruby>`:

```css
.fw { position: relative; }
.fw::after {
  content: attr(data-r);
  position: absolute; left: 50%; bottom: 100%;
  transform: translateX(-50%);
  font-size: clamp(9px, 0.56em, 13px);
  opacity: 0; pointer-events: none;
  transition: opacity .12s ease;
}
```

The reading is out of normal flow, so revealing it **never reflows the page**. That is the
whole point — do not switch to `<ruby>`, which reserves space and makes text jump on tap.

**Three modes** (html class, persisted in `localStorage` key **`koku-furigana`**):

| Class | Behaviour |
|---|---|
| `furi-off` | `content:none` — renders exactly as pre-furigana |
| `furi-click` | **default** — tap a kanji to show, tap again to hide, one at a time |
| `furi-on` | all readings always visible |

Tapped kanji and its reading turn **`--furi-hit` `#ff9d3c`**.

**Line-height bumps apply only in `furi-click`/`furi-on`:** `.example-jp`/`.example-note`/
`.recap-rule` → 1.95; `.description`/`.guide-p` → 2.15; `.gr-name` → 2; `.formula-gram` → 1.5;
`.title` padding-top 17px; `.formula-box-rows` padding-top 13px (it has `overflow:hidden` and
clips readings). Switching mode reflows once, deliberately. Tapping never does.

Click handling is a **capture-phase** listener on `document` with `stopPropagation()` so a tap
on a word doesn't also fire card/example toggles. Escape closes.
`touch-action: manipulation` kills the 300ms delay.

**Reading accuracy — 145 corrections were applied by hand.** Generated with SudachiPy +
sudachidict-core, then corrected over four rounds. **If you regenerate, you reintroduce all of
them.** Fault classes:
1. Context-blind defaults — 辛い → つらい where it means からい; 言う → ゆう instead of いう;
   何時 → いつ instead of なんじ; 間 → ま instead of あいだ.
2. **私 read わたくし in all 59 places** → corrected to わたし.
3. い-adjective stems taking on-readings — 高くて→こう, 寒くなくて→かん, 重い→じゅう, 速さ→そく.
4. Compound sound changes — 一本→いっぽん, 一週間→いっしゅうかん, 九時→くじ, 七時→しちじ,
   綿菓子→わたがし, 十/二十/三十分→…じゅっぷん (but 5分 and 十五分 keep ふん).
5. Card context overrides the tokenizer — CH.19 · CR.1 teaches 開ける/開く so 開 is **あ** there;
   it stays **ひら** only where it means "open a business / hold an event".

**Validator worth re-running after any reading change:** compare each card's title reading
against the same word in its body. It caught two genuine title errors (〜方 ほう vs かた;
〜風 かぜ vs ふう).

### 9.9 Grammar JS

`openChapter(n)` · `closeChapter()` · `openGuide(which)` · `closeGuide()` · `gotoCard(ch,id,from)` ·
`scrollToCard(id)` · `toggleExamples(btn)` · `buildFilterBar(card)` · `filterExamples` ·
`resetFilter` · `updateGroupLabels` · `collapseAllExamples` · `applyLevels` · `toggleN4` ·
`n4Enabled` · `openSettings`/`closeSettings` · `setFuri` · `grammarToTop` · `grReading`.

The filter bar is generated from each card's own etags at runtime — correctly-tagged examples
need **zero** JS changes.

### 9.10 Grammar navigation in-app

The module's chapter/guide header bar was replaced by a **single floating ← icon** (top-left),
reclaiming ~89px per chapter (20px padding ×2 + 48px margin + border). A **↑ scroll-to-top
icon** sits under it, shown only while `body.gr-reading` is set — i.e. only inside a chapter or
guide, never on the chapter menu or elsewhere in the app.

N4 content is toggled by `applyLevels()`, which sets `hide-n4` on **`document.body`**. The
hiding rules must therefore read `body.hide-n4 #grammar .card.lv-n4` — scoping them *inside*
`#grammar` breaks the toggle silently.

### 9.11 Unknown-kanji audit

Goal: every example sentence uses only kanji the user has learned. The known set is derived
from `FINAL KANJI & SKG LIST` (first line of each kanji cell) — 888 kanji currently.

Rules agreed with the user:
1. Kanji **inside** `<span class="jp">…</span>` is the grammar construct — **never change it**,
   even if unknown. Only kanji outside the span counts.
2. No kana-only substitution. Replace the word with another whose kanji are all known.
   Unfamiliar *vocabulary* is fine; unfamiliar *kanji* is not.
3. The replacement must keep the etags true — same word type on the final predicate, same
   tense, polarity, register, question status.
4. The sentence must make real-world sense. No filler.
5. Update `example-en` to match; update `example-note` only if its nuance no longer fits.

**Current state:** 16 sentences use unknown kanji, and this is **accepted, not a defect** —
they were chosen to keep examples sensible: 散 掃 除 磨 遊 簡 背 寿 菜 窓 廊 戻 捨 駐 匂 源.
The permanent fix is to add these to `FINAL KANJI & SKG LIST` on a future batch, after which
the exceptions dissolve by themselves.

There is no tooling for this — the whole check is a few lines: build the known set from the
kanji spreadsheet, strip
`<span class="jp">…</span>` out of every `example-jp`, and flag any remaining CJK character
not in that set.

```python
import re, openpyxl, glob, os
kf = max(glob.glob('FINAL KANJI*.xlsx'), key=os.path.getmtime)
ws = openpyxl.load_workbook(kf, read_only=True, data_only=True).active
KJ = re.compile(r'[一-鿿]')
known = {c for row in ws.iter_rows(values_only=True) for v in row
         if isinstance(v, str) for c in KJ.findall(v.split('\n')[0])}
h = open('index.html', encoding='utf-8').read()
for m in re.finditer(r'<div class="example-jp">(.*?)</div>', h, re.S):
    outside = re.sub(r'<[^>]+>', '', re.sub(r'<span class="jp">.*?</span>', '', m.group(1), flags=re.S))
    bad = {c for c in KJ.findall(outside) if c not in known}
    if bad:
        print(''.join(sorted(bad)), re.sub(r'<[^>]+>', '', m.group(1)).strip())
```

---

# PART 10 — KATAKANA

260 words from `FINAL KATAKANA VOCAB`, grouped by category. The pipeline derives **romaji
automatically** from the katakana via a kana→romaji table (handling small kana, long vowels
and っ), so the spreadsheet does not store romaji.

Each word may carry a note typed as **false-friend / origin / clipping / usage**, rendered
grouped by type. Study view: collapsible category cards, romaji italic and smaller, glow on
closed categories only.

Review: meaning-typed, with a **False-Friends-only** option in the picker. On a correct answer
the romaji appears under the word.

Katakana words are searchable, and the search bar has a **カナ** input button that converts
typed romaji straight to katakana.

---

# PART 11 — CONJUGATION TABLES (Study Conjugation)

`makeVerbTableHTML(verb)` renders one table per example verb from the `VERB_CONJ` data
structure. Columns: **PLAIN REGISTER** (`#c060c0`) · **POLITE REGISTER** (`#a0a0a0`).
Rows: present-aff → past-aff → present-neg → past-neg, plus a て-form row.

Row labels are **colour-coded text, no boxes** (10px, uppercase, `letter-spacing:0.08em`):

| Label | Hex |
|---|---|
| PRESENT | `#4060e0` |
| PAST | `#e09080` |
| AFFIRMATIVE | `#4af080` |
| NEGATIVE | `#f04a4a` |
| て-FORM | `#8080f0` |

Each cell shows the ending tag (`vc-et`, with romaji in `vc-et-rom`) plus the full form and its
romaji. Exceptions get `exc-mark`. Types covered: る-verbs, う-verbs (split う・つ・る / む・ぶ・ぬ /
く / ぐ / す), irregulars, and い/な-adjectives.

The copula (だ/です) cards in the grammar module carry their own **Conjugation — all forms**
table: Register × Tense × Affirmative × Negative, register/tense cells rendered as coloured
`etag` chips, Japanese forms in `--accent` red with italic romaji in brackets
(`.recap-rom`, 11px, `--text-dim`, mono).

---

# PART 12 — VERIFICATION (run after every change)

```
balance        div == , span == , p == , tr ==
cards          200, ids 1..200 contiguous
examples       2198 jp / 2198 en / 2198 note, none empty
etags          9444      furigana 8758, orphan .fw spans 0
navigation     200 coords, 200 toc targets, refs all resolve
italics        <em> open == close, nested 0, empty 0
quotes         quote chars == 2 × quoted runs
typography     curly quotes 0, curly apostrophes 0
css            braces balanced
ids            no duplicate id="…"
app intact     19 screens present, all app functions present, 7 data consts present
javascript     node --check clean on the extracted <script>
tail           ends </html>
```

Extract the script and check it:

```python
js = n[n.find('>', n.find('<script'))+1 : n.rfind('</script>')]
open('/tmp/x.js','w',encoding='utf-8').write(js)
# then: node --check /tmp/x.js
```

Then **look at it** — render at 900px and 390px. Several bugs (orphan `)`, four-dot ellipses,
bullet clutter, the search bar covering the floating icons) only showed up visually.

---

# PART 13 — WORKING AGREEMENTS

- Never act without being asked; never "fix" things the user didn't ask about.
- Flag the approach first for complex work; produce output directly for simple work.
- Back up to `/tmp` before any large structural operation, and verify before installing.
- Bump `sw.js` on every change that ships.
- `index.html` is the only app file — there is no second copy of the grammar module to mirror.
- The spreadsheets are the source of truth for data; `index.html` is the source of truth for
  behaviour; the grammar module is the source of truth for colour.
