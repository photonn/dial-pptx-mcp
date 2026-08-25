# Deck design guidance

Guidance for the agent building a deck through this server. It answers "what
should this slide look like", which no individual tool can tell you: the tools
will happily place 22 lines of 11pt text in a corner and report success.

Read the mode section first — almost everything downstream depends on whether
you were given a corporate template.

---

## 1. Two modes, and they have opposite defaults

**Template mode — the default here, and the one this server is built for.**
The user supplied a `.pptx`/`.potx` that carries their organisation's identity:
masters, layouts, theme colours, fonts, logo placement, and usually a set of
already-designed example slides.

In this mode your design decisions are mostly *selection*, not invention:

- **Never introduce your own palette or fonts.** The template's colours are the
  brand. A slide you styled yourself will not match the one before it, and the
  mismatch is more visible than anything you could gain.
- **Build by duplication.** Find the template slide whose structure fits your
  content, `duplicate_slide` it, and replace the text. That inherits every
  piece of decoration the template designer put there — rules, panels, icons,
  logo positions — none of which `add_slide` gives you, because `add_slide`
  builds a bare slide from a layout and layouts hold only placeholders.
- **Delete what you don't fill.** If a template slide has four team-member
  blocks and you have three people, `delete_slide` is wrong (you need the
  slide) but the fourth block must go entirely — picture, name, title, the lot.
  A blank block reads as a bug. Emptying its text is not enough.
- **Read the template before planning.** `get_presentation_info` and
  `render_slide_previews` together tell you what slides you have to work with.
  Plan the deck against those, not against an ideal deck you then force the
  template to imitate.

**Scratch mode — no template supplied.** Now you own every visual decision, and
the rest of this document applies in full.

---

## 2. Structure before styling

A deck is an argument, not a document. Before placing anything:

- **One idea per slide.** If a slide needs two sentences to summarise, it is
  two slides. Splitting is free; cramming is not.
- **Give it an arc.** Something like: what the situation is → why it matters →
  what you found → what you propose → what you need. Section dividers earn
  their place when the deck runs past ~10 slides.
- **Titles carry the message.** "Q3 revenue" is a label. "Q3 revenue grew 18%,
  driven by enterprise renewals" is the point of the slide, and it means the
  audience gets it even if they read nothing else. Write the assertion.
- **Vary the layout.** Ten consecutive title-and-bullets slides read as a
  transcript. Alternate: a full-bleed statement slide, a two-column
  comparison, a chart slide, a three-card row.
- **Put speaker notes to work.** Detail that does not belong on the slide
  belongs in `manage_speaker_notes`, not in a 10pt paragraph at the bottom.
  Text on a slide is read by the audience; notes are read by the presenter.

---

## 3. Layout

Whatever the mode, the geometry rules are the same. Slide sizes here are
usually 13.33×7.5in (16:9 widescreen) or 10×5.63in — check `slide_width` from
`get_presentation_info` rather than assuming, because placing content for the
wrong canvas silently puts it off the slide.

- **Margins: 0.5in minimum** on every edge, and treat that as a hard floor, not
  a target. Content that runs to 0.2in from the edge looks like it escaped.
- **Pick one gap size and keep it.** 0.3in or 0.5in between blocks, used
  consistently, is what makes a deck look designed. Mixed gaps read as
  accidental even when nothing overlaps.
- **Align to a small number of columns.** Two-column, three-card, and 2×2 grid
  cover most slides. Elements that are conceptually parallel must share an
  x-position (or y-position) to the pixel — visual QA flags misalignment, and
  it is much cheaper to place them right than to repair them.
- **Leave the slide two-thirds full.** White space is what makes the remaining
  third legible. If content will not fit at the sizes in §4, the answer is a
  second slide, not 10pt type.

Layout patterns worth reaching for:

| Pattern | When |
|---|---|
| Statement slide: one sentence, large, centred | Section opener, key finding |
| Two-column: text left, visual right | Explaining something with an example |
| Three or four cards in a row | Parallel items — options, phases, pillars |
| 2×2 grid | Two axes of comparison, or four independent points |
| Large stat + caption | A single number that is the whole point |
| Chart with the takeaway as the title | Any data slide |

---

## 4. Type

**Sizes** (points, for a 16:9 deck):

| Element | Size |
|---|---|
| Slide title | 32–44, bold |
| Section header on a busy slide | 20–24, bold |
| Body text and bullets | 14–18 |
| Card / table body | 12–16 |
| Captions, sources, footnotes | 10–12, muted colour |

The contrast between title and body is what makes a slide scannable. A 24pt
title over 18pt body is not a hierarchy — the reader cannot tell what to read
first. Keep at least a 2× ratio between title and captions.

**Font choice, and why this server cares.** Fonts you write into a `.pptx` are
rendered by the user's PowerPoint. But this server's visual QA renders through
LibreOffice, which does not have Microsoft's fonts and substitutes its own.
Some substitutions are *metric-compatible* — identical character widths, so
lines wrap identically and a QA screenshot tells the truth about fit:

| Font you write | Rendered in QA as | Fit verdict |
|---|---|---|
| Arial, Helvetica | Liberation Sans | exact |
| Times New Roman | Liberation Serif | exact |
| Courier New | Liberation Mono | exact |
| Calibri | Carlito | exact |
| Cambria | Caladea | exact |

Everything else — Georgia, Verdana, Trebuchet MS, Segoe UI, Garamond,
Consolas, Impact, Calibri Light — is substituted by *similarity*, and the
widths differ. Text that overflows in the QA render may fit in PowerPoint, and
text that fits may overflow.

So:

- **Body text, where fit matters, should be a metric-safe font** in scratch
  mode. Then you can trust `visual_inspect_slides` on overflow.
- **Titles and short accents can use anything** — they have slack, and a
  wrong-by-5% width estimate on a three-word title changes nothing.
- **In template mode, use the template's fonts regardless.** Brand beats QA
  precision. `validate_presentation` will report the deck's fonts as an `info`
  problem; that is a note on how to read QA results, not a defect to fix.
- **Never leave Aptos in a deck.** It is Office's post-2023 default, has no
  substitute in this renderer, and is missing from Office installs older than
  2024 — unreliable at both ends.
- **Two font families is the maximum.** One for headings, one for body. A
  serif heading over a sans body is a reliable pairing; so is one family at two
  weights.

---

## 5. Colour

**Template mode: use the theme.** The template's theme colours are already
coherent and already the brand. Take accents from them.

**Scratch mode**, build a small palette and commit to it:

- **One dominant colour** carrying most of the coloured area, **one supporting
  tone**, and **one accent** used sparingly — for the thing you want looked at
  first. Roughly 60/30/10 by area. Three colours at equal weight read as
  indecision.
- **Choose for the subject.** Defaulting to corporate blue for every deck is
  the single most obvious tell of an unconsidered design. Let the topic
  suggest the temperature.
- **Set the light/dark rhythm deliberately.** Dark title and closing slides
  with light content slides between them is a reliable structure; so is
  committing to dark throughout. What does not work is alternating at random.
- **Contrast is a hard requirement, not a preference.** Body text needs to be
  clearly readable against its background — mid-grey on white, or light grey
  on a mid-tone, both fail. This applies to icons and chart elements too, not
  just text. Low contrast is the most common defect the visual reviewer
  reports and the most annoying to repair late.
- **Backgrounds: white or a deliberate brand colour.** Warm off-white and
  beige defaults look dated and were not chosen by anyone.

---

## 6. Charts and tables

- **Every chart needs a title that states the finding**, not the variable.
  "Enterprise renewals drove Q3 growth" beats "Revenue by segment".
- **Label the data, drop the furniture.** Turn data labels on when there are
  few enough points to read; then the value axis and most gridlines are
  redundant. A single-series chart does not need a legend at all — say what it
  is in the title.
- **Pick the form for the comparison.** Bars for comparing categories,
  lines for change over time, stacked bars only when the total genuinely
  matters as well as the parts. Pie charts work for two or three slices and
  fail past that.
- **Charts must be native chart objects** (`add_chart`), never an image of a
  chart. Native charts stay editable, scale cleanly, and can be repaired by
  the visual loop; a picture of a chart can only be deleted.
- **Tables are for lookup, not for narrative.** More than ~6 rows or ~5
  columns and the audience stops reading. If the point is a comparison, make
  it a chart or a set of cards.
- **Table text is text.** It overflows, wraps into unreadable stacks, and gets
  clipped by row heights like anything else. The visual reviewer checks cells;
  give columns enough width up front.

---

## 7. Images

- **Every slide benefits from something that isn't text** — a chart, a photo,
  an icon, a shape that organises the content. A wall-to-wall bulleted list is
  forgettable even when it's correct.
- **Place images with `add_image_from_dial_url`, not base64.** The bytes go
  server-side and never through your context.
- **Never distort.** Use `fit="contain"` (or `"cover"` when the box must be
  filled and cropping is acceptable). `fit="stretch"` exists but nothing in the
  repair whitelist can un-distort a picture afterwards, and
  `validate_presentation` reports it.
- **Give an image real size.** A 1.5in photo in a corner is decoration nobody
  looks at. Half the slide, or a full bleed, or don't bother.

---

## 8. Things that make a deck look machine-made

Avoid these specifically:

- **A thin accent line or bar under every title.** Extremely common in
  generated decks and in almost no professionally designed ones. Use space or
  a background change to separate the title instead.
- **Decorative stripes**: full-width header/footer bands, a vertical strip down
  one edge, a coloured bar along one side of every card. They add no
  information and read as filler.
- **Centred body text.** Centre titles and single statements; left-align
  every paragraph, bullet list, and card body. Centred paragraphs have a ragged
  left edge, which is what makes them hard to read.
- **The same layout on every slide.**
- **One beautifully styled slide and nine plain ones.** Either commit across
  the deck or keep the whole thing simple. Inconsistency is worse than plain.
- **Text shrunk to fit.** If it needs 9pt, it needs a second slide.
- **Content spilling past its box or the slide edge.** The most common
  user-visible defect and the first thing to check.

---

## 9. The build loop

Design quality comes out of the loop, not out of the first attempt:

1. **Plan** the deck's structure — the slide list and each slide's assertion —
   before creating anything.
2. **Inspect the template** (`get_presentation_info`,
   `render_slide_previews`) and map each planned slide onto a template slide.
3. **Build one slide** — `duplicate_slide` the right template slide, then fill
   it (`manage_text`, `populate_placeholder`, `add_chart`,
   `add_image_from_dial_url`).
4. **Look at it** with `visual_inspect_slides(slides=[n])`. Inspecting the one
   slide you just built is cheap and precise, and catching a layout mistake now
   is much cheaper than discovering it in all twelve copies later.
5. **Repeat** for the rest of the deck.
6. **Validate the whole deck**: `validate_presentation` for structure, then
   `visual_repair_slides` for appearance.
7. **Read the text back** with `extract_presentation_text` — check for missing
   sections, duplicated content, and template filler nobody replaced.
8. **Export** with `export_presentation` and give the user the file URL.

Your first render will have a few genuine problems: overflow, an overlap, an
alignment that drifted. Fix those and stop. Iterating past the point where the
issues are real is how the repair budget gets burned on nothing.
