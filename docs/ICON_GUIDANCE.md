# Icon guidance

How to draw an icon as SVG for `render_svg_icon`. You write the SVG; the server
rasterizes it to a transparent PNG, has the vision model check the render, and
stores it in DIAL file storage so `add_image_from_dial_url` can place it.

---

## 1. When a drawn icon is the right answer

Icons are decoration with a job: they label a card or a step so the eye can
find it again. They do not add information. Three or four across a row of
cards, one per process step, one beside a section title — that is the whole
useful range.

Reach for a drawn icon when:

- **The surface is not white.** This is the common case and the one nothing
  else solves. A stock icon file carries its own opaque background, so on a
  tinted card or a brand-coloured panel it shows up as a pale rectangle. An
  SVG you draw has no background at all, so it sits on any colour.
- **The concept is specific to this deck** — "regulatory submission", "cold
  chain", "second-line therapy". No general icon set has it, and a vague
  substitute is worse than none.
- **The set has to be consistent.** Icons pulled from different sources have
  different stroke weights and corner treatments, and a row of them looks
  assembled rather than designed. Drawing all of a deck's icons the same way
  fixes that by construction.

Do not use one when the template already ships icons — in template mode, an
existing slide's icons are the brand's, and `duplicate_slide` brings them
along. Do not use one to illustrate a slide: that is a picture, and pictures
come from the image model through `add_image_from_dial_url`.

---

## 2. The two variants

Every icon in a deck is one of these two, and which one depends only on what it
sits on. Pick the accent colour from the deck itself — the template's own
accent, or the deck's primary colour in scratch mode. `#005DB9` below stands in
for it; substitute yours.

**Line art, for a light surface.** A circle outline with the pictogram drawn
inside it in the same colour and the same stroke weight:

```svg
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 200" width="200" height="200">
  <circle cx="100" cy="100" r="88" fill="none" stroke="#005DB9" stroke-width="4"/>
  <g fill="none" stroke="#005DB9" stroke-width="4" stroke-linecap="round" stroke-linejoin="round">
    <!-- pictogram paths here, inside a ~120x120 area centred on (100,100) -->
  </g>
</svg>
```

**White on a filled disc, for a coloured surface.** The disc is a slightly
lighter or darker shade of the panel it sits on — not the same colour, or it
disappears — and the pictogram is white:

```svg
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 200" width="200" height="200">
  <circle cx="100" cy="100" r="96" fill="#1A5A9E"/>
  <g fill="none" stroke="#FFFFFF" stroke-width="4" stroke-linecap="round" stroke-linejoin="round">
    <!-- pictogram paths here -->
  </g>
</svg>
```

Set `slide_background` on the call to the colour the icon will sit on, so the
review judges its contrast against the real surface rather than against white.

The disc is optional in both cases — a bare pictogram with no circle is a valid
icon. What is not optional is that every icon in one deck makes the same
choice.

---

## 3. Rules that keep a hand-drawn set coherent

- **One stroke weight, everywhere.** 3–4 on a 200×200 viewBox, the circle
  included. Varying it is the fastest way to make an icon look wrong without
  being able to say why.
- **Stroke, don't fill.** `fill="none"` on the pictogram; it is drawn with
  lines. The exceptions are the filled disc of the second variant and small
  solid dots used as indicators.
- **Rounded ends and corners.** `stroke-linecap="round"`,
  `stroke-linejoin="round"`.
- **Six to eight segments, no more.** The icon is displayed at about one inch.
  Detail that survives at 200px becomes a smudge there — the review renders it
  at slide size for exactly this reason.
- **Centre it.** Keep the drawing inside roughly `40..160` on both axes so the
  circle has air around it, and make it visually centred rather than
  numerically centred if the shape is lopsided.
- **No text, ever.** Not a letter, not a number. The renderer refuses text
  elements, because glyphs would come from whatever font it substitutes. A
  label goes in a textbox next to the icon.
- **No gradients, shadows, filters, or embedded images.** Flat line art only.
- **Nothing external.** No `href` to a file or a URL; the SVG must be
  self-contained. The renderer refuses these too.

---

## 4. Worked examples

Adapt these rather than inventing path data from nothing — most concepts are
one of these with a change. Each is the `<g>` contents for the line-art
variant; for the filled variant, change the stroke to `#FFFFFF`.

**Gear — settings, operations, process**

```svg
<circle cx="100" cy="100" r="20"/>
<path d="M100 60V50 M100 150V140 M140 100H150 M50 100H60 M128 72L135 65 M65 135L72 128 M128 128L135 135 M65 65L72 72"/>
```

**Lightbulb — idea, insight, innovation**

```svg
<path d="M82 125L82 115Q62 100 62 80A38 38 0 1 1 138 80Q138 100 118 115L118 125Z"/>
<line x1="82" y1="133" x2="118" y2="133"/>
<line x1="86" y1="141" x2="114" y2="141"/>
<line x1="92" y1="149" x2="108" y2="149"/>
```

**People — team, customers, stakeholders**

```svg
<circle cx="100" cy="72" r="16"/>
<path d="M68 142Q68 112 100 112Q132 112 132 142"/>
<circle cx="55" cy="82" r="12"/>
<path d="M32 142Q32 118 55 118Q66 118 71 124"/>
<circle cx="145" cy="82" r="12"/>
<path d="M168 142Q168 118 145 118Q134 118 129 124"/>
```

**Shield with a check — security, compliance, quality**

```svg
<path d="M100 48L142 68V108Q142 142 100 160Q58 142 58 108V68Z"/>
<polyline points="82,105 95,118 120,88"/>
```

**Target — goal, focus, objective**

```svg
<circle cx="100" cy="100" r="45"/>
<circle cx="100" cy="100" r="28"/>
<circle cx="100" cy="100" r="11"/>
```

**Clock — time, schedule, speed**

```svg
<circle cx="100" cy="100" r="45"/>
<polyline points="100,70 100,100 125,110"/>
```

**Rocket — launch, growth, acceleration**

```svg
<path d="M100 50Q82 78 82 115L100 135L118 115Q118 78 100 50Z"/>
<path d="M82 108L62 125"/>
<path d="M118 108L138 125"/>
<line x1="92" y1="135" x2="92" y2="152"/>
<line x1="108" y1="135" x2="108" y2="152"/>
```

**Handshake — partnership, agreement**

```svg
<path d="M50 105L70 85L95 85L105 75"/>
<path d="M150 105L130 85L105 85L95 75"/>
<path d="M70 105L90 120L110 105L130 120"/>
```

**Leaf — sustainability, growth, natural**

```svg
<path d="M100 145C100 145 55 125 55 80Q55 45 100 45Q145 45 145 80C145 125 100 145 100 145Z"/>
<path d="M100 145Q100 105 80 75"/>
<path d="M100 120Q110 100 125 88"/>
```

**Document with lines — report, policy, submission**

```svg
<path d="M70 48H115L135 68V152H70Z"/>
<polyline points="115,48 115,68 135,68"/>
<line x1="85" y1="92" x2="120" y2="92"/>
<line x1="85" y1="110" x2="120" y2="110"/>
<line x1="85" y1="128" x2="106" y2="128"/>
```

**Upward bar chart — growth, performance, results**

```svg
<polyline points="55,148 145,148"/>
<line x1="75" y1="148" x2="75" y2="118"/>
<line x1="100" y1="148" x2="100" y2="96"/>
<line x1="125" y1="148" x2="125" y2="70"/>
```

---

## 5. Using it

1. `render_svg_icon(svg=..., concept="cold chain", slide_background="#005DB9")`
   — `concept` is what the reviewer compares the drawing against, so name the
   thing, not the shapes.
2. Read the `review` in the response. `passed: false` means no icon was
   stored: fix what the issues name and call again. Two failed attempts means
   the drawing is too ambitious — cut it to three or four large shapes, or drop
   the icon rather than shipping a broken one.
3. Place the returned `image_url` with `add_image_from_dial_url`, in a square
   box (`width == height`), `fit="contain"`. Roughly 0.7in for an icon in a
   card, 1.0–1.2in for one that leads a section.
4. Reuse the same `image_url` wherever that icon repeats. Rendering it again
   costs a vision call and risks getting a subtly different drawing back.

Draw a deck's icons as one batch, in one style, before placing any of them.
That is how the set stays consistent — and it is the same reason the template's
own icons should win whenever they exist.
