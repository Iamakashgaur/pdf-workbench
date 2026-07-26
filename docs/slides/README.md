# Explainer slides

Three 1200×1200 slides that explain what this project does and why the approach
is what it is — written for a general audience rather than a code reader.

| | |
|---|---|
| [`1-problem.png`](1-problem.png) | The flat text a converter actually sees, and the four real cases that break pattern-matching |
| [`2-insight.png`](2-insight.png) | Why column boundaries are calibrated from the widest empty strip instead of the midpoint between header labels |
| [`3-result.png`](3-result.png) | The reconciliation panel, and what happens to a row that cannot be read |

Plus one that is not part of the carousel:

| | |
|---|---|
| [`column-calibration.png`](column-calibration.png) | Slide 2's diagram without the carousel framing, landscape — embedded in the project README where the calibration is explained |

## What is real and what is drawn

Slide 3 embeds [`panel.png`](panel.png), a **real screenshot** of the running app
— captured from the live DOM element, not recreated. Its data comes from a
synthetic invoice, because real supplier data cannot go in a public repository;
the slide says so on its face.

Slides 1 and 2 are drawn diagrams. The geometry in slide 2 is the actual
relationship, not a sketch: header labels centred, data left-aligned, and a long
value genuinely overflowing past the midpoint — which is the bug the calibration
in [`ReportColumnParser._boundaries`](../../pdf_to_excel.py) exists to avoid.

## Regenerating them

The `.html` files are the source. They use the same design tokens as the app, so
the slides and the interface stay one visual system.

```bash
pip install playwright          # not a project dependency; tooling only
python docs/slides/shoot.py docs/slides
```

It drives the system Chrome via `channel="chrome"`, so no browser download is
needed. Edit a `.html`, re-run, and the `.png` beside it is replaced.

`panel.png` is captured separately from the running app rather than by this
script — see the capture note in the project history if it needs refreshing.
