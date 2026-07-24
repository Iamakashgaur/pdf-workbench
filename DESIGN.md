# Design

<!-- impeccable:design-schema 1 -->

The visual system for both surfaces: the **converter app** (`app.py`, Operate)
and the **reference page** (`index.html`, Read). One token set, one type scale,
one component language.

## World

An enterprise console in the restrained tradition — Linear's density and state
discipline, Stripe's respect for numbers, Geist's hairline monochrome, Notion's
willingness to explain itself in place. Executed at that level of finish, not
as a template wearing its clothes.

The surface is a **working ledger**. Data is the hero; chrome recedes to
hairlines and space. Nothing decorative competes with a number.

## Colour strategy

**Restrained** — neutrals plus a signal set. Correct because the visitor came
to operate, not to be persuaded.

The ground is **light by default**. The physical scene forces it: a back-office
team on Windows desktops under office light, with Excel open alongside all day.
A dark tool beside a white spreadsheet costs an eye adaptation on every glance.
Dark mode exists and is fully supported, but it is the alternate, not the default.

### Neutrals

A single cool-grey ramp carries the entire surface.

| Token | Light | Dark | Use |
|---|---|---|---|
| `--bg` | `#fbfbfa` | `#0c0c0d` | Page ground |
| `--surface` | `#ffffff` | `#141416` | Cards, panels, rows |
| `--surface-2` | `#f6f6f5` | `#1b1b1e` | Recessed areas, table headers |
| `--border` | `#e6e6e3` | `#26262a` | Hairlines — 1px, never heavier |
| `--border-strong` | `#d4d4d0` | `#35353a` | Emphasised division |
| `--text` | `#18181b` | `#f4f4f5` | Primary text |
| `--text-2` | `#52525b` | `#a1a1aa` | Secondary |
| `--text-3` | `#6b6b74` | `#9a9aa4` | Tertiary, labels |

`--text-3` carries 11px labels and table headers, so its value is set by the
4.5:1 floor on both `--surface` and `--surface-2`, not by appearance. The
lighter shades an earlier draft used failed AA at that size.

### Signals

Four, each with exactly one meaning. Their scarcity is what makes them legible.

| Token | Value | Means | Rule |
|---|---|---|---|
| `--ink` | `#18181b` / `#fafafa` | Primary action | The primary button is ink, **not blue**. Avoids the SaaS-blue default the incumbent used. |
| `--held` | `#b45309` amber | **Rows excluded from the output** | **Reserved.** Amber appears nowhere else on either surface. Its presence means exactly one thing. |
| `--ok` | `#15803d` green | Reconciled, complete | Confirmation only; never decoration. |
| `--fail` | `#b91c1c` red | Conversion failed | Errors only. |

The amber reservation is the product's central promise made visual: principle
2, *never fail silently*. If amber is ever spent on a hover state, a chart
series or an accent, that promise is broken and the rule has failed.

## Typography

System stacks, on **both** surfaces. This is an Operate/Read pair for a Windows
back office; the native face is the users' own environment, loads instantly,
and needs no network — which matters for a tool that runs from `convert.bat`
offline. No web font is loaded on either surface; a font CDN would be a network
dependency the offline tool must not have.

```
--font: "Segoe UI Variable Text", "Segoe UI", -apple-system,
        BlinkMacSystemFont, system-ui, sans-serif
--font-mono: "Cascadia Mono", Consolas, ui-monospace, "SF Mono", monospace
```

**Numerals are structural.** Every table, count, amount and metric sets
`font-variant-numeric: tabular-nums`. Money, order numbers and certificate ids
set `--font-mono`. Columns of figures must align on the decimal without effort —
this is the Stripe half of the bar and is not optional.

### Scale

A tight scale; an Operate surface earns hierarchy from weight and space, not size.

| Step | Size / line-height | Use |
|---|---|---|
| `display` | 28px / 1.2, -0.02em | Page title only |
| `title` | 18px / 1.35, -0.011em | Section headings |
| `body` | 14px / 1.55 | Default |
| `small` | 13px / 1.5 | Secondary, help text |
| `label` | 11px / 1.4, 0.04em, uppercase, 550 | Field and column labels |

Weights: 400 body, 500 emphasis, 600 headings. Never 700+ — weight that heavy
reads as shouting on a calm surface.

## Space

4px base unit; steps 4, 8, 12, 16, 24, 32, 48, 64. One rhythm throughout.
More space above a heading than below it — the heading belongs to what follows.

## Form

- **Radius:** 6px controls, 10px panels. Nothing pill-shaped; nothing square.
- **Borders:** 1px hairline, `--border`. The primary structural device.
- **Shadow:** almost never. One resting shadow for genuinely floating layers
  (`0 1px 2px rgb(0 0 0 / .04), 0 8px 24px -12px rgb(0 0 0 / .10)`). Panels sit
  on borders, not shadows.
- **Focus:** 2px `--ink` ring at 2px offset. Always visible, never removed.
- Hit targets ≥ 32px; ≥ 44px on coarse pointers.

## Motion

Motion explains a state change or it does not happen. No entrance animation on
static content, no decorative movement.

```
--dur-fast: 120ms   /* hover, focus */
--dur: 180ms        /* state change */
--dur-slow: 260ms   /* panel open */
--ease: cubic-bezier(.2, 0, 0, 1)
```

Animate `opacity` and `transform` only. Every motion rule sits behind
`prefers-reduced-motion: reduce`, which collapses durations to ~1ms rather than
removing the state change.

## Components

- **Panel** — `--surface`, 1px border, 10px radius. The unit of grouping.
- **Data table** — `--surface-2` header with `label` type; 1px row separators;
  no zebra striping (borders already do that work); numeric columns right-aligned
  and tabular; row hover is a background shift only.
- **Metric** — a `label` over a `display`-weight tabular figure. Used for
  page count, rows extracted, reconciliation.
- **Verdict** — the reconciliation panel leads with the difference at 30px
  metric scale (`$0.00` in `--ok`, or the signed shortfall in `--held`) with a
  one-line caption, above the smaller stated-vs-extracted rows. The most
  important number on the screen is the largest. This is the product's trust
  moment; it is never demoted to a ledger row.
- **Status pill** — 11px uppercase label, 1px border, tinted background from the
  signal set. Four only: settled, held, failed, pending.
- **Button** — primary is `--ink` fill with inverted text; secondary is a
  hairline border on `--surface`. One primary action per view.
- **Dropzone** — hairline dashed border, generous interior, `--surface-2` ground.
  It is a target, not a hero.

## Rules

1. **Amber is reserved for excluded rows.** Nowhere else, on either surface.
2. **Numbers get tabular figures.** Any digit a user might compare or sum.
3. **Hairlines carry structure.** Reach for a border or space before a shadow.
4. **One primary action per view.** Everything else is secondary or a link.
5. **States explain themselves.** Empty, loading, warning and error states each
   say what happened and what to do — the team shares this tool at mixed skill
   levels and has no manual.
6. **Motion only for state change**, and always reduced-motion aware.
