# Doc Haus interface system

Doc Haus uses Material 3's semantic roles, hierarchy, shape contrast and adaptive
containment in native React and CSS. This is a professional web workspace: the
canvas and documents stay white, navigation is lightly toned, and dense legal
work keeps its space. There is no Material component framework or new runtime
dependency. Space Grotesk identifies page headings; system sans serves controls
and long reading; document typography continues to respect the document.

## Source and cascade

`src/styles/design-system.css` owns all global tokens. `primitives.css` owns
native buttons, form controls, row actions and switches. `styles.css` contains
remaining legacy component structure; its old global tokens, buttons/forms,
settings, review and issue rules have been extracted to their owning sheets.

`main.tsx` loads legacy structure, primitives and motion, then design foundations,
shell, documents, conversation, library, review, settings and Matters. The last
sheet specialises the shared library layout for Matters. Component imports load
earlier, so late design sheets own final composition. Update an owning rule
instead of appending a competing copy. Use shallow component selectors; avoid
`!important` except the reduced-motion accessibility preference.

## Color mapping

The initial palette was derived using Material Color Utilities 0.4.0:

```ts
const scheme = new SchemeTonalSpot(Hct.fromInt(argbFromHex('#17634F')), false, 0)
```

Generation is a design-time aid. The shipped tokens are static CSS. Generated
primary tone 40 (`#146b55`) remains a reference token; the action/brand role is
intentionally the exact Doc Haus `#17634F`. Container and neutral roles are
curated lighter than the generated scheme to preserve a predominantly white UI.

| Material role (`--md-sys-color-…`) | Value / purpose | Existing alias |
| --- | --- | --- |
| primary / on-primary | #17634F / white; primary actions | --accent / --text-on-accent |
| primary-container / on-primary-container | #EDF4EE / #154C3B; selected controls | --accent-soft |
| surface / on-surface | white / #242C28; work surfaces and ink | --panel / --text |
| surface-container-low | #F8FAF7; navigation/supporting areas | --bg / --panel-hover |
| surface-container / high / highest | #F2F5F1 / #EDF1EC / #E3E9E3 | --bg-sunken |
| on-surface-variant | #526057; secondary reading | --text-secondary |
| outline | #707974; controls needing a visible boundary | --border-strong |
| outline-variant | #DCE4DC; separators and containment | --border |
| error / error-container | #A3302D / #FCEFED | --danger / --danger-soft |
| warning / warning-container | #805311 / #FFF5E3 | --warn / --warn-soft |
| info / info-container | #315E77 / #EDF5FA | --info / --info-soft |

Measured examples: white on brand green is 7.16:1; main ink on white is
14.33:1; secondary text on the low surface is 6.31:1; muted text on the selected
container is 4.87:1; outline on white is 4.49:1. Text pairs must meet 4.5:1,
with 3:1 for relevant non-text boundaries. Decorative
separators may be quieter. Use state labels/icons as well as color. Primary,
selected, warning and error text always use the corresponding paired role.
`--md-sys-state-*` provides hover, pressed, selected, focus and disabled roles.

## Typography, shape and spacing

`--md-sys-typescale-*` are complete CSS `font` shorthand tokens. Display small
and headline large/medium/small use Space Grotesk; title, body and label roles
use system sans. Body medium is 14px with 1.6 line-height. Headline large is
32px, reduced to 28px in smaller page layouts. Workspace headings stay compact.
The existing compact/standard/large appearance preference still scales the
whole UI; no new stored setting or persistence semantics were introduced.
`--ui-scale` tracks that existing CSS zoom. `--viewport-inline` and
`--viewport-block` compensate viewport units so fixed workspaces, dialogs and
composers fit at every preference size.

`--md-sys-shape-corner-*` maps extra-small/small/medium/large/extra-large/full
to 4/8/12/20/28/999px. Actions and selected navigation are rounded; fields and
rows use small/medium; standalone panels and dialogs use large/extra-large.
Tables stay precise. `--radius-*` aliases preserve existing consumers.

Corners: one radius per role.

- Overlays (dialogs, sheets, picker panels, document reader, review modal,
  matter edit, settings panel): `--md-sys-shape-corner-extra-large` (28px).
- Cards, section containers and card-style notices: `--radius-lg` (20px).
- Inputs, selects, textareas, list rows, menus, popovers, toasts:
  `--radius-md` (12px).
- Chips, tooltips and settings segments: `--radius-sm` (8px; segments press to 4px).
- Text buttons are pills; icon and close buttons are circular.
- The composer card keeps its own radius; the review table and document page stay square.
- Nested corners are concentric: inner = max(4px, outer - inset), written with
  `calc()` and the tokens (settings tabs: 28px - `--space-3`; menu items: 12px - 6px).

`--space-1` through `--space-8` are 4, 8, 12, 16, 20, 24, 32 and 40px.
Use 24–40px between page sections and 8–16px inside controls. Supporting
containers have restrained tone/outline, not stacked elevation. The elevation
0–3 tokens reserve shadows for menus and overlays.

## Native component contracts

- Buttons: `.primary`, `.ghost`, `.danger`; `.btn-sm`
  and `.icon-btn` are compact variants. Keep a textual accessible name and use
  `disabled` / `aria-busy` for unavailable or pending work. Only one primary
  action should dominate each task area.
- Fields: native labels, hints via `aria-describedby`, and `aria-invalid` on
  validation errors. Never use a placeholder as the only visible label.
- Selection: native checkboxes/radios and `.switch[role=switch]`; settings
  segments use `aria-pressed`; settings tabs use a roving tab stop and arrow keys.
- Modals: prefer native `dialog.showModal()`. Existing div picker/viewer shells
  share `useModalFocus` for initial focus, Tab containment and opener restoration.
  `ModalHeader` supplies consistent title/close treatment. Nested dialogs own
  focus independently. Desktop document panels remain nonmodal.
- `Tooltip` supports hover (including moving onto its text), keyboard focus, Escape
  and viewport-aware placement.
  Tooltip text supplements an accessible name. Toasts retain status announcements,
  distinct icons and explicit dismiss actions.
- Lists expose actions without hover. Review tables retain columns and intentional
  horizontal scrolling; reviewed/stale states include words, not color alone.
- Loading, error and empty states preserve the task context and an appropriate
  retry/add action. Legal approvals and drafting gates keep their existing APIs.

## Responsive and motion rules

Adapt to available content width, not a stretched phone scaffold. Matters uses a
container query to move the supporting form above the index with a disclosure.
Global navigation becomes a full-width web header with a six-link disclosure on
phones; the saved rail preference also controls that disclosure. Library panes
stack in narrow windows; active desktop workspaces retain independent
scrolling and an accessible composer. Documents use an adjacent panel on wide
screens, an overlay at intermediate sizes and a full-screen modal on phones.
Comparison documents and change lists stack at tablet widths to preserve reading
space. Complex review grids scroll horizontally rather than becoming unreadable cards.

Motion tokens are 120/200/320ms with standard/emphasized easing. Prefer feedback
through color, opacity and short spatial transitions; never animate document
reading. All component animations/transitions respect `prefers-reduced-motion`.
Overlays move on M3 springs (`--spring-spatial-*` 230/320/490ms with a slight
overshoot for position, `--spring-effects-*` 150/240/330ms without one for
opacity and color; each easing has a `-ms` duration). They enter on the default
springs and exit on `--spring-effects-fast`, as transitions, so a reopen during
an exit turns back from where it is. `useExit` (or `<Exit when>`) keeps a
closing overlay mounted for that exit and sets `Closing` for it; focus returns
and the page is live as the exit starts (native dialogs `close()` at once and
`overlay` keeps them painted), so only the picture lingers. Reduced motion
unmounts at once.
Primary touch controls and text fields provide 44px targets; browser zoom and
text preferences must not hide essential actions.

## Verification

From `apps/web`, run `bun run typecheck`, `bun run build`, and `bun test test tests`.
Existing `e2e/README.md` documents isolated fixture-based browser runners.
`e2e/material-expressive.ts` captures representative routes at mobile, tablet and
desktop sizes and can collect accessibility results using an externally supplied
axe script. Fixtures block unexpected backend traffic; never use paid generation
for visual verification. Capture artifacts outside public source.

References: [Material 3](https://m3.material.io/),
[color roles](https://m3.material.io/styles/color/roles),
[Material skill](https://github.com/hamen/material-3-skill),
[color utilities](https://github.com/material-foundation/material-color-utilities),
[Material Web limitations](https://github.com/material-components/material-web).
