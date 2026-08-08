# H-Queex Hub — UI Standards

This document is the standing set of UI rules for `templates/index.html`. It exists so
basic UX quality doesn't depend on being re-requested every time.

**Process requirement:**
- Before making any UI change: check this document.
- After making any UI change: verify the change against every rule below.
- If a design decision requires deviating from a rule here, flag the deviation explicitly
  in your response and explain why — never deviate silently.

## Brand palette (reference)

| Token | Hex | Use |
|---|---|---|
| Navy | `#16294A` | Primary brand colour — headers, primary buttons, nav |
| Gold | `#B08D57` | Accent — primary button border, active/on states |
| Steel blue | `#618096` | Muted/secondary text, hint text, optional field labels |
| Charcoal | `#37373A` | Primary body text |
| Dark grey | `#535356` | Secondary/muted body text |
| Border | `#E2E2E3` | Field and card borders |
| Bg light | `#F7F7F7` | Input backgrounds, alternating table rows |
| Bg deep | `#F2F2F2` | Disabled/readonly backgrounds |
| White | `#FFFFFF` | Base background, text on navy/gold |
| Error/red | `#C62828` | Errors, destructive actions, non-deductible/blocked states |

Font: Segoe UI throughout. No other typeface, no other colours.

## Form fields

- Placeholder text must always be shorter than the field width — **max 25 characters**
  for standard fields, **max 40** for wide/full-row fields.
- Placeholder text is never a sentence describing what the field does — **one or two
  words only** (e.g. `"Search suppliers..."`, not `"Search or type a new supplier /
  payee"`).
- Every input must be wide enough to display its expected content without horizontal
  scrolling (e.g. calculated currency fields sized for at least 8 characters including
  `€` and two decimals; free-text fields like Supplier / Payee spanning the full row).
- Labels sit above the field. Never rely on placeholder text as the only label.
- Mandatory fields show a visible asterisk (`.field.required label::after`).
- Locked/read-only fields must never look editable: no dropdown arrow, no input border —
  render as a badge or plain text instead.
- **Description fields must always be textareas with a minimum of 3 rows, never
  single-line inputs.** Any field labelled Description, Notes, or Details anywhere in
  the app must follow this rule — full width, resizable vertically, same font size and
  Segoe UI as every other field, with comfortable padding for reading and editing.

## Dropdowns

- Primary options (most likely to be picked) are visually prominent — regular weight,
  full-size, standard text colour.
- Secondary actions (add new, exceptions) are visually subdued — smaller text, steel
  blue `#618096`, separated from primary options by a divider line.
- Maximum 5–6 visible options before the list scrolls internally.
- Never show a dropdown arrow on a field the user cannot interact with.

## Warning messages

- Tell the user exactly what to do next, not just what the problem is.
- Never use a "Learn more" control that replaces the current message in place — if more
  detail is needed, use a proper modal or an expand-below, never an in-place swap.
- Severity meaning is fixed:
  - **Red** = blocked or non-deductible — requires action or acknowledgement.
  - **Amber** = needs attention but the user can proceed.
  - **Green/info** = informational only.
- Plain English only, no accounting/tax jargon without explanation.

## Buttons

- Primary action: navy `#16294A` background, white text, gold `#B08D57` bottom border.
- Secondary/cancel: white background, navy `#16294A` text and border.
- Destructive action: red `#C62828`, always requires confirmation before firing.
- Labels describe the outcome, not the mechanical action — `"Confirm and Save"` not
  `"Submit"`, `"Go Back to Form"` not `"Cancel"`, `"I Understand — Save Record"` not
  `"Save Anyway"`.
- Never use ALL CAPS on a button label or any interactive element.
- **Destructive actions must never appear as prominent buttons in list views — they
  belong inside the edit context only.** Archive/Cancel is never a button on a table row
  or a list card. Edit is the sole row-level action everywhere (Services, Clients,
  Suppliers, Income, Expenses, Invoices, Payroll, Subscriptions). The archive/cancel
  action lives at the bottom of the record's edit form or edit modal, below a full-width
  divider, styled as a small red `#C62828` text link ("Archive this record" — "Cancel
  invoice" for Invoices) using the shared `archive_in_modal()` macro (see Component
  Patterns) so every instance across the app is identical.
- Archive/cancel confirmation is an inline step below the link ("Are you sure? This will
  archive the record. Yes, archive / Cancel"), never a full modal on top of a modal and
  never a floating tooltip.
- If archiving a record is blocked by a dependency (e.g. an Income entry linked to a paid
  invoice), the archive link does not render at all — show a plain explanatory line
  instead (steel blue `#618096`) telling the user where to make the change.

## Breadcrumbs

- Every page reached from a parent section (e.g. a Finance module reached from the
  Finance hub) shows a breadcrumb above its page title: small steel blue `#618096` text,
  format `Parent › Current Page`, with the large navy bold page title directly beneath
  it. This is how a user knows where they are and how to get back — never rely on the
  main nav alone to convey hierarchy.

## Layout

- Fields in the same logical group are visually grouped together.
- Related fields sharing a row have equal visual weight unless one is clearly more
  important.
- No element may be truncated or overflow its container on a standard 1280px-wide
  screen.
- Binary yes/no options use a toggle switch, not a checkbox styled to look like a legal
  disclaimer.

## Typography

- Never use ALL CAPS for user-facing labels. Title Case for labels, sentence case for
  hint text and messages.
- Hint text is always smaller than label text and always steel blue `#618096`.
- Error messages are always red `#C62828` — never orange or yellow.

## Consistency

- Every form follows the same field order convention: Date → Title/Name → Description →
  Category → Amounts → secondary fields.
- Every list/table uses the same column styling: header in navy, alternating rows in
  white and `#F7F7F7`. Numeric columns (price, amount) right-aligned; status columns
  centred; name/description columns left-aligned; action columns (Edit/Archive)
  right-aligned — consistently, in every table, not decided per-page.
- Every page uses the same header structure: module title in navy, subtitle in steel
  blue.
- **The same interactive component must never be hand-rolled twice.** The in-modal
  archive/cancel action, Edit button, and the supplier/client smart-search field are each
  defined exactly once as a Jinja2 macro at the top of `templates/index.html`
  (`archive_in_modal()`, `edit_button()`, `supplier_search_field()`) and called everywhere
  they're needed. If a page needs a
  variant (different label, different confirm text), extend the macro's parameters —
  never copy-paste the markup and tweak it in place. A second hand-written copy of an
  existing component is a bug, not a new feature.
- **System status information never belongs in the main KPI/dashboard area.** Backup
  status, server/sync status, and similar operational information belong in a dedicated
  status bar or footer, in small steel blue `#618096` text — not as a KPI card competing
  with business metrics like Income or Net Cashflow.

## Component Patterns

General rule: **before building any interactive UI component, identify the standard,
universally understood pattern for that component type and follow it exactly.** Do not
invent a custom layout for something that already has a convention users recognise on
sight.

### Autocomplete / search fields

- The search results dropdown must always appear immediately below the input field with
  **zero gap**, so input and dropdown read as a single unified component.
- The dropdown must use `position: absolute` so it floats above the form layout and
  never pushes other elements down.
- The input field's border becomes gold `#B08D57` while the dropdown is open.
- The dropdown connects seamlessly to the bottom of the input field — no visible
  separation or double border between them (drop the dropdown's top border/radius so it
  reads as a continuation of the field, not a separate box).
- Secondary options or toggles related to the field always appear **below the closed
  dropdown**, never between the input and the results.
- Maximum 5 results visible before scrolling.
- This is a standard browser autocomplete pattern — follow it exactly, do not invent a
  new layout.
