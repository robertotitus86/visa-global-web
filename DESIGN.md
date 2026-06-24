# Design System — Asesoría Visa Global

## Color Strategy: Committed
The dark surface carries 70% — gold is reserved for active states, CTAs and key accents only.

## Colors
- Background: oklch(8% 0.02 250) — near-black, blue-tinted
- Surface: oklch(14% 0.025 250)
- Gold active: oklch(72% 0.16 85) — #F5B429
- Gold dim: oklch(72% 0.16 85 / 15%)
- Text primary: oklch(92% 0.01 250)
- Text secondary: oklch(60% 0.015 250)
- Green confirm: oklch(62% 0.18 155)

## Typography
- Font: Inter (system-ui fallback)
- Question size: clamp(28px, 6vw, 52px) — large and confident
- Body: 15–16px
- Labels: 11px uppercase, 0.08em tracking
- Line height: 1.3 for headings, 1.65 for body

## Motion
- Ease: cubic-bezier(0.16, 1, 0.3, 1) — ease-out-expo
- Question entry: translateY(40px) → 0, opacity 0→1, 400ms
- Question exit: translateY(-20px), opacity 1→0, 250ms
- No bounce, no spring

## Components
- Option buttons: full-width pill, 56px min-height, large font
- Text inputs: borderless bottom-only or full border minimal
- Progress: 3px line at top, gold fill
- CTA: full-width at bottom, 56px, gold background
