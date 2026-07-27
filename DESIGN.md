# Design System — Asesoría Visa Global (v4, marketing pages, 2026-07-27)

Covers the marketing/landing pages (index.html + service pages). The intake-form flow
(intake.html, ds160.html, portal.html, etc.) is a separate, unrelated design and is not
covered by this file.

## Why we changed it
The old marketing system was navy (#0E2143/#060E1C) + gold (#C8861A) with dark-glow
shadows, a decorative background grid, an infinite marquee, and Space Grotesk/Playfair
Display — the default "AI-generated fintech/legal-services" look. Roberto called the
site generic. Navy+gold is the first thing anyone pictures for "professional advisory,"
so it was replaced outright rather than cleaned up in place.

## Color Strategy: Committed
A warm, dark espresso-brown surface (not blue-black navy) carries ~70% of the interface.
Terracotta/clay carries emphasis: CTAs, active states, numerals, key words in headings.
No gold, no navy blue, no decorative glow.

## Colors
- Background: `#1C1410` — warm espresso, near-black but brown-tinted, not blue
- Background 2/3/4: `#241A14` / `#2C2018` / `#34281C` — layered warm dark surfaces
- Card surface: `#291F17`
- Terracotta (accent): `#B8501F` (deep) / `#E2703A` (bright) / `#EA8552` (text-safe on dark)
- Text primary: `#F2E9DE` (warm off-white, never pure #fff)
- Text secondary: `#B8A99A`
- Text muted: `#9C8B7A` (~5.5:1 contrast on `#1C1410`)
- Confirm green: `#6B9E5C` (earthy, not neon)
- WhatsApp CTA: `#1E9E52 → #0F6E5F` (darkened from the brand green so white text clears AA)

## Typography
- Headings: **Newsreader** (serif, editorial, warm — replaces Playfair Display)
- Body/UI: **DM Sans** (kept, works well, not on the overused-AI-font list)
- Removed: Space Grotesk (overused AI-generated-site font)

## Motion
- Ease: `cubic-bezier(.16,1,.3,1)` (ease-out-expo) everywhere. No bounce/overshoot curves.
- Removed `.wa-pulse` pulsing dot and `cta-pulse`/`pulse-gold` glow-pulse animations —
  pure decoration with no real-time data behind them.
- Marquee converted from infinite auto-scroll to a static wrapped row.

## Removed AI-slop patterns
- Decorative background grid (`.hero-bg::after`, 60px linear-gradient grid) — deleted.
- Ambient page-wide `body::before` radial-gradient glow — deleted.
- Colored blurred box-shadow glow on buttons/cards — replaced with neutral
  `rgba(0,0,0,...)` elevation shadows; accent rings kept at low opacity only.
- Bounce/overshoot transitions (`cubic-bezier(.22,.68,0,1.2)`) — replaced with ease-out-expo.

## Components
- Buttons: pill radius 100px, terracotta gradient (`#C25A20 → #9C4318`) with cream text
  (`#FBF3E9`), neutral dark elevation shadow on hover (no colored glow).
- Cards: `#291F17` surface, 1px low-opacity white border, neutral shadow on hover.
- Badges/pills: terracotta text on 8-10% terracotta-tint background.

## Contrast
Body/secondary text targets AA (4.5:1) against the espresso background; large text/labels
target 3:1. Re-verified with `npx impeccable --json index.html` after the palette swap;
a handful of small-badge edge cases remain marginally below 4.5:1 and are flagged for a
follow-up pass, not blocking.

## Pages on this system
index.html, visa-usa-ecuador.html, asesoria-visas-quito.html, asesoria-visas-guayaquil.html,
visa-espana-ecuatorianos.html, visa-canada-ecuatorianos.html, visa-rechazada-que-hacer.html,
migracion-circular-ecuador.html. `design.css` is the shared source of truth for index.html;
the service pages carry the same hex values inline (they don't currently link design.css).
