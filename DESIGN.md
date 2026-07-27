# Design System — Asesoría Visa Global (v5, marketing pages, 2026-07-27)

Covers the marketing/landing pages (index.html + service pages). The intake-form flow
(intake.html, ds160.html, portal.html, etc.) is a separate, unrelated design and is not
covered by this file.

## Why we changed it (again)
v4 (espresso-brown `#1C1410` + terracotta `#B8501F`) was rejected by Roberto: too warm/tierra,
doesn't read as "legal/immigration advisory" — he wants confidence, not a cafe. Full rebuild
of the palette to a light, professional theme: petrol (deep blue-green) as the anchor,
warm bone (not pure white) as the field, emerald as the one live accent color for action.

## Color Strategy: Committed
Bone carries the field (~70%), deep petrol carries text/headings/dark chrome (nav-on-scroll,
high-contrast emphasis), emerald is reserved for CTAs and active/success states only. No gold,
no terracotta, no navy-black.

## Colors
- Background: `#F7F2E7` — warm bone, not pure white
- Background 2/3/4: `#F0E9DA` / `#E8DFCC` / `#DED2B8` — layered warm-bone surfaces
- Card surface: `#FFFFFF` (clean white card lifts off the bone field)
- Petrol (text/heading anchor): `#16332F` (near-black petrol, headings/emphasis) /
  `#3F5B57` (body text) / `#5A7873` (muted/secondary)
- Emerald (accent — CTAs, buttons, active states): `#0E8F63` → `#0A6B4A` (button gradient),
  `#12B37E` / `#0E9A6C` (bright text-safe accent). Chosen over coral/terracotta-adjacent tones
  because Roberto explicitly rejected anything reading as "warm/earth"; emerald is vivid,
  reads as growth/approval/"go," and has zero overlap with the rejected terracotta family.
- Dark chrome (nav-on-scroll only): `rgba(14,36,34,.95)` petrol-black, not pure black
- WhatsApp CTA: kept its own brand green (`#157A40`/`#0F6E5F`, `#6B9E5C`) — untouched,
  brand recognition matters more than palette purity for that one button
- Borders/hairlines: `rgba(22,51,47,.07–.13)` — dark petrol at low opacity on the light field

## Typography
Unchanged from v4 — **Newsreader** (serif headings) + **DM Sans** (body/UI). Both already
avoid the overused-AI-font list (Space Grotesk/Inter/Playfair/Roboto/Geist/Plus Jakarta Sans).

## Positioning / copy angle change
v4 spoke in a corporate "nosotros." Roberto's ask: make himself the visible guarantee, since
sales close on his personal presence, not an anonymous agency. Hero and key CTAs now speak as
Roberto directly ("Yo mismo reviso cada caso," "Cuando hablas con nosotros, hablas conmigo")
while keeping the core promise "Tú solo cuéntanos, nosotros hacemos todo." No invented bio
details (no fabricated years/numbers) — only the point-of-view shifted from agency to person.

## Motion
Unchanged from v4 — ease-out-expo everywhere, no bounce/glow/marquee-autoscroll.

## Components
- Buttons: pill radius 100px, emerald gradient (`#0E8F63 → #0A6B4A` or `#12B37E → #0C7A54`)
  with white text, neutral dark elevation shadow on hover (no colored glow).
- Cards: `#FFFFFF` surface on bone field, 1px low-opacity petrol border, neutral shadow on hover.
- Badges/pills: emerald text on 8-10% emerald-tint background (`rgba(18,179,126,.08-.22)`).

## Contrast
Body text `#3F5B57` on `#F7F2E7` and headings `#16332F` both clear AA comfortably (light-on-dark
flipped to dark-on-light preserves the same relative contrast steps as v4). Re-verify with
`npx impeccable --json index.html` after any further copy/color edits.

## Pages on this system
index.html, visa-usa-ecuador.html, asesoria-visas-quito.html, asesoria-visas-guayaquil.html,
visa-espana-ecuatorianos.html, visa-canada-ecuatorianos.html, visa-rechazada-que-hacer.html,
migracion-circular-ecuador.html. `design.css` is the shared source of truth for index.html;
the service pages carry the same hex values inline (they don't currently link design.css).
