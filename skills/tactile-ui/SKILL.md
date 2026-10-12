---
name: tactile-ui
description: Design or review tactile web interfaces with translucent materials, rounded forms and responsive interaction. Use when implementing glass-like surfaces, material palettes or physical-feeling controls while preserving legibility, accessibility and performance.
license: MIT
metadata:
  author: Daniel Netzer
  version: "0.2.0"
---

# Tactile UI

Make the surface inviting and the action obvious. Material is a way to communicate depth and affordance, not a reason to distort the task.

## Start with the task

Identify the action, its frequency and the existing design system. Keep familiar semantics, information order and hit areas. Choose one expressive surface first. A checkout, editor or dense table may need almost no material treatment.

Read [material recipes](references/materials.md), [component recipes](references/components.md) and, when needed, [optics](references/optics.md) for implementation and [sources](references/sources.md) for provenance. The Sea Glass clip is a visual reference, not evidence of an Android release or usable production code.

## Read the whole reference

Inspect the full reference before choosing an implementation. Record the material, geometry, interaction and state changes separately. A clip may move from abstract glass to actual controls; do not deliver decorative blobs when the useful reference is a slider, checkbox or phone interface. Build a small component board first, then show the same tokens in context. Identify what is observed versus your own interpretation.

## 1. Separate the readable layer

Put labels, values, errors and focus rings on a stable, sufficiently opaque plane. Place tint, highlights and decorative texture underneath. Test contrast against the worst visible backdrop, including overlapping layers; an attractive screenshot alone proves nothing.

## 2. Give the surface a construction

Use a tinted base, restrained edge highlight and soft contact shadow. Keep one light direction across related objects. Distinguish translucent blur from optical refraction: CSS backdrop blur does not bend the scene. Label approximations honestly. Start with CSS; add a shader only when actual distortion is needed and a measured device budget allows it.

## 3. Keep shape and target distinct

Organic outlines suit decorative specimens and isolated accents. Keep input geometry and clickable regions predictable. Never stretch text with a squashing parent. Color may identify a material, but selection also needs a label, border or checkmark.

## 4. Let interaction set the motion

Give press feedback immediately; commit native button actions on click, allowing cancellation. Animate only the properties that explain a state change. A frequently repeated action should not incur a flourish or wait. For a draggable object preserve the grab offset, capture the pointer, and handle cancellation. Retarget from current position and velocity; do not disable input until an animation ends. Do not call an eased CSS transition a physical spring.

## 5. Provide a quiet version

Honor reduced motion, including changes while the page is open. Remove decorative movement and large spatial transitions; retain visible state feedback. Supply keyboard and touch equivalents for functional interactions; decorative pointer effects can be omitted entirely. Offer an opaque fallback without backdrop-filter and a readable forced-colors view. Essential content must never depend on an effect loading.

## Deliver and verify

Deliver a working interaction, reusable tokens, the unembellished comparison and a short report of checks actually performed. Exercise keyboard activation, focus, touch, cancellation where relevant, rapid reversal, narrow screens, zoom and reduced motion. Check the worst backdrop and profile the interaction on a stated device before claiming performance. Compare with the plain baseline under the same conditions; agree a frame and input-latency budget for the target hardware, then remove effects that exceed it. Do not invent a usability lift, attribution, original creator or benchmark result.

If the effect competes with reading or slows the action, simplify the effect. Preserve the user's scope; a polish request is not permission to redesign the entire product.
