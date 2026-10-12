# Material recipes

These are starting values, not universal rules or a reproduction of the video's renderer.

## A readable glass-like panel

Keep foreground content separate from the decorative surface. Blur is optional; opacity provides the fallback.

```css
.material {
  color: #173630;
  background: #e2eee8;
  border: 1px solid #90aca0;
  border-radius: 1.5rem;
  box-shadow: 0 12px 28px -20px #17363080, inset 0 1px 0 #fff;
}
@supports (backdrop-filter: blur(12px)) {
  .material { background: rgb(240 248 242 / 94%); backdrop-filter: blur(12px); }
}
.material button { min-height: 44px; }
.material button:focus-visible { outline: 3px solid #173630; outline-offset: 4px; }
.material button:active { transform: translateY(1px); }
@media (prefers-reduced-motion: reduce) {
  .material, .material * { animation: none; transition: none; }
}
@media (forced-colors: active) {
  .material { background: Canvas; color: CanvasText; border: 1px solid CanvasText; box-shadow: none; }
}
```

For decorative pebbles, combine an irregular rounded shape with a radial highlight, a tint gradient and an inset rim. Keep those shapes out of the text plane and mark them aria-hidden. This is an illustration of material, not a control or real refraction.

## Motion decisions

| Interaction | Starting choice | Verify |
| --- | --- | --- |
| Press | Immediate 1px displacement, no input delay | Release outside cancels normally |
| Selection | Short background/indicator transition | Text and selected state update immediately |
| Repeated save | Immediate label/icon change | No decorative delay, no duplicated action |
| Drag | Pointer tracking, then velocity-aware settle | Re-grab, cancel and keyboard alternative |
| Decorative ambient motion | Off by default | Adds meaning before enabling; pause and reduced-motion support |

Use a frame trace before claiming smoothness. Watch composite layers, large blurred areas and painting costs; do not add permanent will-change everywhere. Avoid React state updates on every pointer frame. With WebGPU, handle device loss and unavailable adapters, stop work offscreen, and keep controls in semantic DOM.

## Review prompts

- Which visual detail shows what can be operated?
- Does the plain version complete exactly the same task?
- Can a person read the value with the most difficult backdrop?
- Is the quiet version complete, or did it lose information?
- Which device and interaction support the performance claim?
