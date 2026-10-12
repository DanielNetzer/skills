# A component family, not just a glass card

Start with a shared material: tint, dark ink, white upper rim, soft lower rim, contact shadow and radius. Keep labels on a stable foreground. Test a button, checkbox, switch, slider and segmented choice with the same tokens before decorating a full screen.

## Map shape to behavior

| Component | Construction | Behavior and checks |
| --- | --- | --- |
| Button | Thick tinted pill, upper highlight, shallow contact shadow | Native button; immediate press feedback; click commits; release outside cancels. Icon and label reflect state. |
| Checkbox chip | Thin translucent pill, checkmark when selected | Native checkbox inside a label. Selection is not color alone. Keep visible focus on the chip when the visually hidden input is focused. |
| Favorite | Isolated coral pebble with a stable heart icon | Toggle button with aria-pressed and a descriptive accessible name. Preserve at least a 44px target. |
| Switch | Pale capsule track with a smaller glass thumb | Native checkbox with role=switch. Update its accessible checked state; move only the thumb, never the label. |
| Slider | Thin inset track, tinted fill, taller soft thumb | Native range. Bind the visible number with output/htmlFor. Arrow keys, Home and End work without custom handlers. Thumb shape can compress while pressed; its hit target must stay predictable. |
| Segmented choice | Shared glass trough, tinted selected pill | Native radio group with a legend. Use one name for the group; arrow keys move selection. |
| Player | Frosted panel, colored play control, thin wave track | Real play/pause/seek state. If audio is absent, explicitly call it a silent preview. Pause hidden-page work and clean up timers. |

## Native range with a soft thumb

```html
<label for="light">Light level</label>
<output for="light">42%</output>
<input id="light" type="range" min="0" max="100" value="42" />
```

```css
input[type=range] {
  appearance: none;
  width: 100%; height: 18px; margin: 16px 0;
  border: 1px solid white; border-radius: 20px;
  background: linear-gradient(to right, var(--tint) 0 var(--fill), #eef2ee var(--fill) 100%);
  box-shadow: inset 0 2px 5px #648d7a20;
  cursor: pointer;
}
input[type=range]::-webkit-slider-thumb {
  appearance: none; width: 31px; height: 37px;
  border-radius: 48% 52% 45% 55%; border: 1px solid white;
  background: linear-gradient(125deg, #effbf4, var(--tint));
  box-shadow: inset 2px 2px 3px white, 0 2px 5px #52746525;
}
input[type=range]::-moz-range-thumb {
  width: 31px; height: 37px; border-radius: 48%;
  background: var(--tint); border: 1px solid white;
}
input[type=range]:focus-visible { outline: 3px solid #294b40; outline-offset: 5px; }
```

Update the number and `--fill` from the same state. Do not use visual fill as the only value representation. Allow vertical scrolling on touch; do not capture every pointer gesture for decoration.

## Context and comparison

A quiet neutral board reveals edges and shape. A detailed wallpaper reveals blur and transparency. Show both, and keep a plain-surface option with the same controls, text, state and hit areas. A wallpaper is not permission to sacrifice contrast: inspect its darkest and brightest patches, and increase panel opacity as needed. Do not make essential labels inherit decorative transparency.

Use original or licensed artwork. Do not claim that a recreated platform concept is an official implementation. Keep simulated device settings local and identify them as a demo.
