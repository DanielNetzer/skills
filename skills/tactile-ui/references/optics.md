# When the reference actually bends light

CSS backdrop blur softens what is behind a panel. It does not magnify letters or bend a grid. If those are the defining details, use an optical approximation and explain its boundary.

The included `examples/glass.wgsl` is an original height-field lens used in the live Tactile UI study. It is not a physically accurate ray tracer or an extraction of the reference video. It samples a supplied background texture; it does not capture arbitrary DOM content.

## Construction

1. Make a signed distance field for a rounded capsule or an organic radial contour.
2. Estimate its outward gradient with central differences.
3. Scale local coordinates toward the lens center for broad-face magnification; add a curved offset near the rim.
4. Sample the background at that displaced coordinate. Slightly separate RGB samples near the edge for restrained dispersion.
5. Add a directional upper rim, a darker opposite edge and a contact shadow. Use one light direction.
6. Keep controls and essential text outside the shader in semantic DOM.

## Shader contract

The fragment shader expects normalized, top-origin UVs. It exposes `params` (time in seconds, aspect ratio, position 0–1, depth 0–1, shape 0=capsule/1=organic, material index 0–4), `src` (2D background texture) and `samp` (linear sampler). Supply your own background; the live study renders its own “Stay curious.” text to a canvas texture. The grid is procedural.

The example is compatible with the installed vgpu 0.5.0 effect workflow. Read the documentation for the version in the target project rather than assuming this recipe is a stable API specification. Validate the WGSL, compile its color-target signature, draw and await GPU validation before declaring it ready.

## Runtime boundary

Load only when needed. Cap pixel ratio and animation frequency; dispose the device, textures, observers and canvas on unmount. Stop the loop offscreen and while the document is hidden. Honor reduced motion at startup and if it changes: keep a static shader frame and let native controls redraw on demand. Supply a labeled static fallback for unavailable WebGPU, initialization errors and device loss. Redraw static surfaces after resizing without recursively acquiring a frame inside the surface's resize callback.

Verify rendering visually, not just shader compilation. Move the lens across a straight line or letters; inspect the inside and outside at the same time. Check edge clipping, mobile aspect ratios and the failure path. A successful build or one desktop preview is not a mobile GPU performance benchmark.
