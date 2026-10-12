struct Params { time: f32, aspect: f32, position: f32, depth: f32, shape: f32, material: f32 }
@group(0) @binding(0) var<uniform> params: Params;
@group(0) @binding(1) var src: texture_2d<f32>;
@group(0) @binding(2) var samp: sampler;
fn tint() -> vec3f {
  if (params.material < 0.5) { return vec3f(.62,.83,.74); }
  if (params.material < 1.5) { return vec3f(.48,.61,.91); }
  if (params.material < 2.5) { return vec3f(.95,.74,.32); }
  if (params.material < 3.5) { return vec3f(.86,.49,.44); }
  return vec3f(.76,.82,.84);
}
fn distanceToGlass(p: vec2f) -> f32 {
  let a = atan2(p.y,p.x);
  let organic = length(p) - (.25 + .018 * sin(a * 6.0 + params.time*.5) + .009 * cos(a * 3.0 - params.time*.3));
  let q = abs(p) - vec2f(.14,.015);
  let capsule = length(max(q,vec2f(0))) + min(max(q.x,q.y),0.0) - .17;
  return mix(capsule,organic,params.shape);
}
fn background(uv: vec2f) -> vec3f {
  let text = textureSampleLevel(src,samp,uv,0.0).rgb;
  let xy = uv * vec2f(params.aspect,1.0);
  let grid = min(abs(fract(xy.x*35.0)-.5),abs(fract(xy.y*35.0)-.5));
  let g = 1.0 - smoothstep(.015,.045,grid);
  return mix(text,vec3f(.96,.96,.93)-vec3f(g*.12),params.shape);
}
@fragment fn main(@location(0) uv: vec2f) -> @location(0) vec4f {
  let center = vec2f(params.position,.5 + .015*sin(params.time*.6));
  let p = (uv-center)*vec2f(params.aspect,1.0);
  let d = distanceToGlass(p);
  let shadow = exp(-max(distanceToGlass(p-vec2f(.012,.026)),0.0)*50.0)*.12;
  var color = background(uv)*(1.0-shadow);
  let mask = 1.0-smoothstep(-.0015,.0015,d);
  if (d < .002) {
    let eps = .0005;
    let gradient = normalize(vec2f(distanceToGlass(p+vec2f(eps,0))-distanceToGlass(p-vec2f(eps,0)),distanceToGlass(p+vec2f(0,eps))-distanceToGlass(p-vec2f(0,eps)))+vec2f(.00001));
    let thickness = .035 + params.depth*.065;
    let inner = clamp(-d/.17,0.0,1.0);
    let edge = pow(1.0-inner,3.0);
    // Magnification through the broad face, curved refraction near its rim.
    let bend = (-p * thickness*2.7 + gradient*edge*thickness*.8)/vec2f(params.aspect,1.0);
    let dispersion = gradient * edge * .003 / vec2f(params.aspect,1.0);
    let sampleUV = uv + bend;
    var glass = vec3f(background(sampleUV+dispersion).r,background(sampleUV).g,background(sampleUV-dispersion).b);
    glass = mix(glass,glass*tint()+tint()*.15,params.shape*.55);
    let rim = exp(-abs(d+.004)*260.0);
    let light = max(dot(gradient,normalize(vec2f(-.5,-1.0))),0.0);
    glass += vec3f(pow(light,5.0)*edge*.42 + rim*.35);
    glass *= 1.0 - edge*.11*(1.0-light);
    let highlight = exp(-length((p-vec2f(-.1,-.105))*vec2f(1.0,5.0))*32.0)*.13;
    glass += vec3f(highlight);
    color = mix(color,glass,mask);
  }
  return vec4f(color,1.0);
}
