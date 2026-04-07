const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "Chunk 4 — Lower Bounds";
pres.title = "Monotone Classification — Lower Bounds & Proofs";

// ── Color Palette ──
const C = {
  bg: "0F1117", surface: "1A1D2E", surface2: "232640",
  accent: "6C8CFF", accent2: "A78BFA", green: "34D399",
  red: "F87171", orange: "FB923C", yellow: "FBBF24",
  text: "E2E8F0", dim: "94A3B8", border: "2D3154", white: "FFFFFF",
};
const mkShadow = () => ({ type: "outer", blur: 6, offset: 2, angle: 135, color: "000000", opacity: 0.18 });

function addCard(slide, x, y, w, h, accentColor) {
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: C.surface }, line: { color: C.border, width: 0.5 }, shadow: mkShadow() });
  if (accentColor) slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.06, h, fill: { color: accentColor } });
}
function addSlideNum(slide, n, total) {
  slide.addText(`${n} / ${total}`, { x: 8.8, y: 5.15, w: 1, h: 0.35, fontSize: 10, color: C.dim, align: "right", fontFace: "Calibri" });
}
function addTag(slide, x, y, label, color) {
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 1.6, h: 0.28, fill: { color, transparency: 80 } });
  slide.addText(label, { x, y, w: 1.6, h: 0.28, fontSize: 10, fontFace: "Calibri", bold: true, color, align: "center", valign: "middle", margin: 0 });
}
const T = 9; // total slides

// ── SLIDE 1: Title ──
{
  const s = pres.addSlide(); s.background = { color: C.bg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addText("Lower Bounds & Proofs", { x: 0.8, y: 1.4, w: 8.4, h: 1, fontSize: 42, fontFace: "Calibri", bold: true, color: C.accent, margin: 0 });
  s.addText("Monotone Classification with Relative Approximations", { x: 0.8, y: 2.4, w: 8.4, h: 0.6, fontSize: 18, fontFace: "Calibri", color: C.dim, margin: 0 });
  s.addText("Chunk 4 \u2014 Why the algorithms are near-optimal", { x: 0.8, y: 3.1, w: 8.4, h: 0.5, fontSize: 14, fontFace: "Calibri", color: C.dim, italic: true, margin: 0 });
  ["Exact case: \u03B5 = 0 needs \u03A9(n) probes", "Constant-ratio lower bound (Thm 13)", "Arbitrary \u03B5 lower bound: \u03A9(w/\u03B5\u00B2) (Thm 14)", "Synthesis: matching upper & lower bounds"].forEach((t, i) => {
    s.addText(`${i+1}.  ${t}`, { x: 1.2, y: 3.8 + i*0.38, w: 7, h: 0.36, fontSize: 13, fontFace: "Calibri", color: C.text, margin: 0 });
  });
  addSlideNum(s, 1, T);
}

// ── SLIDE 2: Recap ──
{
  const s = pres.addSlide(); s.background = { color: C.bg };
  s.addText("Recap: What We Know So Far", { x: 0.6, y: 0.3, w: 9, h: 0.6, fontSize: 30, fontFace: "Calibri", bold: true, color: C.text, margin: 0 });
  s.addText("The upper bounds leave a question \u2014 are they tight?", { x: 0.6, y: 0.9, w: 9, h: 0.4, fontSize: 14, fontFace: "Calibri", color: C.dim, margin: 0 });
  // RPE card
  addCard(s, 0.5, 1.5, 4.2, 2.3, C.accent); addTag(s, 0.7, 1.6, "UPPER BOUND", C.accent);
  s.addText("RPE Algorithm (Thm 1)", { x: 0.7, y: 1.95, w: 3.8, h: 0.35, fontSize: 16, fontFace: "Calibri", bold: true, color: C.accent, margin: 0 });
  s.addText([{ text: "Expected error \u2264 2k*", options: { breakLine: true, fontSize: 14, color: C.text } }, { text: "Expected cost: O(w \u00B7 Log(n/w))", options: { breakLine: true, fontSize: 14, color: C.text } }, { text: "", options: { breakLine: true, fontSize: 6 } }, { text: "\u2022 Simple random-probe-and-eliminate", options: { breakLine: true, fontSize: 12, color: C.dim } }, { text: "\u2022 Ratio 2 is tight for RPE", options: { fontSize: 12, color: C.dim } }], { x: 0.7, y: 2.35, w: 3.8, h: 1.3, valign: "top", margin: 0 });
  // Coreset card
  addCard(s, 5.3, 1.5, 4.2, 2.3, C.accent); addTag(s, 5.5, 1.6, "UPPER BOUND", C.accent);
  s.addText("Coreset Algorithm (Thm 6)", { x: 5.5, y: 1.95, w: 3.8, h: 0.35, fontSize: 16, fontFace: "Calibri", bold: true, color: C.accent, margin: 0 });
  s.addText([{ text: "Error \u2264 (1+\u03B5)k*  w.h.p.", options: { breakLine: true, fontSize: 14, color: C.text } }, { text: "Cost: O(w/\u03B5\u00B2 \u00B7 Log(n/w) \u00B7 log n)", options: { breakLine: true, fontSize: 14, color: C.text } }, { text: "", options: { breakLine: true, fontSize: 6 } }, { text: "\u2022 Relative-comparison coreset", options: { breakLine: true, fontSize: 12, color: C.dim } }, { text: "\u2022 Works for any \u03B5 > 0", options: { fontSize: 12, color: C.dim } }], { x: 5.5, y: 2.35, w: 3.8, h: 1.3, valign: "top", margin: 0 });
  // Key question
  addCard(s, 1.5, 4.1, 7, 1.2, C.red);
  s.addText("The Key Question", { x: 1.8, y: 4.2, w: 6.5, h: 0.3, fontSize: 16, fontFace: "Calibri", bold: true, color: C.red, margin: 0 });
  s.addText("Can we do better than these costs? Or are these algorithms near-optimal? We need lower bounds to answer this.", { x: 1.8, y: 4.5, w: 6.5, h: 0.6, fontSize: 13, fontFace: "Calibri", color: C.text, margin: 0 });
  addSlideNum(s, 2, T);
}

// ── SLIDE 3: Theorem 10 ──
{
  const s = pres.addSlide(); s.background = { color: C.bg };
  s.addText("Exact Optimality Needs \u03A9(n) Probes", { x: 0.6, y: 0.3, w: 9, h: 0.6, fontSize: 30, fontFace: "Calibri", bold: true, color: C.text, margin: 0 });
  addCard(s, 0.5, 1.0, 9, 1.0, C.red); addTag(s, 0.7, 1.08, "THEOREM 10", C.red);
  s.addText("Any algorithm that finds an optimal monotone classifier (error = k*) with probability > 2/3 must probe \u03A9(n) elements in expectation \u2014 even in 1D, even if k* is known.", { x: 0.7, y: 1.4, w: 8.6, h: 0.5, fontSize: 13, fontFace: "Calibri", color: C.text, margin: 0 });
  s.addText("The Hard Family Construction", { x: 0.6, y: 2.2, w: 9, h: 0.4, fontSize: 18, fontFace: "Calibri", bold: true, color: C.accent, margin: 0 });
  s.addText([{ text: "Group n elements into n/2 pairs: (1,2), (3,4), \u2026, (n\u22121, n)", options: { breakLine: true } }, { text: "Normal pair: odd element = +1, even element = \u22121", options: { breakLine: true } }, { text: "Each input has exactly one anomaly pair \u2014 both elements share the same label", options: {} }], { x: 0.6, y: 2.6, w: 9, h: 0.9, fontSize: 13, fontFace: "Calibri", color: C.text, margin: 0 });
  // Pair visualization
  const py = 3.7, pw = 0.9, pg = 0.15, dr = 0.22;
  const pairData = [["pair 1",false],["pair 2",false],[null,false],["anomaly!",true],[null,false],["pair n/2",false]];
  pairData.forEach(([lbl,isAnom], i) => {
    const px = 1.0 + i*(pw+pg);
    if (!lbl) { s.addText("\u2026", { x: px, y: py+0.3, w: pw, h: 0.4, fontSize: 22, color: C.dim, align: "center", margin: 0 }); return; }
    addCard(s, px, py, pw, 1.1, isAnom ? C.yellow : C.border);
    const c1 = isAnom ? C.yellow : C.accent, c2 = isAnom ? C.yellow : C.red;
    const l1 = isAnom ? "\u22121" : "+1", l2 = "\u22121";
    s.addShape(pres.shapes.OVAL, { x: px+0.33, y: py+0.18, w: dr, h: dr, fill: { color: c1 } });
    s.addText(l1, { x: px+0.33, y: py+0.18, w: dr, h: dr, fontSize: 8, color: isAnom ? C.bg : C.white, align: "center", valign: "middle", margin: 0 });
    s.addShape(pres.shapes.OVAL, { x: px+0.33, y: py+0.5, w: dr, h: dr, fill: { color: c2 } });
    s.addText(l2, { x: px+0.33, y: py+0.5, w: dr, h: dr, fontSize: 8, color: isAnom ? C.bg : C.white, align: "center", valign: "middle", margin: 0 });
    s.addText(lbl, { x: px, y: py+0.8, w: pw, h: 0.25, fontSize: 9, fontFace: "Calibri", color: isAnom ? C.yellow : C.dim, align: "center", margin: 0 });
  });
  s.addText("k* = n/2 \u2212 1 for every input. The algorithm must find the anomaly to output the optimal classifier.", { x: 0.6, y: 5.0, w: 9, h: 0.4, fontSize: 12, fontFace: "Calibri", color: C.dim, margin: 0 });
  addSlideNum(s, 3, T);
}

// ── SLIDE 4: Proof steps ──
{
  const s = pres.addSlide(); s.background = { color: C.bg };
  s.addText("Proof of \u03A9(n): Deterministic to Randomized", { x: 0.6, y: 0.3, w: 9, h: 0.6, fontSize: 28, fontFace: "Calibri", bold: true, color: C.text, margin: 0 });
  s.addText("Reducing the problem through counting and Yao\u2019s minimax", { x: 0.6, y: 0.85, w: 9, h: 0.35, fontSize: 14, fontFace: "Calibri", color: C.dim, margin: 0 });
  [{ n:"1", t:"Proposition 5", d:"No monotone classifier is optimal for both P\u208B\u2081(i) and P\u208A\u2081(i). Must find and identify the anomaly." },
   { n:"2", t:"Free labels boost", d:"Strengthen A: probing one element in a pair reveals both labels for free. Lower bound still holds." },
   { n:"3", t:"Lemma 11 (counting)", d:"A_det probes pairs x\u2081,...,x_t. If family-err \u2264 cn/2, then family-cost \u2265 n\u00B2(1\u2212c\u00B2)/4." },
   { n:"4", t:"Yao\u2019s minimax \u2192 Corollary 12", d:"Any randomized alg with family-err < n/3 has E[family-cost] = \u03A9(n\u00B2). At least one input needs \u03A9(n) probes." }
  ].forEach((st, i) => {
    const sy = 1.35 + i*1.0;
    addCard(s, 0.5, sy, 9, 0.85, C.accent);
    s.addShape(pres.shapes.OVAL, { x: 0.7, y: sy+0.22, w: 0.38, h: 0.38, fill: { color: C.accent } });
    s.addText(st.n, { x: 0.7, y: sy+0.22, w: 0.38, h: 0.38, fontSize: 14, fontFace: "Calibri", bold: true, color: C.bg, align: "center", valign: "middle", margin: 0 });
    s.addText(st.t, { x: 1.2, y: sy+0.1, w: 8, h: 0.3, fontSize: 14, fontFace: "Calibri", bold: true, color: C.accent, margin: 0 });
    s.addText(st.d, { x: 1.2, y: sy+0.4, w: 8, h: 0.4, fontSize: 12, fontFace: "Calibri", color: C.dim, margin: 0 });
  });
  addSlideNum(s, 4, T);
}

// ── SLIDE 5: Lemma 11 ──
{
  const s = pres.addSlide(); s.background = { color: C.bg };
  s.addText("Lemma 11: The Counting Argument", { x: 0.6, y: 0.3, w: 9, h: 0.6, fontSize: 28, fontFace: "Calibri", bold: true, color: C.text, margin: 0 });
  s.addText("If a deterministic algorithm errs on few inputs, it must probe many pairs", { x: 0.6, y: 0.85, w: 9, h: 0.35, fontSize: 13, fontFace: "Calibri", color: C.dim, margin: 0 });
  s.addText("A_det probes a fixed sequence of pairs: x\u2081, x\u2082, ..., x_t. Stops when anomaly found, or outputs fixed h_det.", { x: 0.6, y: 1.3, w: 9, h: 0.5, fontSize: 13, fontFace: "Calibri", color: C.text, margin: 0 });
  const hdr = { bold: true, color: C.accent, fill: { color: C.surface2 }, fontSize: 12 };
  const cel = { color: C.text, fill: { color: C.surface }, fontSize: 12 };
  s.addTable([[{ text: "Input type", options: hdr },{ text: "Anomaly pair", options: hdr },{ text: "Probes", options: hdr }],[{ text: "i \u2209 {x\u2081,...,x_t}", options: cel },{ text: "Never found", options: {...cel, color: C.dim} },{ text: "t probes each", options: cel }],[{ text: "i = x_j", options: cel },{ text: "Found at pos j", options: {...cel, color: C.dim} },{ text: "j probes", options: cel }]], { x: 0.8, y: 1.9, w: 8.4, colW: [3,2.7,2.7], border: { pt: 0.5, color: C.border } });
  [{ l: "ERRORS", c: C.red, t: "family-err \u2265 n/2 \u2212 t   (must fail on P\u208B\u2081(i) or P\u208A\u2081(i) for un-probed i)" },
   { l: "TOTAL COST", c: C.accent, t: "family-cost = 2t(n/2 \u2212 t) + 2\u03A3j = nt \u2212 t\u00B2 + t" },
   { l: "CONSTRAINT", c: C.green, t: "family-err \u2264 cn/2  \u21D2  t \u2265 n(1\u2212c)/2" }
  ].forEach((d, i) => {
    const dy = 3.15 + i*0.6;
    addCard(s, 0.5, dy, 9, 0.5, d.c);
    s.addText(d.l, { x: 0.7, y: dy+0.08, w: 1.6, h: 0.3, fontSize: 10, fontFace: "Calibri", bold: true, color: d.c, margin: 0 });
    s.addText(d.t, { x: 2.3, y: dy+0.08, w: 7, h: 0.3, fontSize: 12, fontFace: "Calibri", color: C.text, margin: 0 });
  });
  addCard(s, 1.5, 4.85, 7, 0.55, C.red);
  s.addText("For t \u2208 [n(1\u2212c)/2, n/2]:   family-cost \u2265 nt \u2212 t\u00B2 \u2265 n\u00B2(1\u2212c\u00B2)/4", { x: 1.8, y: 4.9, w: 6.5, h: 0.45, fontSize: 16, fontFace: "Calibri", bold: true, color: C.red, align: "center", valign: "middle", margin: 0 });
  addSlideNum(s, 5, T);
}

// ── SLIDE 6: Theorem 13 ──
{
  const s = pres.addSlide(); s.background = { color: C.bg };
  s.addText("Lower Bound for Constant Approximation", { x: 0.6, y: 0.3, w: 9, h: 0.6, fontSize: 28, fontFace: "Calibri", bold: true, color: C.text, margin: 0 });
  addCard(s, 0.5, 0.95, 9, 0.9, C.red); addTag(s, 0.7, 1.03, "THEOREM 13", C.red);
  s.addText("For any constant c \u2265 1, any algorithm guaranteeing expected error \u2264 c\u00B7k* must probe \u03A9(w\u2032 \u00B7 Log(n\u2032/w\u2032)) elements in expectation.", { x: 0.7, y: 1.35, w: 8.6, h: 0.4, fontSize: 13, fontFace: "Calibri", color: C.text, margin: 0 });
  addCard(s, 0.5, 2.1, 4.3, 2.0, C.orange);
  s.addText("Key Idea: Dummy Points", { x: 0.7, y: 2.2, w: 3.9, h: 0.3, fontSize: 15, fontFace: "Calibri", bold: true, color: C.orange, margin: 0 });
  s.addText([{ text: "Start from realizable hard family (w\u2032 boxes, n\u2032/w\u2032 points each).", options: { breakLine: true } }, { text: "", options: { breakLine: true, fontSize: 6 } }, { text: "Add 2ck dummy points between consecutive points. Dummy box with 2k points forces k* = k.", options: { breakLine: true } }, { text: "", options: { breakLine: true, fontSize: 6 } }, { text: "Each non-dummy point surrounded by ck matching dummies.", options: {} }], { x: 0.7, y: 2.55, w: 3.9, h: 1.4, fontSize: 12, fontFace: "Calibri", color: C.text, margin: 0 });
  addCard(s, 5.2, 2.1, 4.3, 2.0, C.green);
  s.addText("Reduction to Realizable Case", { x: 5.4, y: 2.2, w: 3.9, h: 0.3, fontSize: 15, fontFace: "Calibri", bold: true, color: C.green, margin: 0 });
  s.addText([{ text: "Misclassifying any non-dummy point \u2192 \u2265 ck+1 errors \u2192 violates c\u00B7k* guarantee.", options: { breakLine: true } }, { text: "", options: { breakLine: true, fontSize: 6 } }, { text: "Algorithm must correctly classify all non-dummy points.", options: { breakLine: true } }, { text: "", options: { breakLine: true, fontSize: 6 } }, { text: "\u21D2 Inherits \u03A9(w\u2032 Log(n\u2032/w\u2032)) lower bound from realizable case.", options: { bold: true } }], { x: 5.4, y: 2.55, w: 3.9, h: 1.4, fontSize: 12, fontFace: "Calibri", color: C.text, margin: 0 });
  s.addText("Visual: w\u2032 independent boxes + dummy box", { x: 0.6, y: 4.3, w: 9, h: 0.3, fontSize: 13, fontFace: "Calibri", bold: true, color: C.accent, margin: 0 });
  ["B\u2081","B\u2082","\u2026","B_w\u2032","+","Dummy"].forEach((lbl, i) => {
    const bx = 1.0 + i*1.3, bc = i===5 ? C.red : C.accent;
    if (i===2||i===4) { s.addText(lbl, { x: bx, y: 4.7, w: 1, h: 0.5, fontSize: 18, color: C.dim, align: "center", valign: "middle", margin: 0 }); }
    else { s.addShape(pres.shapes.RECTANGLE, { x: bx, y: 4.65, w: 1.0, h: 0.65, fill: { color: C.surface }, line: { color: bc, width: 1.5 } }); s.addText(lbl, { x: bx, y: 4.65, w: 1.0, h: 0.65, fontSize: 13, fontFace: "Calibri", bold: true, color: bc, align: "center", valign: "middle", margin: 0 }); }
  });
  addSlideNum(s, 6, T);
}

// ── SLIDE 7: Theorem 14 ──
{
  const s = pres.addSlide(); s.background = { color: C.bg };
  s.addText("Lower Bound for Arbitrary \u03B5", { x: 0.6, y: 0.3, w: 9, h: 0.6, fontSize: 28, fontFace: "Calibri", bold: true, color: C.text, margin: 0 });
  addCard(s, 0.5, 0.95, 9, 0.85, C.red); addTag(s, 0.7, 1.03, "THEOREM 14", C.red);
  s.addText("For 0 < \u03B5 \u2264 1/10, any algorithm guaranteeing expected error (1+\u03B5)k* must probe \u03A9(w/\u03B5\u00B2) elements. Holds in d = 2.", { x: 0.7, y: 1.35, w: 8.6, h: 0.4, fontSize: 13, fontFace: "Calibri", color: C.text, margin: 0 });
  s.addText("Hard Input Construction", { x: 0.6, y: 2.0, w: 9, h: 0.35, fontSize: 16, fontFace: "Calibri", bold: true, color: C.accent, margin: 0 });
  s.addText("Place n/w points at each of w locations x\u2081,...,x_w (no dominance). Each location gets a random bias \u03BC[i] \u2208 {\u03BC\u2081, \u03BC\u2082}.", { x: 0.6, y: 2.35, w: 9, h: 0.4, fontSize: 13, fontFace: "Calibri", color: C.text, margin: 0 });
  [["x\u2081","\u03BC\u2081",C.red],["x\u2082","\u03BC\u2082",C.accent],["x\u2083","\u03BC\u2081",C.red],[null],["x_w","\u03BC\u2082",C.accent]].forEach((d, i) => {
    const lx = 1.5 + i*1.4;
    if (!d[0]) { s.addText("\u2026", { x: lx, y: 3.1, w: 1, h: 0.5, fontSize: 20, color: C.dim, align: "center", margin: 0 }); return; }
    s.addShape(pres.shapes.RECTANGLE, { x: lx, y: 2.85, w: 1.0, h: 1.3, fill: { color: C.surface }, line: { color: C.border, width: 0.5 } });
    s.addText(d[1], { x: lx+0.15, y: 2.9, w: 0.7, h: 0.25, fontSize: 10, fontFace: "Calibri", bold: true, color: d[2], align: "center", margin: 0 });
    for (let j = 0; j < 5; j++) { const dc = Math.random() > (d[2]===C.red ? 0.55 : 0.45) ? C.accent : C.red; s.addShape(pres.shapes.OVAL, { x: lx+0.38, y: 3.2+j*0.18, w: 0.12, h: 0.12, fill: { color: dc } }); }
    s.addText(d[0], { x: lx, y: 4.15, w: 1.0, h: 0.2, fontSize: 10, fontFace: "Calibri", color: C.dim, align: "center", margin: 0 });
  });
  addCard(s, 0.5, 4.5, 4.3, 0.9, C.orange);
  s.addText([{ text: "\u03BC\u2081 = (1\u2212\u03B3)/2  slightly below 1/2", options: { breakLine: true, color: C.orange } }, { text: "\u03BC\u2082 = (1+\u03B3)/2  slightly above 1/2", options: { breakLine: true, color: C.accent } }, { text: "where \u03B3 = 9\u03B5", options: { color: C.dim } }], { x: 0.7, y: 4.55, w: 3.9, h: 0.8, fontSize: 12, fontFace: "Calibri", margin: 0 });
  addCard(s, 5.2, 4.5, 4.3, 0.9, C.red);
  s.addText([{ text: "Core difficulty: distinguishing \u03BC\u2081 from \u03BC\u2082", options: { breakLine: true, color: C.red } }, { text: "requires M = \u0398(1/\u03B5\u00B2) samples per location", options: { breakLine: true, color: C.text } }, { text: "(Anthony-Bartlett / Le Cam)", options: { color: C.dim } }], { x: 5.4, y: 4.55, w: 3.9, h: 0.8, fontSize: 12, fontFace: "Calibri", margin: 0 });
  addSlideNum(s, 7, T);
}

// ── SLIDE 8: RP-1 & RP-2 ──
{
  const s = pres.addSlide(); s.background = { color: C.bg };
  s.addText("Proof Sketch: Two Random Processes", { x: 0.6, y: 0.3, w: 9, h: 0.55, fontSize: 26, fontFace: "Calibri", bold: true, color: C.text, margin: 0 });
  s.addText("Showing E[R\u2082] > 1 + \u03B5 when the probe budget is too small", { x: 0.6, y: 0.8, w: 9, h: 0.3, fontSize: 13, fontFace: "Calibri", color: C.dim, margin: 0 });
  addCard(s, 0.4, 1.2, 4.1, 2.1, C.accent);
  s.addText("RP-1: All Labels Upfront", { x: 0.6, y: 1.3, w: 3.7, h: 0.3, fontSize: 14, fontFace: "Calibri", bold: true, color: C.accent, margin: 0 });
  ["1. Sample \u03BC \u2208 {\u03BC\u2081,\u03BC\u2082}^w","2. Label every point with Pr[+1]=\u03BC[i]","3. Run algorithm A on labeled P","4. Measure R\u2081 = err(h_A)/k*"].forEach((t,i) => { s.addText(t, { x: 0.6, y: 1.65+i*0.35, w: 3.7, h: 0.3, fontSize: 11, fontFace: "Calibri", color: C.text, margin: 0 }); });
  s.addText("\u21D4", { x: 4.55, y: 1.8, w: 0.9, h: 0.4, fontSize: 24, color: C.green, align: "center", valign: "middle", margin: 0 });
  s.addText("Lemma 15\nE[R\u2081]=E[R\u2082]", { x: 4.45, y: 2.2, w: 1.1, h: 0.5, fontSize: 9, fontFace: "Calibri", color: C.green, align: "center", margin: 0 });
  addCard(s, 5.5, 1.2, 4.1, 2.1, C.accent2);
  s.addText("RP-2: Labels On-Demand", { x: 5.7, y: 1.3, w: 3.7, h: 0.3, fontSize: 14, fontFace: "Calibri", bold: true, color: C.accent2, margin: 0 });
  ["1. Sample \u03BC \u2208 {\u03BC\u2081,\u03BC\u2082}^w","2. When A probes p, assign label","3. After A finishes, label rest","4. Measure R\u2082 = err(h_A)/k*"].forEach((t,i) => { s.addText(t, { x: 5.7, y: 1.65+i*0.35, w: 3.7, h: 0.3, fontSize: 11, fontFace: "Calibri", color: C.text, margin: 0 }); });
  [{ tag: "Lemma 17", c: C.yellow, t: "Distinguishing \u03BC\u2081 from \u03BC\u2082 with > 2/3 prob needs M = \u0398(1/\u03B5\u00B2) samples" },
   { tag: "Lemma 18", c: C.accent, t: "With budget wM/8, more than w/2 locations are \u201Clight\u201D (probed < M times)" },
   { tag: "Lemma 19", c: C.accent2, t: "At each light location, algorithm guesses \u03BC[i] wrong with prob > 1/4" },
   { tag: "Result", c: C.green, t: "Each wrong guess adds \u2265 n\u03B3/(2w) excess error \u21D2 E[R\u2082] > 1 + \u03B5" }
  ].forEach((l, i) => {
    const ly = 3.5 + i*0.5;
    addCard(s, 0.4, ly, 9.2, 0.44, l.c);
    s.addText(l.tag, { x: 0.6, y: ly+0.06, w: 1.3, h: 0.28, fontSize: 10, fontFace: "Calibri", bold: true, color: l.c, margin: 0 });
    s.addText(l.t, { x: 2.0, y: ly+0.06, w: 7.4, h: 0.28, fontSize: 11, fontFace: "Calibri", color: C.text, margin: 0 });
  });
  addSlideNum(s, 8, T);
}

// ── SLIDE 9: Summary ──
{
  const s = pres.addSlide(); s.background = { color: C.bg };
  s.addText("Near-Optimality: The Full Picture", { x: 0.6, y: 0.25, w: 9, h: 0.55, fontSize: 28, fontFace: "Calibri", bold: true, color: C.text, margin: 0 });
  s.addText("Upper and lower bounds nearly match across all regimes of \u03B5", { x: 0.6, y: 0.75, w: 9, h: 0.3, fontSize: 13, fontFace: "Calibri", color: C.dim, margin: 0 });
  const h = { bold: true, color: C.accent, fill: { color: C.surface2 }, fontSize: 11, align: "center", valign: "middle" };
  const c = (cl) => ({ color: cl||C.text, fill: { color: C.surface }, fontSize: 11, valign: "middle" });
  s.addTable([
    [{ text: "\u03B5", options: h },{ text: "Upper Bound", options: h },{ text: "Lower Bound", options: h },{ text: "Gap", options: h }],
    [{ text: "\u03B5 = 0", options: {...c(C.red), bold: true} },{ text: "O(n)", options: c() },{ text: "\u03A9(n)  [Thm 10]", options: c(C.red) },{ text: "Tight", options: {...c(C.green), bold: true} }],
    [{ text: "const c>1", options: {...c(C.orange), bold: true} },{ text: "O(w\u00B7Log(n/w)) [RPE]", options: c() },{ text: "\u03A9(w\u00B7Log(n/(k*+1)w)) [Thm 13]", options: c(C.red) },{ text: "Tight*", options: {...c(C.green), bold: true} }],
    [{ text: "any \u03B5>0", options: {...c(C.accent), bold: true} },{ text: "O(w/\u03B5\u00B2\u00B7Log(n/w)\u00B7log n) [Thm 6]", options: c() },{ text: "\u03A9(w/\u03B5\u00B2) [Thm 14]", options: c(C.red) },{ text: "O(Log\u00B7log)", options: {...c(C.yellow), bold: true} }],
  ], { x: 0.5, y: 1.15, w: 9, colW: [1.3,3.2,3.2,1.3], border: { pt: 0.5, color: C.border } });
  s.addText("* Tight when k* \u2264 (n/w)^{1\u2212\u03B4} for any small \u03B4 > 0.", { x: 0.6, y: 2.8, w: 9, h: 0.25, fontSize: 10, fontFace: "Calibri", italic: true, color: C.dim, margin: 0 });
  [{ t: "\u03B5 = 0 is Hopeless", c: C.red, d: "Finding the exact optimum requires reading essentially the entire input." },
   { t: "Width w is the Right Parameter", c: C.orange, d: "Both bounds scale with w, confirming dominance width captures intrinsic difficulty." },
   { t: "1/\u03B5\u00B2 Dependence is Necessary", c: C.accent, d: "The \u03A9(w/\u03B5\u00B2) lower bound shows the quadratic dependence on 1/\u03B5 is unavoidable." },
   { t: "Only Polylog Gap Remains", c: C.green, d: "Gap is O(Log(n/w)\u00B7log n). Closing it is an open problem." }
  ].forEach((tk, i) => {
    const col = i%2, row = Math.floor(i/2), tx = 0.5+col*4.7, ty = 3.15+row*1.15;
    addCard(s, tx, ty, 4.3, 1.0, tk.c);
    s.addText(tk.t, { x: tx+0.2, y: ty+0.08, w: 3.9, h: 0.28, fontSize: 13, fontFace: "Calibri", bold: true, color: tk.c, margin: 0 });
    s.addText(tk.d, { x: tx+0.2, y: ty+0.38, w: 3.9, h: 0.55, fontSize: 11, fontFace: "Calibri", color: C.dim, margin: 0 });
  });
  addSlideNum(s, 9, T);
}

pres.writeFile({ fileName: "C:/Users/kex03/OneDrive/Desktop/sc2320-presentation/Chunk4_LowerBounds.pptx" })
  .then(() => console.log("DONE"))
  .catch(e => console.error(e));
