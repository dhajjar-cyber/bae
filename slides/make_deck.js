const pptxgen = require('pptxgenjs');

const DEEP = '065A82', TEAL = '1C7293', MID = '21295C';
const INK = '1A1F2E', BODY = '3F4A5A', MUTE = '7C8798', WHITE = 'FFFFFF', TINT = 'EEF4F7';
const HEAD = 'Cambria', SANS = 'Calibri';

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.author = 'bae';
pres.title = 'bae — Bundle Adjustment in the Eager Mode';

const W = 13.3, H = 7.5;

function titleBar(slide, text, sub) {
  slide.addText(text, { x: 0.7, y: 0.5, w: 11.9, h: 0.8, fontSize: 38, bold: true,
    color: INK, fontFace: HEAD, isTextBox: true, margin: 0 });
  if (sub) slide.addText(sub, { x: 0.7, y: 1.35, w: 11.9, h: 0.45, fontSize: 15,
    color: MUTE, fontFace: SANS, isTextBox: true, margin: 0 });
}

function dot(slide, x, y, label, fill) {
  slide.addShape(pres.ShapeType.ellipse, { x, y, w: 0.52, h: 0.52, fill: { color: fill } });
  slide.addText(label, { x, y, w: 0.52, h: 0.52, align: 'center', valign: 'middle',
    fontSize: 15, bold: true, color: WHITE, fontFace: SANS, isTextBox: true, margin: 0 });
}

/* 1 — Title */
{
  const s = pres.addSlide();
  s.background = { color: MID };
  s.addShape(pres.ShapeType.ellipse, { x: 9.6, y: -1.6, w: 6.2, h: 6.2, fill: { color: DEEP, transparency: 35 } });
  s.addShape(pres.ShapeType.ellipse, { x: 11.2, y: 3.6, w: 3.4, h: 3.4, fill: { color: TEAL, transparency: 55 } });
  s.addText('bae', { x: 0.9, y: 2.0, w: 8, h: 1.5, fontSize: 88, bold: true, color: WHITE,
    fontFace: HEAD, isTextBox: true, margin: 0 });
  s.addText('Bundle Adjustment in the Eager Mode', { x: 0.95, y: 3.5, w: 8.6, h: 0.6,
    fontSize: 24, color: 'CADCFC', fontFace: SANS, isTextBox: true, margin: 0 });
  s.addText('Second-order optimization for robotics, built on PyTorch', { x: 0.95, y: 4.15,
    w: 8.6, h: 0.5, fontSize: 15, italic: true, color: '9FB3C8', fontFace: SANS, isTextBox: true, margin: 0 });
  s.addText('Working deck — draft v0.1', { x: 0.95, y: 6.3, w: 6, h: 0.4, fontSize: 12,
    color: '7C8798', fontFace: SANS, isTextBox: true, margin: 0 });
  s.addNotes('Starter deck for the bae library. Placeholder for venue/date and speaker names.');
}

/* 2 — The problem */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  titleBar(s, 'The problem', 'Classical solvers are fast but closed; deep-learning stacks are open but slow.');
  const rows = [
    ['C++ solvers are a wall', 'Ceres, g2o and GTSAM are fast, but live outside the autograd graph — no gradients through the optimizer, no easy custom residuals.'],
    ['Python BA is too slow', 'Naive sparse ops in eager mode blow up memory and runtime on problems with thousands of cameras and millions of points.'],
    ['Research needs both', 'Modern SfM and SLAM pipelines want differentiable, composable optimization that still scales to real datasets.'],
  ];
  rows.forEach(([h, b], i) => {
    const y = 2.15 + i * 1.55;
    dot(s, 0.7, y, String(i + 1), [DEEP, TEAL, MID][i]);
    s.addText(h, { x: 1.45, y: y - 0.04, w: 10.9, h: 0.4, fontSize: 21, bold: true,
      color: INK, fontFace: SANS, isTextBox: true, margin: 0 });
    s.addText(b, { x: 1.45, y: y + 0.42, w: 10.6, h: 0.85, fontSize: 14.5, color: BODY,
      fontFace: SANS, isTextBox: true, margin: 0 });
  });
  s.addNotes('Frame why an eager-mode, PyTorch-native BA library is worth building.');
}

/* 3 — What bae is */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  titleBar(s, 'What bae is', 'A PyTorch-native library for large-scale sparse second-order optimization.');
  const cards = [
    ['Sparse block algebra', 'Block-sparse matrix types and operations sized for BA and PGO Hessians.', DEEP],
    ['CUDA acceleration', 'Custom kernels plus optional CUDSS for the heavy linear algebra.', TEAL],
    ['Levenberg-Marquardt', 'Trust-region LM optimizer with PCG and direct sparse solvers.', MID],
    ['Autograd throughout', 'Residuals are ordinary PyTorch modules; jacobians come from the graph.', DEEP],
  ];
  cards.forEach(([h, b, c], i) => {
    const x = 0.7 + (i % 2) * 6.15, y = 2.2 + Math.floor(i / 2) * 2.3;
    s.addShape(pres.ShapeType.roundRect, { x, y, w: 5.75, h: 2.0, rectRadius: 0.12,
      fill: { color: TINT }, shadow: { type: 'outer', color: '99A6B5', blur: 10, offset: 2, angle: 90, opacity: 0.25 } });
    s.addShape(pres.ShapeType.ellipse, { x: x + 0.35, y: y + 0.35, w: 0.46, h: 0.46, fill: { color: c } });
    s.addText(h, { x: x + 1.0, y: y + 0.32, w: 4.5, h: 0.45, fontSize: 19, bold: true,
      color: INK, fontFace: SANS, isTextBox: true, margin: 0 });
    s.addText(b, { x: x + 0.35, y: y + 1.0, w: 5.05, h: 0.85, fontSize: 14, color: BODY,
      fontFace: SANS, isTextBox: true, margin: 0 });
  });
  s.addNotes('Four pillars, straight from the README feature list.');
}

/* 4 — Architecture */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  titleBar(s, 'Where the code lives', 'Repository layout as of today.');
  const mods = [
    ['bae/sparse', 'Block-sparse tensors, py_ops, solve — the core data structures and CUDA kernels.'],
    ['bae/optim', 'The LM optimizer and trust-region strategies.'],
    ['bae/autograd', 'Custom autograd functions for sparse ops.'],
    ['bae/utils', 'PCG and other Python-side solvers, helpers.'],
    ['datapipes', 'BAL loader; entry point for 1DSfM and G2O data.'],
    ['InstantSFM', 'Downstream SfM pipeline built on top of bae.'],
  ];
  mods.forEach(([name, desc], i) => {
    const y = 2.1 + i * 0.83;
    s.addShape(pres.ShapeType.roundRect, { x: 0.7, y, w: 2.75, h: 0.62, rectRadius: 0.08,
      fill: { color: i < 4 ? DEEP : TEAL } });
    s.addText(name, { x: 0.7, y, w: 2.75, h: 0.62, align: 'center', valign: 'middle',
      fontSize: 15, bold: true, color: WHITE, fontFace: 'Courier New', isTextBox: true, margin: 0 });
    s.addText(desc, { x: 3.7, y: y + 0.06, w: 8.9, h: 0.55, fontSize: 14.5, color: BODY,
      valign: 'middle', fontFace: SANS, isTextBox: true, margin: 0 });
  });
  s.addNotes('Blue = core library, teal = data and downstream users.');
}

/* 5 — How it works */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  titleBar(s, 'One optimization step', 'The LM loop, end to end.');
  const steps = [
    ['Residuals', 'A PyTorch module maps cameras and points to reprojection error.'],
    ['Jacobian', 'Autograd builds the sparse block jacobian of that module.'],
    ['Normal eqns', 'Form JᵀJ + λD and solve with PCG or CUDSS.'],
    ['Update', 'Trust-region strategy accepts or rejects, adapts λ, repeats.'],
  ];
  steps.forEach(([h, b], i) => {
    const x = 0.7 + i * 3.08;
    s.addShape(pres.ShapeType.roundRect, { x, y: 2.5, w: 2.75, h: 2.5, rectRadius: 0.12,
      fill: { color: i % 2 ? TINT : 'E3EDF2' } });
    s.addText(String(i + 1), { x: x + 0.25, y: 2.7, w: 1, h: 0.75, fontSize: 46, bold: true,
      color: [DEEP, TEAL, MID, DEEP][i], fontFace: HEAD, isTextBox: true, margin: 0 });
    s.addText(h, { x: x + 0.25, y: 3.5, w: 2.3, h: 0.4, fontSize: 18, bold: true, color: INK,
      fontFace: SANS, isTextBox: true, margin: 0 });
    s.addText(b, { x: x + 0.25, y: 3.95, w: 2.3, h: 0.95, fontSize: 13, color: BODY,
      fontFace: SANS, isTextBox: true, margin: 0 });
    if (i < 3) s.addText('→', { x: x + 2.78, y: 3.4, w: 0.3, h: 0.5, fontSize: 22,
      color: MUTE, align: 'center', fontFace: SANS, isTextBox: true, margin: 0 });
  });
  s.addText('Everything stays inside the autograd graph, so the whole loop is composable with learned components.',
    { x: 0.7, y: 5.5, w: 11.9, h: 0.5, fontSize: 15, italic: true, color: TEAL,
      fontFace: SANS, isTextBox: true, margin: 0 });
  s.addNotes('Placeholder for a figure of the sparse Hessian structure.');
}

/* 6 — Code */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  titleBar(s, 'The API in practice', 'Bundle adjustment on a BAL problem.');
  s.addShape(pres.ShapeType.roundRect, { x: 0.7, y: 2.1, w: 7.3, h: 4.5, rectRadius: 0.1,
    fill: { color: MID } });
  const code = [
    'strategy = pp.optim.strategy.TrustRegion(',
    '    up=2.0, down=0.5**4)',
    'solver   = PCG(tol=1e-4, maxiter=250)',
    'optimizer = LM(model, strategy=strategy,',
    '               solver=solver, reject=30)',
    '',
    'for idx in range(20):',
    '    loss = optimizer.step(input)',
    '    print(idx, loss.item())',
  ].join('\n');
  s.addText(code, { x: 1.0, y: 2.4, w: 6.8, h: 3.9, fontSize: 14, color: 'CADCFC',
    fontFace: 'Courier New', isTextBox: true, margin: 0, lineSpacing: 22 });
  const notes = [
    ['Familiar shape', 'The same optimizer / step pattern as any torch.optim loop.'],
    ['Swappable solvers', 'PCG today, CUDSS or a direct sparse solve with one argument.'],
    ['Your own residuals', 'model is just an nn.Module — subclass it for new problems.'],
  ];
  notes.forEach(([h, b], i) => {
    const y = 2.3 + i * 1.5;
    dot(s, 8.4, y, String.fromCharCode(65 + i), [DEEP, TEAL, MID][i]);
    s.addText(h, { x: 9.1, y: y - 0.02, w: 3.5, h: 0.4, fontSize: 17, bold: true, color: INK,
      fontFace: SANS, isTextBox: true, margin: 0 });
    s.addText(b, { x: 9.1, y: y + 0.42, w: 3.5, h: 0.9, fontSize: 13.5, color: BODY,
      fontFace: SANS, isTextBox: true, margin: 0 });
  });
  s.addNotes('Code is lifted from the README example; trim further if the slide runs long.');
}

/* 7 — Data + eval */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  titleBar(s, 'What we run it on', 'Standard benchmarks, already wired up.');
  const stats = [['3', 'dataset families supported'], ['LM', 'trust-region optimizer'], ['CUDA', 'custom sparse kernels']];
  stats.forEach(([n, l], i) => {
    const x = 0.7 + i * 4.1;
    s.addText(n, { x, y: 2.15, w: 3.7, h: 1.1, fontSize: 64, bold: true, color: [DEEP, TEAL, MID][i],
      fontFace: HEAD, isTextBox: true, margin: 0 });
    s.addText(l, { x, y: 3.3, w: 3.7, h: 0.4, fontSize: 14, color: MUTE, fontFace: SANS,
      isTextBox: true, margin: 0 });
  });
  const ds = [
    ['BAL', 'Bundle Adjustment in the Large — the ladybug and related problem sets.'],
    ['1DSfM', 'Large-scale structure-from-motion scenes.'],
    ['G2O', 'Pose graph optimization benchmarks.'],
  ];
  ds.forEach(([h, b], i) => {
    const y = 4.1 + i * 0.85;
    s.addShape(pres.ShapeType.roundRect, { x: 0.7, y, w: 11.9, h: 0.7, rectRadius: 0.08,
      fill: { color: i % 2 ? TINT : 'F6FAFC' } });
    s.addText(h, { x: 1.0, y, w: 1.6, h: 0.7, fontSize: 16, bold: true, color: DEEP,
      valign: 'middle', fontFace: SANS, isTextBox: true, margin: 0 });
    s.addText(b, { x: 2.7, y, w: 9.6, h: 0.7, fontSize: 14, color: BODY, valign: 'middle',
      fontFace: SANS, isTextBox: true, margin: 0 });
  });
  s.addNotes('Runtime and accuracy numbers still to be filled in — do not quote figures until measured.');
}

/* 8 — Next */
{
  const s = pres.addSlide();
  s.background = { color: MID };
  s.addShape(pres.ShapeType.ellipse, { x: -2.0, y: 4.2, w: 6.0, h: 6.0, fill: { color: DEEP, transparency: 45 } });
  s.addText('Where we take it next', { x: 0.8, y: 0.7, w: 11, h: 0.9, fontSize: 38, bold: true,
    color: WHITE, fontFace: HEAD, isTextBox: true, margin: 0 });
  s.addText('Open threads for this deck to grow into.', { x: 0.8, y: 1.55, w: 11, h: 0.45,
    fontSize: 15, color: '9FB3C8', fontFace: SANS, isTextBox: true, margin: 0 });
  const items = [
    ['Benchmarks', 'Measured runtime and memory against Ceres, g2o and GTSAM on BAL.'],
    ['Figures', 'Sparsity-pattern and convergence plots to replace text-heavy slides.'],
    ['Story', 'Decide the audience: paper talk, lab meeting, or adoption pitch.'],
    ['API stability', 'What we freeze before calling the library non-experimental.'],
  ];
  items.forEach(([h, b], i) => {
    const x = 0.8 + (i % 2) * 6.0, y = 2.5 + Math.floor(i / 2) * 1.9;
    dot(s, x, y, String(i + 1), TEAL);
    s.addText(h, { x: x + 0.75, y: y - 0.02, w: 4.9, h: 0.4, fontSize: 20, bold: true,
      color: WHITE, fontFace: SANS, isTextBox: true, margin: 0 });
    s.addText(b, { x: x + 0.75, y: y + 0.44, w: 4.9, h: 0.9, fontSize: 13.5, color: 'CADCFC',
      fontFace: SANS, isTextBox: true, margin: 0 });
  });
  s.addText('arxiv.org/abs/2409.12190', { x: 0.8, y: 6.5, w: 6, h: 0.4, fontSize: 13,
    color: '7C8798', fontFace: SANS, isTextBox: true, margin: 0 });
  s.addNotes('Turn these into owners and dates once we know the venue.');
}

pres.writeFile({ fileName: '/home/user/bae/slides/bae-overview.pptx' }).then(f => console.log('wrote', f));
