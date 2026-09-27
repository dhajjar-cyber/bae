const pptxgen = require('pptxgenjs');

const NAVY = '1E2761', SLATE = '3E5C8A', ICE = 'CADCFC', MIST = 'EDF2FA';
const INK = '16203A', BODY = '44506B', MUTE = '8792AB', WHITE = 'FFFFFF', ACCENT = 'B8862B';
const HEAD = 'Cambria', SANS = 'Calibri';
const TBD = 'to be defined';

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.title = 'EAIND — FSI Strategy Development Plan';
pres.author = 'Engineering, AI and Data';

function head(s, title, sub) {
  s.addText(title, { x: 0.7, y: 0.45, w: 11.9, h: 0.75, fontSize: 36, bold: true,
    color: INK, fontFace: HEAD, isTextBox: true, margin: 0 });
  if (sub) s.addText(sub, { x: 0.7, y: 1.24, w: 11.9, h: 0.45, fontSize: 14.5,
    color: MUTE, fontFace: SANS, isTextBox: true, margin: 0 });
}
function disc(s, x, y, label, fill, d) {
  d = d || 0.5;
  s.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: fill } });
  s.addText(label, { x, y, w: d, h: d, align: 'center', valign: 'middle', fontSize: d > 0.45 ? 15 : 12,
    bold: true, color: WHITE, fontFace: SANS, isTextBox: true, margin: 0 });
}
function openTag(s, x, y, w, text) {
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h: 0.33, rectRadius: 0.16,
    fill: { color: WHITE }, line: { color: ACCENT, width: 1 } });
  s.addText(text, { x, y, w, h: 0.33, align: 'center', valign: 'middle', fontSize: 10.5,
    italic: true, color: ACCENT, fontFace: SANS, isTextBox: true, margin: 0 });
}

/* 1 — Title */
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addShape(pres.ShapeType.ellipse, { x: 9.3, y: -1.8, w: 6.6, h: 6.6, fill: { color: SLATE, transparency: 45 } });
  s.addShape(pres.ShapeType.ellipse, { x: 11.0, y: 3.9, w: 3.6, h: 3.6, fill: { color: ICE, transparency: 80 } });
  s.addText('Engineering, AI and Data', { x: 0.9, y: 1.9, w: 9, h: 0.45, fontSize: 16,
    color: ICE, charSpacing: 2, fontFace: SANS, isTextBox: true, margin: 0 });
  s.addText('FSI Strategy\nDevelopment Plan', { x: 0.88, y: 2.45, w: 9, h: 2.0, fontSize: 48,
    bold: true, color: WHITE, fontFace: HEAD, isTextBox: true, margin: 0, lineSpacing: 52 });
  s.addText('The journey to a defined go-to-market and activation strategy for financial services',
    { x: 0.93, y: 4.55, w: 8.6, h: 0.6, fontSize: 15, italic: true, color: ICE,
      fontFace: SANS, isTextBox: true, margin: 0 });
  s.addText('Framework only — content to be developed over a three-week cycle',
    { x: 0.93, y: 6.35, w: 8, h: 0.4, fontSize: 12, color: '9AA7C4', fontFace: SANS, isTextBox: true, margin: 0 });
  s.addNotes('This deck sets out how we will build the FSI strategy. It deliberately carries no answers — markets, accounts, offerings, targets and owners are outputs of the cycle described here.');
}

/* 2 — How to use */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, 'What this document is', 'A method, not a conclusion.');
  const pts = [
    ['A plan to build the strategy', 'It describes the journey, the questions and the sequence — it does not pre-empt the answers.'],
    ['Every element is deliberately open', 'Markets, accounts, offerings, targets and owners are left blank by design and filled through the workstreams.'],
    ['One cycle, one sign-off', 'The strategy is treated as a single three-week cycle ending in a decision, not a rolling exercise.'],
  ];
  pts.forEach(([h, b], i) => {
    const y = 2.0 + i * 1.35;
    disc(s, 0.7, y, String(i + 1), [NAVY, SLATE, NAVY][i]);
    s.addText(h, { x: 1.45, y: y - 0.03, w: 6.5, h: 0.4, fontSize: 19, bold: true, color: INK,
      fontFace: SANS, isTextBox: true, margin: 0 });
    s.addText(b, { x: 1.45, y: y + 0.42, w: 6.6, h: 0.75, fontSize: 14, color: BODY,
      fontFace: SANS, isTextBox: true, margin: 0 });
  });
  s.addShape(pres.ShapeType.roundRect, { x: 8.6, y: 1.95, w: 4.0, h: 4.2, rectRadius: 0.12, fill: { color: MIST } });
  s.addText('What "done" looks like', { x: 8.95, y: 2.25, w: 3.4, h: 0.4, fontSize: 17, bold: true,
    color: NAVY, fontFace: SANS, isTextBox: true, margin: 0 });
  const done = ['An agreed service catalogue', 'A named target account list', 'Revenue and pricing targets', 'An owner against every account', 'A first-90-days activation plan'];
  s.addText(done.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < done.length - 1 } })),
    { x: 8.95, y: 2.8, w: 3.4, h: 3.0, fontSize: 13.5, color: BODY, fontFace: SANS,
      isTextBox: true, margin: 0, paraSpaceAfter: 10 });
  s.addNotes('Set expectations up front: reviewers should not look for answers in this pack.');
}

/* 3 — Framework */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, 'Five questions to answer', 'Each dimension is owned by one workstream and closed at sign-off.');
  const dims = [
    ['Where to play', ['Which markets, in what priority?', 'Which client segments?', 'Which named accounts?']],
    ['What we offer', ['Which integrated propositions across the three pillars?', 'Which partial offerings are permitted, and when?']],
    ['How we win', ['What differentiates us?', 'Which alliances and credentials are needed?', 'What delivery model applies?']],
    ['How much', ['What revenue by market and by offering?', 'What deal sizes and price bands does the market pay?']],
    ['Who covers what', ['Which partner owns which account?', 'What coverage model across markets?', 'What capacity is required?']],
  ];
  dims.forEach(([t, qs], i) => {
    const x = 0.7 + i * 2.42;
    s.addShape(pres.ShapeType.roundRect, { x, y: 2.05, w: 2.22, h: 3.85, rectRadius: 0.12,
      fill: { color: i % 2 ? MIST : 'E4EBF7' } });
    disc(s, x + 0.25, 2.3, String(i + 1), NAVY, 0.42);
    s.addText(t, { x: x + 0.25, y: 2.85, w: 1.8, h: 0.6, fontSize: 16, bold: true, color: INK,
      fontFace: SANS, isTextBox: true, margin: 0 });
    s.addText(qs.map((q, j) => ({ text: q, options: { bullet: true, breakLine: j < qs.length - 1 } })),
      { x: x + 0.25, y: 3.5, w: 1.78, h: 2.2, fontSize: 11.5, color: BODY, fontFace: SANS,
        isTextBox: true, margin: 0, paraSpaceAfter: 7 });
  });
  openTag(s, 4.65, 6.15, 4.0, 'answers ' + TBD + ' in the cycle');
  s.addNotes('Framing follows a standard strategy cascade, adapted to a consulting portfolio.');
}

/* 4 — Principles */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, 'Principles we hold going in', 'Constraints on the answers, agreed before the work starts.');
  const pr = [
    ['Integrated, not siloed', 'Cross-pillar propositions are the default catalogue; single-pillar sales are an entry point, not the destination.'],
    ['Named accounts, not averages', 'The output is a specific target list with owners, not a market-size view.'],
    ['Sized and priced', 'Every proposition carries a buyer, an indicative deal size and a price band.'],
    ['Covered by a partner', 'Every priority account has one accountable owner and a defined partner play.'],
    ['Deliverable by design', 'Nothing enters the catalogue that we cannot staff and deliver to standard.'],
  ];
  pr.forEach(([h, b], i) => {
    const x = 0.7 + (i % 2) * 6.15, y = 1.95 + Math.floor(i / 2) * 1.62;
    s.addShape(pres.ShapeType.roundRect, { x, y, w: 5.75, h: 1.4, rectRadius: 0.1, fill: { color: MIST } });
    disc(s, x + 0.3, y + 0.28, String(i + 1), i % 2 ? SLATE : NAVY, 0.42);
    s.addText(h, { x: x + 0.92, y: y + 0.22, w: 4.6, h: 0.35, fontSize: 17, bold: true, color: INK,
      fontFace: SANS, isTextBox: true, margin: 0 });
    s.addText(b, { x: x + 0.92, y: y + 0.63, w: 4.55, h: 0.65, fontSize: 12.5, color: BODY,
      fontFace: SANS, isTextBox: true, margin: 0 });
  });
  s.addText('These are the rules the strategy must obey. Anything else remains open.',
    { x: 6.85, y: 6.05, w: 5.6, h: 0.5, fontSize: 13.5, italic: true, color: SLATE,
      fontFace: SANS, isTextBox: true, margin: 0 });
  s.addNotes('Confirm these principles at kick-off; they bound every later decision.');
}

/* 5 — Approach */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, 'How we will get there', 'A five-step approach, run once, end to end.');
  const steps = [
    ['Mobilise', ['Confirm scope, principles and owners', 'Agree inputs and cadence'], 'Week 1'],
    ['Baseline', ['Where we stand today: credentials, pipeline, capability', 'What we have already won, and where'], 'Week 1'],
    ['Define', ['Markets, segments and candidate accounts', 'Integrated proposition catalogue'], 'Weeks 1–2'],
    ['Quantify', ['Deal sizes and market price bands', 'Revenue targets by market and offering'], 'Week 2'],
    ['Commit', ['Account ownership and coverage', 'Activation plan and sign-off'], 'Week 3'],
  ];
  steps.forEach(([name, pts, wk], i) => {
    const x = 0.7 + i * 2.42, cw = 2.22;
    s.addShape(pres.ShapeType.chevron, { x, y: 2.05, w: cw, h: 0.78,
      fill: { color: i % 2 ? SLATE : NAVY } });
    s.addText(name, { x: x + 0.18, y: 2.05, w: cw - 0.3, h: 0.78, align: 'center', valign: 'middle',
      fontSize: 15, bold: true, color: WHITE, fontFace: SANS, isTextBox: true, margin: 0 });
    disc(s, x + cw / 2 - 0.19, 3.05, String(i + 1), ACCENT, 0.38);
    s.addText(pts.map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < pts.length - 1 } })),
      { x, y: 3.65, w: cw, h: 1.7, fontSize: 12, color: BODY, fontFace: SANS, isTextBox: true,
        margin: 0, paraSpaceAfter: 8 });
    s.addText(wk, { x, y: 5.45, w: cw, h: 0.35, align: 'center', fontSize: 11.5, bold: true,
      color: MUTE, fontFace: SANS, isTextBox: true, margin: 0 });
  });
  s.addText('Each step closes before the next opens; the gate is a decision, not a document.',
    { x: 0.7, y: 6.15, w: 11.9, h: 0.4, fontSize: 13.5, italic: true, color: SLATE,
      fontFace: SANS, isTextBox: true, margin: 0 });
  s.addNotes('This is the process view. The workstream and calendar views that follow are the same journey cut by owner and by week.');
}

/* 6 — Workstreams */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, 'Five workstreams, five artefacts', 'Each workstream owns one dimension and produces one deliverable.');
  const rows = [
    [{ text: 'Workstream', options: { bold: true } }, { text: 'Question it closes', options: { bold: true } }, { text: 'Artefact produced', options: { bold: true } }],
    ['1 · Proposition & catalogue', 'What we offer', 'Integrated proposition catalogue with permitted variants'],
    ['2 · Market & account targeting', 'Where to play', 'Prioritised markets, segments and named account list'],
    ['3 · Commercial sizing & pricing', 'How much', 'Deal-size bands, price benchmarks, revenue targets'],
    ['4 · Coverage & ownership', 'Who covers what', 'Account-to-partner map and coverage model'],
    ['5 · Delivery model & readiness', 'How we win', 'Delivery approach and capability or capacity gaps'],
  ];
  s.addTable(rows, {
    x: 0.7, y: 2.1, w: 11.9, colW: [3.3, 2.6, 6.0], fontSize: 13, fontFace: SANS, color: BODY,
    border: { type: 'solid', color: 'D8E0EE', pt: 1 }, align: 'left', valign: 'middle',
    rowH: [0.5, 0.72, 0.72, 0.72, 0.72, 0.72], margin: 10,
    fill: { color: WHITE },
  });
  s.addShape(pres.ShapeType.rect, { x: 0.7, y: 2.1, w: 11.9, h: 0.5, fill: { color: NAVY } });
  s.addText([{ text: 'Workstream', options: { bold: true } }],
    { x: 0.85, y: 2.1, w: 3.2, h: 0.5, valign: 'middle', fontSize: 13, color: WHITE, fontFace: SANS, isTextBox: true, margin: 0 });
  s.addText([{ text: 'Question it closes', options: { bold: true } }],
    { x: 4.15, y: 2.1, w: 2.5, h: 0.5, valign: 'middle', fontSize: 13, color: WHITE, fontFace: SANS, isTextBox: true, margin: 0 });
  s.addText([{ text: 'Artefact produced', options: { bold: true } }],
    { x: 6.75, y: 2.1, w: 5.6, h: 0.5, valign: 'middle', fontSize: 13, color: WHITE, fontFace: SANS, isTextBox: true, margin: 0 });
  s.addText('Workstreams run in parallel; dependencies are resolved at the weekly checkpoint.',
    { x: 0.7, y: 6.3, w: 11.9, h: 0.4, fontSize: 13, italic: true, color: MUTE, fontFace: SANS, isTextBox: true, margin: 0 });
  s.addNotes('Artefact templates are blank at this stage; they are populated during the cycle.');
}

/* 7 — Three-week journey */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, 'The three-week journey', 'Three weeks, three gates, one decision.');
  const weeks = [
    ['Week 1', 'Frame', ['Baseline current position and credentials', 'Prioritise markets and segments', 'Build the account longlist', 'Draft the proposition catalogue'], 'Gate: agreed market and segment priorities'],
    ['Week 2', 'Define', ['Lock propositions and permitted variants', 'Benchmark deal sizes and pricing', 'Shortlist target accounts', 'Draft the coverage map'], 'Gate: agreed catalogue and shortlist'],
    ['Week 3', 'Commit', ['Set revenue targets by market and offering', 'Assign an owner to every account', 'Build the first-90-days activation plan', 'Review and sign off'], 'Gate: strategy signed off'],
  ];
  weeks.forEach(([w, label, acts, gate], i) => {
    const x = 0.7 + i * 4.07;
    s.addShape(pres.ShapeType.roundRect, { x, y: 2.0, w: 3.75, h: 4.2, rectRadius: 0.12, fill: { color: MIST } });
    s.addShape(pres.ShapeType.rect, { x, y: 2.0, w: 3.75, h: 0.75, fill: { color: [NAVY, SLATE, NAVY][i] } });
    s.addText(w + '  ·  ' + label, { x: x + 0.3, y: 2.0, w: 3.2, h: 0.75, valign: 'middle',
      fontSize: 18, bold: true, color: WHITE, fontFace: SANS, isTextBox: true, margin: 0 });
    s.addText(acts.map((a, j) => ({ text: a, options: { bullet: true, breakLine: j < acts.length - 1 } })),
      { x: x + 0.3, y: 3.0, w: 3.15, h: 2.1, fontSize: 12.5, color: BODY, fontFace: SANS,
        isTextBox: true, margin: 0, paraSpaceAfter: 9 });
    s.addText(gate, { x: x + 0.3, y: 5.35, w: 3.15, h: 0.65, fontSize: 12, bold: true,
      color: ACCENT, fontFace: SANS, isTextBox: true, margin: 0 });
    if (i < 2) s.addText('→', { x: x + 3.78, y: 3.8, w: 0.28, h: 0.5, fontSize: 20,
      color: MUTE, align: 'center', fontFace: SANS, isTextBox: true, margin: 0 });
  });
  s.addText('Start date ' + TBD + ' at kick-off.', { x: 0.7, y: 6.4, w: 6, h: 0.35, fontSize: 12,
    italic: true, color: MUTE, fontFace: SANS, isTextBox: true, margin: 0 });
  s.addNotes('Timebox is deliberate: three weeks forces decisions rather than analysis.');
}

/* 8 — Governance */
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addShape(pres.ShapeType.ellipse, { x: -2.2, y: 4.4, w: 6.2, h: 6.2, fill: { color: SLATE, transparency: 55 } });
  s.addText('How we run it', { x: 0.8, y: 0.6, w: 11, h: 0.8, fontSize: 36, bold: true,
    color: WHITE, fontFace: HEAD, isTextBox: true, margin: 0 });
  s.addText('Cadence, inputs and the decisions we are asking for.', { x: 0.8, y: 1.42, w: 11, h: 0.45,
    fontSize: 14.5, color: ICE, fontFace: SANS, isTextBox: true, margin: 0 });
  const cols = [
    ['Cadence', ['Kick-off to confirm scope and principles', 'Weekly working session per workstream', 'Checkpoint at the end of each week', 'Sign-off session in week three']],
    ['Inputs we need', ['Current pipeline and revenue by pillar', 'Existing account relationships', 'Alliance and partnership status', 'Delivery capacity and skills view', 'Historical deal values']],
    ['Decisions requested', ['Priority markets and segments', 'Target account list', 'Revenue and margin targets', 'Account ownership', 'Investment and hiring asks']],
  ];
  cols.forEach(([h, items], i) => {
    const x = 0.8 + i * 4.05;
    disc(s, x, 2.15, String(i + 1), ICE === ICE ? SLATE : SLATE, 0.42);
    s.addText(h, { x: x + 0.6, y: 2.13, w: 3.2, h: 0.4, fontSize: 19, bold: true, color: WHITE,
      fontFace: SANS, isTextBox: true, margin: 0 });
    s.addText(items.map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < items.length - 1 } })),
      { x, y: 2.85, w: 3.6, h: 2.8, fontSize: 13, color: ICE, fontFace: SANS, isTextBox: true,
        margin: 0, paraSpaceAfter: 10 });
  });
  s.addText('Owners, dates and participants are set at kick-off.', { x: 0.8, y: 6.4, w: 7, h: 0.4,
    fontSize: 12, italic: true, color: '9AA7C4', fontFace: SANS, isTextBox: true, margin: 0 });
  s.addNotes('Close on two asks: confirm the principles, and nominate the workstream owners.');
}

pres.writeFile({ fileName: '/home/user/bae/FSI strategy/EAIND-FSI-Strategy-Development-Plan.pptx' })
  .then(f => console.log('wrote', f));
