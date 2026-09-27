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
  s.addNotes(`CONTEXT FOR THE TEAM — do not put on the slide.

The portfolio is EAIND: Engineering, AI and Data. It is a consulting portfolio — think technology advisory and delivery. We design, advise on and implement end-to-end technology transformation for the financial services industry, primarily banking.

Three pillars:
1. Industry Solutions — core-led transformation: core modernisation, core banking and other core packages. Also full systems-integration capability: we can take the SI role and own the end-to-end setup, including design governance, functional governance, delivery management and individual stream management, to stand up a greenfield bank or modernise an incumbent.
2. Data and AI — enterprise AI, core AI, AI use-case activation, and standard data work: data platforms, warehouses, ETL, reporting, all for banking.
3. Engineering and Cloud — cloud advisory and implementation, plus custom engineered solutions: integration management, microservices, API gateways, and bespoke core engineering. Less common in banking; likeliest in investments and wealth.

This deck is the plan for developing the FSI strategy, not the strategy. Target: strategy finalised within three weeks.`);
}

/* 2 — Outline */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, 'What is in this pack', 'Ten sections: the framing, the dimensions, the plan.');
  const items = [
    ['What this document is', 'Method, not conclusion'],
    ['Eight questions to answer', 'The dimensions the strategy must close'],
    ['Principles we hold going in', 'Constraints agreed before the work starts'],
    ['How we will get there', 'The five-step approach'],
    ['Eight workstreams, eight artefacts', 'Who owns what, and what each produces'],
    ['Team and talent', 'Dedicated, shared and the gap to close'],
    ['Eminence and positioning', 'How the market comes to know us'],
    ['Alliances', 'Who we go to market with, and for what'],
    ['The three-week journey', 'Week by week, with a gate each week'],
    ['How we run it', 'Cadence, inputs and decisions requested'],
  ];
  items.forEach(([t, d], i) => {
    const x = 0.7 + (i % 2) * 6.15, y = 1.95 + Math.floor(i / 2) * 0.93;
    s.addShape(pres.ShapeType.roundRect, { x, y, w: 5.75, h: 0.78, rectRadius: 0.08,
      fill: { color: i % 2 ? MIST : 'F2F6FC' } });
    disc(s, x + 0.22, y + 0.17, String(i + 1), i % 2 ? SLATE : NAVY, 0.44);
    s.addText(t, { x: x + 0.85, y: y + 0.08, w: 4.7, h: 0.34, fontSize: 15, bold: true, color: INK,
      fontFace: SANS, isTextBox: true, margin: 0 });
    s.addText(d, { x: x + 0.85, y: y + 0.42, w: 4.7, h: 0.3, fontSize: 11.5, color: MUTE,
      fontFace: SANS, isTextBox: true, margin: 0 });
  });
  s.addNotes(`Walk the pack in two halves. Sections 1 to 5 are the method: what this document is, the questions, the principles, the approach and the workstreams. Sections 6 to 8 are the dimensions that were called out as needing their own treatment — talent, eminence and alliances. Sections 9 and 10 are the plan and the governance.

If time is short, the three slides that carry the decision are: the eight questions, the five-step approach, and how we run it.`);
}

/* 3 — How to use */
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
  s.addNotes(`Direction given: this pack carries no answers. We know things already — focus markets, likely accounts — but none of it belongs here. Everything is a blank to be filled through the cycle, so reviewers engage with the method first and do not anchor on half-formed conclusions.

Still open before we finalise this pack:
- Audience: leadership approval pack, or team working document?
- What "partner" means where we say account ownership: our own consulting partners (the people who carry accounts) versus alliance and vendor partners. Both matter; they are different lists.
- Start date for the three-week cycle.`);
}

/* 3 — Framework */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, 'Eight questions to answer', 'Each dimension is owned by one workstream and closed at sign-off.');
  const dims = [
    ['Where to play', ['Which markets, in what priority?', 'Which client segments?', 'Which named accounts?']],
    ['What we offer', ['Which integrated propositions across the three pillars?', 'Which partial offerings are permitted?']],
    ['How we win', ['What differentiates us?', 'Which alliances and credentials are needed?']],
    ['How much', ['What revenue by market and by offering?', 'What deal sizes and price bands does the market pay?']],
    ['Who covers what', ['Which partner owns which account?', 'What coverage model across markets?']],
    ['Who delivers it', ['Who is dedicated to FSI?', 'What shared capacity can we draw on?', 'What must we recruit?']],
    ['How we are known', ['Where must we be visible?', 'What do we publish, host or speak at?', 'How do we measure it?']],
    ['Who we go with', ['Which alliances matter for which offering?', 'What do we leverage each one for?', 'What do they expect back?']],
  ];
  dims.forEach(([t, qs], i) => {
    const x = 0.7 + (i % 4) * 3.0, y = 2.0 + Math.floor(i / 4) * 2.15;
    s.addShape(pres.ShapeType.roundRect, { x, y, w: 2.85, h: 1.95, rectRadius: 0.12,
      fill: { color: i % 2 ? MIST : 'E4EBF7' } });
    disc(s, x + 0.25, y + 0.28, String(i + 1), NAVY, 0.4);
    s.addText(t, { x: x + 0.78, y: y + 0.26, w: 1.95, h: 0.4, fontSize: 14.5, bold: true, color: INK,
      fontFace: SANS, isTextBox: true, margin: 0 });
    s.addText(qs.map((q, j) => ({ text: q, options: { bullet: true, breakLine: j < qs.length - 1 } })),
      { x: x + 0.28, y: y + 0.78, w: 2.4, h: 1.05, fontSize: 10.5, color: BODY, fontFace: SANS,
        isTextBox: true, margin: 0, paraSpaceAfter: 4 });
  });
  openTag(s, 4.65, 6.6, 4.0, 'answers ' + TBD + ' in the cycle');
  s.addNotes(`Framing is loosely a strategy cascade (where to play, how to play), extended for a consulting portfolio with commercial, coverage, talent and eminence dimensions.

Context we already hold, to be validated rather than assumed during the cycle:
- Primary markets: Saudi Arabia, UAE, Qatar.
- Secondary markets: Kuwait, Oman, Jordan — where we focus our key relationships.
- The output must go further than markets: a specific, named target account list, with account mapping — who covers whom.
- Offerings must be expressed as products and services with a buyer attached: who buys this, and what do they typically pay.`);
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
  s.addNotes(`The hard rule from the portfolio: no siloed solutions. Our primary go-to-market catalogue is integrated propositions across the three pillars. Even where we sell a focused solution — enterprise AI, for example — it is sold as an entry point into an integrated play, and the catalogue is built that way. Flavours and partial offerings are permitted as variants of an integrated proposition, not as the default.

Delivery model: we expect to inherit the standard firm delivery model rather than invent one. The work here is to confirm it and note where FSI-specific governance (design authority, functional governance, stream management on large SI programmes) needs to be explicit.

Confirm all five principles at kick-off; they bound every later decision.`);
}

/* 5 — Approach */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, 'How we will get there', 'A five-step approach, run once, end to end.');
  const steps = [
    ['Mobilise', ['Confirm scope, principles and owners', 'Agree inputs and cadence'], 'Week 1'],
    ['Baseline', ['Credentials, pipeline and current wins', 'Who we have today: dedicated and shared'], 'Week 1'],
    ['Define', ['Markets, segments and candidate accounts', 'Integrated proposition catalogue'], 'Weeks 1–2'],
    ['Quantify', ['Deal sizes and market price bands', 'Revenue targets and the capacity to service them'], 'Week 2'],
    ['Commit', ['Account ownership and coverage', 'Team, eminence and activation plans signed off'], 'Week 3'],
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
  s.addNotes(`The five steps are the process view; the workstream and calendar views that follow are the same journey cut by owner and by week.

Note on Quantify: pricing will not reduce to a single number per offering. Core banking modernisation and migration varies by retail, corporate or SME scope; by single-country versus multi-country; and by whether full data migration is in scope. So the artefact is price bands and deal-size ranges with the drivers named, not a price list. Same logic for AI and data platform work, and for cloud and integration engagements.

Baseline also covers people: who we have today, dedicated and shared.`);
}

/* 6 — Workstreams */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, 'Eight workstreams, eight artefacts', 'Each workstream owns one dimension and produces one deliverable.');
  const rows = [
    [{ text: 'Workstream', options: { bold: true } }, { text: 'Question it closes', options: { bold: true } }, { text: 'Artefact produced', options: { bold: true } }],
    ['1 · Proposition & catalogue', 'What we offer', 'Integrated proposition catalogue with permitted variants'],
    ['2 · Market & account targeting', 'Where to play', 'Prioritised markets, segments and named account list'],
    ['3 · Commercial sizing & pricing', 'How much', 'Deal-size bands, price benchmarks, revenue targets'],
    ['4 · Coverage & ownership', 'Who covers what', 'Account-to-partner map and coverage model'],
    ['5 · Delivery model & readiness', 'How we win', 'Delivery approach and delivery standards'],
    ['6 · Team & talent', 'Who delivers it', 'Dedicated FSI team, shared capacity view, recruitment plan'],
    ['7 · Eminence & positioning', 'How we are known', 'Eminence plan: content, events and speaking calendar'],
    ['8 · Alliances', 'Who we go with', 'Alliance map: role, offerings and accounts per partner'],
  ];
  s.addTable(rows, {
    x: 0.7, y: 2.0, w: 11.9, colW: [3.3, 2.6, 6.0], fontSize: 12.5, fontFace: SANS, color: BODY,
    border: { type: 'solid', color: 'D8E0EE', pt: 1 }, align: 'left', valign: 'middle',
    rowH: [0.46, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5], margin: 6, fontSize: 11.5,
    fill: { color: WHITE },
  });
  s.addShape(pres.ShapeType.rect, { x: 0.7, y: 2.0, w: 11.9, h: 0.48, fill: { color: NAVY } });
  s.addText([{ text: 'Workstream', options: { bold: true } }],
    { x: 0.85, y: 2.0, w: 3.2, h: 0.48, valign: 'middle', fontSize: 12.5, color: WHITE, fontFace: SANS, isTextBox: true, margin: 0 });
  s.addText([{ text: 'Question it closes', options: { bold: true } }],
    { x: 4.15, y: 2.0, w: 2.5, h: 0.48, valign: 'middle', fontSize: 12.5, color: WHITE, fontFace: SANS, isTextBox: true, margin: 0 });
  s.addText([{ text: 'Artefact produced', options: { bold: true } }],
    { x: 6.75, y: 2.0, w: 5.6, h: 0.48, valign: 'middle', fontSize: 12.5, color: WHITE, fontFace: SANS, isTextBox: true, margin: 0 });
  s.addText('Workstreams run in parallel; dependencies are resolved at the weekly checkpoint.',
    { x: 0.7, y: 6.6, w: 11.9, h: 0.4, fontSize: 12.5, italic: true, color: MUTE, fontFace: SANS, isTextBox: true, margin: 0 });
  s.addNotes(`Artefact templates are blank at this stage and are populated during the cycle.

Notes per workstream:
- WS1 Proposition and catalogue: integrated propositions first, permitted variants second. Must cover all three pillars and the cross-pillar plays.
- WS2 Market and account targeting: markets are broadly known (primary: Saudi, UAE, Qatar; secondary: Kuwait, Oman, Jordan). The real output is the named account list and the segment logic behind it.
- WS3 Commercial sizing and pricing: size of each product, who buys it, and the typical price the market pays — with the variation drivers spelled out (retail/corporate/SME, single/multi-country, data migration in or out of scope).
- WS4 Coverage and ownership: two distinct mappings — our partners to accounts (who owns the relationship) and alliance partners to accounts (who we go to market with, and what they drive).
- WS5 Delivery model and readiness: largely inherited; confirm and note gaps.
- WS6 Team and talent: see the talent slide.
- WS7 Eminence and positioning: see the eminence slide.
- WS8 Alliances: see the alliances slide. Primary alliances today are Oracle, Google and Amazon; specialised alliances are Temenos and Intellect.`);
}

/* 7 — Team and talent */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, 'Team and talent', 'The capacity question, answered in three parts.');
  const cols = [
    ['The dedicated team', NAVY, ['Who works on FSI exclusively, across the three pillars?', 'Which roles and levels do they hold?', 'What is realistically deliverable with them alone?']],
    ['Shared capacity', SLATE, ['Which resources are shared with other portfolios?', 'How much of their time can we count on?', 'Which skills do we borrow rather than own?']],
    ['The gap to close', NAVY, ['How many partners, directors and delivery staff are missing?', 'Build, borrow or buy for each gap?', 'By when, to support the targets we set?']],
  ];
  cols.forEach(([t, c, qs], i) => {
    const x = 0.7 + i * 4.07;
    s.addShape(pres.ShapeType.roundRect, { x, y: 2.0, w: 3.75, h: 3.7, rectRadius: 0.12, fill: { color: MIST } });
    s.addShape(pres.ShapeType.rect, { x, y: 2.0, w: 3.75, h: 0.72, fill: { color: c } });
    s.addText(t, { x: x + 0.3, y: 2.0, w: 3.15, h: 0.72, valign: 'middle', fontSize: 17, bold: true,
      color: WHITE, fontFace: SANS, isTextBox: true, margin: 0 });
    s.addText(qs.map((q, j) => ({ text: q, options: { bullet: true, breakLine: j < qs.length - 1 } })),
      { x: x + 0.3, y: 2.95, w: 3.15, h: 2.5, fontSize: 12.5, color: BODY, fontFace: SANS,
        isTextBox: true, margin: 0, paraSpaceAfter: 10 });
  });
  openTag(s, 3.9, 5.95, 5.5, 'names, numbers and dates ' + TBD + ' in the cycle');
  s.addText('Talent is treated as a strategy output, not an afterthought: targets we cannot staff are not targets.',
    { x: 0.7, y: 6.55, w: 11.9, h: 0.4, fontSize: 13, italic: true, color: SLATE,
      fontFace: SANS, isTextBox: true, margin: 0 });
  s.addNotes(`This was called out as an important dimension in its own right.

Three things to establish:
1. The dedicated team — who in the portfolio works on FSI exclusively, across Industry Solutions, Data and AI, and Engineering and Cloud. Roles, levels, and what we can credibly deliver with them alone.
2. Generic and shared capacity — who we can draw on from the wider portfolio and firm, how much of their time we can actually count on, and which skills we borrow rather than own.
3. The recruitment gap — what remains once dedicated and shared capacity is counted, expressed by level: partners, directors, and other delivery staff. Then build, borrow or buy, and by when.

The link to the rest of the strategy is direct: a revenue target we cannot staff is not a target. Capacity is sized against the targets set in the commercial workstream.`);
}

/* 8 — Eminence */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, 'Eminence and positioning', 'How the market comes to know us for this.');
  const ch = [
    ['Published views', ['Points of view', 'Articles', 'Benchmark reports']],
    ['Stages', ['Industry conferences', 'Flagship FSI events', 'Panels and keynotes']],
    ['Client formats', ['Executive workshops', 'Seminars', 'Roundtables']],
    ['Digital', ['Podcasts', 'Webinars', 'Social presence']],
    ['Partner-led', ['Joint content', 'Alliance events', 'Co-branded assets']],
  ];
  ch.forEach(([t, items], i) => {
    const x = 0.7 + i * 2.42;
    s.addShape(pres.ShapeType.roundRect, { x, y: 2.05, w: 2.22, h: 2.9, rectRadius: 0.12,
      fill: { color: i % 2 ? MIST : 'E4EBF7' } });
    disc(s, x + 0.25, 2.3, String(i + 1), SLATE, 0.4);
    s.addText(t, { x: x + 0.25, y: 2.82, w: 1.85, h: 0.4, fontSize: 15, bold: true, color: INK,
      fontFace: SANS, isTextBox: true, margin: 0 });
    s.addText(items.map((q, j) => ({ text: q, options: { bullet: true, breakLine: j < items.length - 1 } })),
      { x: x + 0.25, y: 3.35, w: 1.8, h: 1.4, fontSize: 11.5, color: BODY, fontFace: SANS,
        isTextBox: true, margin: 0, paraSpaceAfter: 6 });
  });
  s.addShape(pres.ShapeType.roundRect, { x: 0.7, y: 5.25, w: 11.9, h: 1.05, rectRadius: 0.1, fill: { color: NAVY } });
  s.addText('For each channel the plan must answer: which ones we commit to, at what frequency, who fronts them, which accounts they target, and how we measure the return.',
    { x: 1.0, y: 5.25, w: 11.3, h: 1.05, valign: 'middle', fontSize: 13.5, color: ICE,
      fontFace: SANS, isTextBox: true, margin: 0 });
  openTag(s, 4.65, 6.55, 4.0, 'channel mix ' + TBD + ' in the cycle');
  s.addNotes(`Raised as a distinct part of the plan: how we enhance our positioning in the market, specifically in FSI, around what we do.

Channels mentioned: workshops, seminars, podcasts, articles, and participation in key industry events — Money20/20 was named as an example of the class of event we should consider. The plan should decide which of these are relevant to us rather than assume all of them.

For each channel the plan answers: do we commit, how often, who fronts it, which accounts or markets it targets, and how we measure return — inbound enquiries, meetings created, inclusion on shortlists and RFP invitations.

Worth linking eminence to the account list: eminence activity should point at the accounts and markets the strategy has prioritised, not at general brand awareness.`);
}

/* 9 — Alliances */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, 'Alliances', 'Who we go to market with, and what we use each one for.');
  const groups = [
    ['Primary alliances', NAVY, ['Oracle', 'Google', 'Amazon'], 0.7, 5.6],
    ['Specialised alliances', SLATE, ['Temenos', 'Intellect'], 6.9, 5.7],
  ];
  groups.forEach(([t, c, names, x, w]) => {
    s.addShape(pres.ShapeType.roundRect, { x, y: 2.0, w, h: 2.25, rectRadius: 0.12, fill: { color: MIST } });
    s.addText(t, { x: x + 0.3, y: 2.15, w: w - 0.6, h: 0.4, fontSize: 17, bold: true, color: c,
      fontFace: SANS, isTextBox: true, margin: 0 });
    const cw = (w - 0.6 - 0.2 * (names.length - 1)) / names.length;
    names.forEach((n, j) => {
      const nx = x + 0.3 + j * (cw + 0.2);
      s.addShape(pres.ShapeType.roundRect, { x: nx, y: 2.7, w: cw, h: 1.15, rectRadius: 0.1, fill: { color: c } });
      s.addText(n, { x: nx, y: 2.7, w: cw, h: 1.15, align: 'center', valign: 'middle', fontSize: 17,
        bold: true, color: WHITE, fontFace: SANS, isTextBox: true, margin: 0 });
    });
  });
  openTag(s, 4.4, 4.4, 4.5, 'further alliances ' + TBD + ' in the cycle');
  s.addText('For every alliance the plan must answer:', { x: 0.7, y: 5.05, w: 11.9, h: 0.35,
    fontSize: 14, bold: true, color: INK, fontFace: SANS, isTextBox: true, margin: 0 });
  const qs = [
    ['What we leverage it for', 'Which propositions and which pillars it sits behind'],
    ['What the relationship is', 'Co-sell, resell, delivery partner or referral'],
    ['Where it opens doors', 'Which markets and which accounts it gives us access to'],
    ['What it costs us', 'Certifications, credentials and commitments they expect back'],
  ];
  qs.forEach(([h, b], i) => {
    const x = 0.7 + (i % 2) * 6.15, y = 5.5 + Math.floor(i / 2) * 0.75;
    disc(s, x, y, String(i + 1), i < 2 ? NAVY : SLATE, 0.34);
    s.addText(h, { x: x + 0.5, y: y - 0.04, w: 2.6, h: 0.35, fontSize: 13, bold: true, color: INK,
      fontFace: SANS, isTextBox: true, margin: 0 });
    s.addText(b, { x: x + 3.1, y: y - 0.04, w: 2.75, h: 0.45, fontSize: 12, color: BODY,
      fontFace: SANS, isTextBox: true, margin: 0 });
  });
  s.addNotes(`Alliances were called out as important in implementation and technology work.

Current alliances named: Oracle, Google and Amazon as the primary ones; Temenos and Intellect as specialised, banking-platform alliances. They are on the slide as current state, not as a conclusion — what we do with each is the open question.

The work is to decide, per alliance: what we leverage it for, which of our integrated propositions it sits behind, what kind of relationship it is (co-sell, resell, delivery, referral), which markets and accounts it opens, and what it demands from us in certifications, trained staff and pipeline commitments.

Two links to the rest of the strategy: the alliance map must reconcile with the account coverage map (an alliance that opens an account should show up against that account), and with the talent plan (certifications and accredited staff are a recruitment and training cost).

If other alliances belong here — hyperscaler or platform vendors we do not yet work with, or regional players — the cycle is where we add them.`);
}

/* 10 — Three-week journey */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, 'The three-week journey', 'Three weeks, three gates, one decision.');
  const weeks = [
    ['Week 1', 'Frame', ['Baseline current position and credentials', 'Prioritise markets and segments', 'Build the account longlist', 'Draft the proposition catalogue', 'Baseline the team we have today'], 'Gate: agreed market and segment priorities'],
    ['Week 2', 'Define', ['Lock propositions and permitted variants', 'Benchmark deal sizes and pricing', 'Shortlist target accounts', 'Map alliances to offerings', 'Size the capacity each target implies'], 'Gate: agreed catalogue and shortlist'],
    ['Week 3', 'Commit', ['Set revenue targets by market and offering', 'Assign an owner to every account', 'Agree the team model and recruitment plan', 'Set the eminence calendar', 'Build the activation plan and sign off'], 'Gate: strategy signed off'],
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
  s.addNotes(`Three weeks is the timebox given: the strategy should be finalised in no more than about three weeks. The constraint is deliberate — it forces decisions rather than analysis.

Each week ends at a gate, and a gate is a decision taken, not a document circulated. If a gate cannot be closed, we escalate at the checkpoint rather than let the week slip.

Start date to be set at kick-off.`);
}

/* 11 — Governance */
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
    ['Decisions requested', ['Priority markets and segments', 'Target account list', 'Revenue and margin targets', 'Account ownership', 'Alliance priorities', 'Investment and hiring asks']],
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
  s.addNotes(`Close on two asks: confirm the principles, and nominate the workstream owners.

Inputs are the long pole — pipeline, revenue by pillar and market, existing relationships, alliance status, delivery capacity and historical deal values all sit with different people. Chase them in the first 48 hours or week 1 slips.

Decisions requested at sign-off: priority markets and segments, the target account list, revenue and margin targets, account ownership, and the investment and hiring asks coming out of the talent and eminence workstreams.

Open items to resolve at kick-off: audience for this pack, the definition of partner for account ownership, and the start date.`);
}

pres.writeFile({ fileName: '/home/user/bae/FSI strategy/EAIND-FSI-Strategy-Development-Plan.pptx' })
  .then(f => console.log('wrote', f));
