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

The portfolio is EAIND: Engineering, AI and Data. It is a consulting portfolio — think technology advisory and delivery. We design, advise on and implement end-to-end technology transformation for the financial services industry. FSI here means three segments, in priority order: banking first, then funds (sovereign wealth funds, national and development funds, fund and investment management), then insurance.

Three pillars:
1. Industry Solutions — core-led transformation: core modernisation and core platform packages, whatever the segment calls its core (core banking, policy administration and claims in insurance, portfolio, fund accounting and investment platforms in funds). Also full systems-integration capability: we can take the SI role and own the end-to-end setup, including design governance, functional governance, delivery management and individual stream management, to stand up a greenfield institution or modernise an incumbent.
2. Data and AI — enterprise AI, core AI, AI use-case activation, and standard data work: data platforms, warehouses, ETL, reporting, across all three segments.
3. Engineering and Cloud — cloud advisory and implementation, plus custom engineered solutions: integration management, microservices, API gateways, and bespoke core engineering. Less common around the core; likeliest in investments, wealth and fund operations.

This deck is the plan for developing the FSI strategy, not the strategy. Target: strategy finalised within three weeks.`);
}

/* 2 — Outline */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, 'What is in this pack', 'Sixteen sections, plus an appendix of blank templates.');
  const items = [
    ['What this document is', 'Method, not conclusion'],
    ['What the strategy must cover', 'Ambition, four pillars, operating discipline'],
    ['The industry we mean', 'Banking, funds and insurance, in priority order'],
    ['Where we start: the vision', 'Three candidate statements to deliberate'],
    ['Pillar 1 · Tier one and tier two', 'Where the energy goes'],
    ['Principles we hold going in', 'Constraints agreed before the work starts'],
    ['How we will get there', 'The five-step approach'],
    ['Five workstreams, one per pillar', 'Who owns what, and what each produces'],
    ['Who is doing this', 'The strategy development team'],
    ['Pillar 3 · Delivery and pricing economics', 'Construct, price, ADR and margin'],
    ['Pillar 3 · Team and talent', 'Dedicated, shared and the gap to close'],
    ['Pillar 4 · Eminence and positioning', 'How the market comes to know us'],
    ['Pillar 4 · Alliances', 'Who we go to market with, and for what'],
    ['The three-week journey', 'Week by week, with a gate each week'],
    ['The discipline · Governing the execution', 'Rhythm and scorecard after sign-off'],
    ['How we run the cycle', 'Cadence, inputs and decisions requested'],
  ];
  items.forEach(([t, d], i) => {
    const x = 0.7 + (i % 2) * 6.15, y = 1.9 + Math.floor(i / 2) * 0.72;
    s.addShape(pres.ShapeType.roundRect, { x, y, w: 5.75, h: 0.63, rectRadius: 0.08,
      fill: { color: i % 2 ? MIST : 'F2F6FC' } });
    disc(s, x + 0.2, y + 0.11, String(i + 1), i % 2 ? SLATE : NAVY, 0.4);
    s.addText(t, { x: x + 0.78, y: y + 0.03, w: 4.8, h: 0.32, fontSize: 13.5, bold: true, color: INK,
      fontFace: SANS, isTextBox: true, margin: 0 });
    s.addText(d, { x: x + 0.78, y: y + 0.33, w: 4.8, h: 0.28, fontSize: 10.5, color: MUTE,
      fontFace: SANS, isTextBox: true, margin: 0 });
  });
  s.addText('Appendix: a blank template for every artefact the cycle produces — the strategy log itself.',
    { x: 0.7, y: 6.95, w: 11.9, h: 0.4, fontSize: 12, italic: true, color: SLATE, fontFace: SANS,
      isTextBox: true, margin: 0 });
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
  head(s, 'What the strategy must cover', 'The shape of the strategy we are going to build \u2014 not the strategy itself.');
  s.addShape(pres.ShapeType.roundRect, { x: 0.7, y: 1.9, w: 11.9, h: 0.78, rectRadius: 0.1, fill: { color: NAVY } });
  s.addText('THE AMBITION', { x: 1.0, y: 1.9, w: 2.6, h: 0.78, valign: 'middle', fontSize: 13, bold: true,
    color: ICE, charSpacing: 1.5, fontFace: SANS, isTextBox: true, margin: 0 });
  s.addText('What winning looks like, in numbers, by when', { x: 3.7, y: 1.9, w: 8.6, h: 0.78,
    valign: 'middle', fontSize: 15, color: WHITE, fontFace: SANS, isTextBox: true, margin: 0 });
  const pillars = [
    ['Market\nand clients', ['Markets and FSI segments', 'Tier one and tier two', 'Named target accounts', 'Offerings and revenue per account', 'Coverage and ownership']],
    ['Offerings\n(what we sell)', ['Integrated propositions', 'Use cases we lead with', 'Flavours and permitted variants']],
    ['Delivery\nand economics', ['Delivery model and construct', 'Pricing, ADR and margin', 'Team, capacity and recruitment', 'Revenue targets']],
    ['Route\nto market', ['Alliances and what they unlock', 'Eminence and positioning', 'How we differentiate']],
  ];
  pillars.forEach(([t, items], i) => {
    const x = 0.7 + i * 3.02, w = 2.84;
    s.addShape(pres.ShapeType.roundRect, { x, y: 2.85, w, h: 3.25, rectRadius: 0.12,
      fill: { color: i % 2 ? MIST : 'E4EBF7' } });
    s.addShape(pres.ShapeType.rect, { x, y: 2.85, w, h: 0.92, fill: { color: i % 2 ? SLATE : NAVY } });
    s.addText(t, { x: x + 0.22, y: 2.85, w: w - 0.44, h: 0.92, valign: 'middle', fontSize: 15, bold: true,
      color: WHITE, fontFace: SANS, isTextBox: true, margin: 0, lineSpacing: 17 });
    s.addText(items.map((q, j) => ({ text: q, options: { bullet: true, breakLine: j < items.length - 1 } })),
      { x: x + 0.22, y: 3.95, w: w - 0.42, h: 2.05, fontSize: 11.5, color: BODY, fontFace: SANS,
        isTextBox: true, margin: 0, paraSpaceAfter: 7 });
  });
  s.addShape(pres.ShapeType.roundRect, { x: 0.7, y: 6.25, w: 11.9, h: 0.72, rectRadius: 0.1, fill: { color: NAVY } });
  s.addText('THE DISCIPLINE', { x: 1.0, y: 6.25, w: 2.6, h: 0.72, valign: 'middle', fontSize: 13, bold: true,
    color: ICE, charSpacing: 1.5, fontFace: SANS, isTextBox: true, margin: 0 });
  s.addText('Operating rhythm, scorecard and decision rights that keep the four pillars honest',
    { x: 3.7, y: 6.25, w: 8.6, h: 0.72, valign: 'middle', fontSize: 14, color: WHITE, fontFace: SANS,
      isTextBox: true, margin: 0 });
  s.addNotes(`Framing is loosely a strategy cascade (where to play, how to play), extended for a consulting portfolio with commercial, coverage, talent and eminence dimensions.

Context we already hold, to be validated rather than assumed during the cycle:
- Primary markets: Saudi Arabia, UAE, Qatar.
- Secondary markets: Kuwait, Oman, Jordan — where we focus our key relationships.
- The output must go further than markets: a specific, named target account list, with account mapping — who covers whom.
- Offerings must be expressed as products and services with a buyer attached: who buys this, and what do they typically pay.`);
}

/* 4 — Segments */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, 'The industry we mean', 'Financial services is three segments, and they are not interchangeable.');
  const segs = [
    ['1', 'Banking', 'Priority segment', ['Retail, corporate and SME banking', 'Islamic banking', 'Payments and cards', 'Greenfield and challenger banks'], NAVY],
    ['2', 'Funds', 'Second priority', ['Sovereign wealth funds', 'National and development funds', 'Fund and investment management', 'Investment operations and custody'], SLATE],
    ['3', 'Insurance', 'Third priority', ['Life and general insurance', 'Takaful', 'Health insurance', 'Brokers and reinsurance'], NAVY],
  ];
  segs.forEach(([k, name, rank, items, c], i) => {
    const x = 0.7 + i * 4.03;
    s.addShape(pres.ShapeType.roundRect, { x, y: 2.0, w: 3.85, h: 3.5, rectRadius: 0.12, fill: { color: MIST } });
    s.addShape(pres.ShapeType.rect, { x, y: 2.0, w: 3.85, h: 0.95, fill: { color: c } });
    s.addText(name, { x: x + 0.3, y: 2.08, w: 2.7, h: 0.42, fontSize: 19, bold: true, color: WHITE,
      fontFace: SANS, isTextBox: true, margin: 0 });
    s.addText(rank, { x: x + 0.3, y: 2.5, w: 2.7, h: 0.35, fontSize: 11.5, color: ICE,
      fontFace: SANS, isTextBox: true, margin: 0 });
    s.addText(k, { x: x + 3.0, y: 2.1, w: 0.6, h: 0.7, align: 'right', fontSize: 30, bold: true,
      color: 'FFFFFF', fontFace: HEAD, isTextBox: true, margin: 0 });
    s.addText(items.map((q, j) => ({ text: q, options: { bullet: true, breakLine: j < items.length - 1 } })),
      { x: x + 0.3, y: 3.15, w: 3.25, h: 2.1, fontSize: 12, color: BODY, fontFace: SANS,
        isTextBox: true, margin: 0, paraSpaceAfter: 8 });
  });
  s.addShape(pres.ShapeType.roundRect, { x: 0.7, y: 5.7, w: 11.9, h: 1.0, rectRadius: 0.1, fill: { color: NAVY } });
  s.addText('Segment is a dimension throughout: every proposition says which segments it serves, every account carries its segment, and targets are set by market and segment, not for FSI as a whole.',
    { x: 1.0, y: 5.7, w: 11.3, h: 1.0, valign: 'middle', fontSize: 13, color: ICE, fontFace: SANS,
      isTextBox: true, margin: 0 });
  s.addText('Sub-segments are indicative; the cycle confirms which we actually pursue.',
    { x: 0.7, y: 6.8, w: 11.9, h: 0.35, fontSize: 11, italic: true, color: MUTE, fontFace: SANS,
      isTextBox: true, margin: 0 });
  s.addNotes(`Correcting the frame: this is a financial services strategy, not a banking strategy. Three segments, in this priority order.

1. Banking — the priority segment, and where most of the existing capability and credentials sit.
2. Funds — sovereign wealth funds, national and development funds, and fund or investment management. Often a small number of very large institutions rather than a market of many, which changes how coverage and pursuit work: fewer accounts, longer cycles, far higher stakes per relationship.
3. Insurance — life, general, takaful, health, brokers and reinsurance. Different core platforms, different regulators, different buyers.

They are not interchangeable. The core platform, the regulator, the buyer and the language all change between them, which is exactly why the domain-specific principle is stated per segment rather than per industry: a core banking reference does not win a policy administration programme.

Practical consequence for the cycle: segment is a column in the market prioritisation, the account list, the proposition catalogue and the revenue targets. Where we are thin — and funds and insurance are where we are thin — the strategy has to say whether we build, borrow, hire or partner into them, and the alliance map has to answer the platform question for each.`);
}

/* 5 — Vision options */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, 'Where we start: the vision', 'All three carry the integrated, end-to-end theme. They differ on the edge we claim on top of it.');
  const opts = [
    ['A', 'The end-to-end partner', 'One partner for everything technology in financial services: whatever a bank, a fund or an insurer needs designed, built, modernised or run, we bring it end to end \u2014 our capabilities and our alliances working as one team.',
     'Edge: breadth and single accountability.', 'Risk: breadth alone is claimable by every large SI.'],
    ['B', 'AI-native, end to end', 'We make financial institutions AI-native: core, data, engineering and alliances brought as one team, so intelligence runs through the institution rather than beside it \u2014 and one partner is accountable for all of it.',
     'Edge: AI as the through-line of everything we sell.', 'Risk: must be provable in every proposition, or it is a slogan.'],
    ['C', 'Known for the industry, built to execute', 'The team banks, funds and insurers call first: we know their business as deeply as their technology, and we bring everything needed to execute \u2014 ours and our partners\u2019 \u2014 as one team.',
     'Edge: domain depth plus end-to-end delivery.', 'Risk: quieter claim, harder to headline.'],
  ];
  opts.forEach(([k, name, text, bet, risk], i) => {
    const x = 0.7 + i * 4.03;
    s.addShape(pres.ShapeType.roundRect, { x, y: 2.0, w: 3.85, h: 3.95, rectRadius: 0.12,
      fill: { color: i % 2 ? MIST : 'E4EBF7' } });
    disc(s, x + 0.3, 2.25, k, i % 2 ? SLATE : NAVY, 0.46);
    s.addText(name, { x: x + 0.95, y: 2.22, w: 2.7, h: 0.55, valign: 'middle', fontSize: 15.5, bold: true,
      color: INK, fontFace: SANS, isTextBox: true, margin: 0 });
    s.addText('\u201c' + text + '\u201d', { x: x + 0.3, y: 2.95, w: 3.25, h: 2.0, fontSize: 11.5,
      italic: true, color: NAVY, fontFace: SANS, isTextBox: true, margin: 0 });
    s.addText(bet, { x: x + 0.3, y: 5.0, w: 3.25, h: 0.5, fontSize: 11.5, bold: true, color: BODY,
      fontFace: SANS, isTextBox: true, margin: 0 });
    s.addText(risk, { x: x + 0.3, y: 5.45, w: 3.25, h: 0.45, fontSize: 10.5, color: MUTE,
      fontFace: SANS, isTextBox: true, margin: 0 });
  });
  s.addShape(pres.ShapeType.roundRect, { x: 0.7, y: 6.1, w: 11.9, h: 0.9, rectRadius: 0.1, fill: { color: NAVY } });
  s.addText('Common to all three: one partner, end to end, our capabilities and our alliances as one team. The test is what we claim on top of that \u2014 and whether a competitor could lift it word for word.',
    { x: 1.0, y: 6.1, w: 11.3, h: 0.9, valign: 'middle', fontSize: 12.5, color: ICE, fontFace: SANS,
      isTextBox: true, margin: 0 });
  s.addNotes(`Deliberation material, not a recommendation. All three now carry the portfolio's core claim — integrated propositions, integrated teams, integrated capabilities, one partner accountable for execution end to end, with the alliances counted as part of what we bring rather than a separate arrangement. They differ on the edge claimed on top of that.

A. The end-to-end partner — leads with breadth and single accountability: whatever the institution needs designed, built, modernised or run. Closest to how the portfolio is already built and easiest to evidence. Weakest on differentiation, since scale is claimable by every large SI.

B. AI-native, end to end — keeps the end-to-end claim and adds AI as the through-line of everything we sell. Strongest differentiator available now and matched to regional ambition, but it obliges us to prove it: if the catalogue does not visibly carry AI in every proposition, the vision undercuts itself.

C. Known for the industry, built to execute — keeps the end-to-end claim and adds domain depth per segment, plus the ability to both advise and build. Quietest of the three and hardest to headline, but the one clients reward, and defensible against global SIs and product vendors alike. It obliges depth in funds and insurance, not only banking.

Wording to avoid: "one-stop shop" carries a commodity, lowest-common-denominator connotation. The same idea lands better as one partner, one accountable team, end to end.

A combination is plausible and may be where this lands — for example C as the base claim with B as the edge. Resist merging all three into a sentence that commits to nothing.

Test to apply: what we do, for whom, what makes us different, and whether a competitor could lift it word for word. If they could, it is positioning, not vision. Whichever is chosen, the ambition then puts numbers and a date against it.`);
}

/* 6 — Tiering */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, 'Tier one and tier two', 'Pillar 1 · where the energy goes, and where we stay present without over-investing.');
  const tiers = [
    ['Tier one', NAVY, 'Priority geographies and named accounts',
     ['Active, deliberate coverage — people walking the corridors', 'Named individual account targets, each with an owner', 'Individual revenue targets, gross and net', 'Where the metrics, the eminence and the investment concentrate']],
    ['Tier two', SLATE, 'Emerging and reactive',
     ['Managed as clusters and groups, not one by one', 'Reactive pursuit: we respond, we do not camp there', 'Light coverage — a named watcher rather than an owner', 'Promoted to tier one when the signal justifies it']],
  ];
  tiers.forEach(([t, c, sub, items], i) => {
    const x = 0.7 + i * 6.1;
    s.addShape(pres.ShapeType.roundRect, { x, y: 2.0, w: 5.8, h: 3.6, rectRadius: 0.12, fill: { color: MIST } });
    s.addShape(pres.ShapeType.rect, { x, y: 2.0, w: 5.8, h: 0.95, fill: { color: c } });
    s.addText(t, { x: x + 0.3, y: 2.08, w: 5.2, h: 0.42, fontSize: 20, bold: true, color: WHITE,
      fontFace: SANS, isTextBox: true, margin: 0 });
    s.addText(sub, { x: x + 0.3, y: 2.5, w: 5.2, h: 0.35, fontSize: 12, color: ICE, fontFace: SANS,
      isTextBox: true, margin: 0 });
    s.addText(items.map((q, j) => ({ text: q, options: { bullet: true, breakLine: j < items.length - 1 } })),
      { x: x + 0.3, y: 3.15, w: 5.2, h: 2.3, fontSize: 12.5, color: BODY, fontFace: SANS,
        isTextBox: true, margin: 0, paraSpaceAfter: 10 });
  });
  s.addShape(pres.ShapeType.roundRect, { x: 0.7, y: 5.8, w: 11.9, h: 0.95, rectRadius: 0.1, fill: { color: NAVY } });
  s.addText('Tiering applies to geographies and to accounts, across all three segments. Tier two keeps us present in what is emerging without diluting tier one.',
    { x: 1.0, y: 5.8, w: 11.3, h: 0.95, valign: 'middle', fontSize: 13, color: ICE, fontFace: SANS,
      isTextBox: true, margin: 0 });
  s.addText('Promotion and demotion criteria, and the tier one list itself, are set in the cycle.',
    { x: 0.7, y: 6.85, w: 11.9, h: 0.35, fontSize: 11, italic: true, color: MUTE, fontFace: SANS,
      isTextBox: true, margin: 0 });
  s.addNotes(`Tiering applies twice — to geographies and to accounts — and independently of segment. A tier one account can sit in a tier two market, and occasionally the reverse.

Tier one is where the energy goes: deliberate coverage with people physically present working the relationships, named individual targets each with an owner, individual gross and net revenue targets, and the eminence and investment concentrated there. The scorecard is read against tier one.

Tier two is deliberately lighter. Accounts and markets are handled as clusters or groups rather than one by one, pursuit is reactive rather than campaigned, and coverage is a named watcher rather than an accountable owner. It exists for two reasons: to keep us present in what is emerging, and to stop us over-concentrating everything in a handful of names.

Two things the cycle must define rather than leave to instinct: what promotes a tier two account or market to tier one — a real signal, not enthusiasm — and what minimum tier two coverage actually means, so it is not simply a list nobody reads.

Watch the balance. If almost everything lands in tier one, nothing is prioritised. If tier one is too narrow, the plan is fragile to one or two decisions going against us.`);
}

/* 7 — Principles */
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
    ['Domain-specific, not agnostic', 'Each FSI segment is specialised and clients expect it: propositions, credentials and CVs speak that segment\u2019s language.'],
    ['AI and innovation embedded', 'Not a separate offering: every proposition carries an AI and innovation component by design.'],
    ['One integrated market team', 'Every partner and director knows the whole catalogue, reads the signals, positions level one, then pulls in the specialist.'],
  ];
  pr.forEach(([h, b], i) => {
    const x = 0.7 + (i % 2) * 6.15, y = 1.85 + Math.floor(i / 2) * 1.28;
    s.addShape(pres.ShapeType.roundRect, { x, y, w: 5.75, h: 1.15, rectRadius: 0.1, fill: { color: MIST } });
    disc(s, x + 0.28, y + 0.22, String(i + 1), i % 2 ? SLATE : NAVY, 0.4);
    s.addText(h, { x: x + 0.85, y: y + 0.16, w: 4.7, h: 0.32, fontSize: 15.5, bold: true, color: INK,
      fontFace: SANS, isTextBox: true, margin: 0 });
    s.addText(b, { x: x + 0.85, y: y + 0.52, w: 4.7, h: 0.58, fontSize: 11.5, color: BODY,
      fontFace: SANS, isTextBox: true, margin: 0 });
  });
  s.addText('These are the rules the strategy must obey. Anything else remains open.',
    { x: 0.7, y: 6.95, w: 11.9, h: 0.4, fontSize: 12, italic: true, color: SLATE,
      fontFace: SANS, isTextBox: true, margin: 0 });
  s.addNotes(`Four rules carry the most weight here.

AI and innovation embedded: AI is not a separate line item we sell when asked. Every proposition in the catalogue carries an AI or innovation component by design — a core modernisation that lands with intelligent operations and automated testing, an integration programme that ships with AI-assisted engineering, a data platform that arrives with use cases already running. Two reasons: it is where client expectation is heading in this region, and it is what stops the catalogue reading like commodity systems integration. Practically it becomes a test applied to the catalogue in the offerings workstream: if a proposition has no AI or innovation content, say why, or change it.

Domain-specific, not agnostic: financial services is specialised enough that generic technology credentials do not travel, and the three segments do not transfer to each other either. A banking client expects us to know core banking, payments and regulatory reporting; a fund expects portfolio, fund accounting, valuations and investor reporting; an insurer expects policy administration, claims, underwriting and actuarial. Credibility is per segment, not per industry. Practically this means propositions are written in banking language and mapped to banking outcomes; credentials and case studies are FSI ones; CVs put domain experience forward; and where we borrow capability from a horizontal pool, it is fronted by people who know the domain. It also bounds what we take on: a generic engagement we could win anywhere is not automatically ours to chase.

One integrated market team: every partner and director in the portfolio is expected to know the full catalogue, not only their own practice. In a client conversation they should recognise the signal — a core replacement being considered, a data estate that cannot answer regulatory questions, an integration layer at end of life, an AI ambition with no platform underneath it — position the first level of our answer credibly, and then pull in the specialist lead. Nobody is expected to be expert in everything; everybody is expected to open the door rather than let it pass because it sits outside their pillar.

This is what makes the integrated-propositions principle real in practice rather than on paper. It also creates an enablement obligation: the proposition one-pagers in the appendix are the tool for it, and the cycle should say how the portfolio is trained on them and how often it is refreshed.

The hard rule from the portfolio: no siloed solutions. Our primary go-to-market catalogue is integrated propositions across the three pillars. Even where we sell a focused solution — enterprise AI, for example — it is sold as an entry point into an integrated play, and the catalogue is built that way. Flavours and partial offerings are permitted as variants of an integrated proposition, not as the default.

Delivery model: we expect to inherit the standard firm delivery model rather than invent one. The work here is to confirm it and note where FSI-specific governance (design authority, functional governance, stream management on large SI programmes) needs to be explicit.

Confirm all five principles at kick-off; they bound every later decision.`);
}

/* 5 — Approach */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, 'How we will get there', 'A five-step approach, run once, end to end.');
  const steps = [
    ['Mobilise', ['Choose the vision, then frame the ambition', 'Confirm scope, principles and owners'], 'Week 1'],
    ['Baseline', ['Credentials, pipeline and current wins', 'Who we have today: dedicated and shared'], 'Week 1'],
    ['Define', ['Markets, segments and candidate accounts', 'Integrated proposition catalogue'], 'Weeks 1–2'],
    ['Quantify', ['Price bands, ADR and target margin', 'Revenue per account, gross and net, against capacity'], 'Week 2'],
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

Note on Quantify: pricing will not reduce to a single number per offering. Core modernisation and migration varies by segment and by scope — retail, corporate or SME in banking; life versus general in insurance; fund accounting versus front-office in funds — and by single-country versus multi-country, and by whether full data migration is in scope. So the artefact is price bands and deal-size ranges with the drivers named, not a price list. Same logic for AI and data platform work, and for cloud and integration engagements.

Baseline also covers people: who we have today, dedicated and shared.`);
}

/* 6 — Workstreams */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, 'Five workstreams, one per pillar', 'Each pillar has one owner and a defined set of artefacts.');
  const ws = [
    ['1 · Market and clients', 'Markets, segments, tiering, named accounts, offerings and revenue per account, coverage',
     'Market prioritisation · target account list · account revenue plan · coverage map'],
    ['2 · Offerings', 'What we sell: integrated propositions, use cases and permitted flavours',
     'Proposition catalogue'],
    ['3 · Delivery and economics', 'Delivery model and construct, pricing, margin, team and targets',
     'Pricing bands · offering economics · team model · revenue targets'],
    ['4 · Route to market', 'Alliances, eminence and differentiation',
     'Alliance map · eminence calendar'],
    ['5 · Operating discipline', 'Rhythm, measures and decision rights after sign-off',
     'Operating rhythm · KPI scorecard · review calendar'],
  ];
  ws.forEach(([t, scope, art], i) => {
    const y = 2.05 + i * 0.95;
    s.addShape(pres.ShapeType.roundRect, { x: 0.7, y, w: 11.9, h: 0.82, rectRadius: 0.1,
      fill: { color: i % 2 ? MIST : 'F2F6FC' } });
    s.addText(t, { x: 1.0, y: y + 0.04, w: 4.0, h: 0.34, fontSize: 13.5, bold: true, color: NAVY,
      fontFace: SANS, isTextBox: true, margin: 0 });
    s.addText(scope, { x: 1.0, y: y + 0.4, w: 4.0, h: 0.36, fontSize: 10.5, color: MUTE,
      fontFace: SANS, isTextBox: true, margin: 0 });
    s.addText(art, { x: 5.3, y, w: 7.0, h: 0.82, valign: 'middle', fontSize: 12.5, color: BODY,
      fontFace: SANS, isTextBox: true, margin: 0 });
  });
  s.addText('Level 1 is the pillar and its owner. Level 2 is the topic, and every topic ends in one artefact in the appendix.',
    { x: 0.7, y: 6.9, w: 11.9, h: 0.4, fontSize: 12, italic: true, color: MUTE, fontFace: SANS,
      isTextBox: true, margin: 0 });
  s.addNotes(`The structure is two levels on purpose. Level 1 is the pillar — one owner, one accountable partner. Level 2 is the topic inside it, and each topic ends in exactly one artefact in the appendix.

What each pillar carries:
- Market and clients: markets and segments (primary expected to be Saudi, UAE and Qatar; secondary Kuwait, Oman and Jordan, all to be validated), the named account list, and the coverage map. Two mappings live here — our own partners who own the relationship, and the alliance partner behind the account.
- Offerings: the catalogue itself, and nothing else — integrated propositions, the use cases we lead with, and the permitted flavours or partial offerings. It stands alone because it is the What: everything else in the strategy either sells it, delivers it or prices it.
- Delivery and economics: the delivery model and the construct behind each offering price bands with the drivers that move them, blended ADR and target margin, and the revenue targets that fall out. Delivery model and pricing sit in one pillar because they determine each other: the construct sets the cost, the cost sets the achievable price and margin.
  Capability sits in the same pillar: the dedicated FSI team, the shared capacity we can genuinely draw on, and the recruitment gap by level. It belongs here because the construct and the margin only hold if the people behind them exist — a target we cannot staff is not a target.
- Route to market: alliances (Oracle, Google, Amazon as primary; Temenos and Intellect as specialised) and the eminence plan.
- Operating discipline: the rhythm, scorecard and decision rights that keep the rest alive after sign-off.

Dependencies to watch: revenue targets cannot be set without the capacity view, and the alliance map must reconcile with the coverage map.`);
}

/* 8 — The team */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, 'Who is doing this', 'The strategy development team.');
  s.addShape(pres.ShapeType.roundRect, { x: 0.7, y: 2.0, w: 11.9, h: 1.15, rectRadius: 0.1, fill: { color: NAVY } });
  disc(s, 1.05, 2.28, 'D', SLATE, 0.6);
  s.addText('Dany', { x: 1.85, y: 2.22, w: 3.2, h: 0.4, fontSize: 20, bold: true, color: WHITE,
    fontFace: SANS, isTextBox: true, margin: 0 });
  s.addText('Overall lead — owns the cycle, chairs the checkpoints, takes the strategy to sign-off',
    { x: 1.85, y: 2.66, w: 10.3, h: 0.4, fontSize: 13, color: ICE, fontFace: SANS, isTextBox: true, margin: 0 });
  const leads = [
    ['Tanay', 'Industry Solutions', 'Core-led transformation and core platforms across banking, funds and insurance, full SI'],
    ['Rachel', 'AI and Data', 'Enterprise and core AI, use-case activation, data platforms'],
    ['Clifford', 'Engineering', 'Cloud, integration, microservices, API and bespoke build'],
  ];
  leads.forEach(([n, role, scope], i) => {
    const x = 0.7 + i * 4.03;
    s.addShape(pres.ShapeType.roundRect, { x, y: 3.4, w: 3.85, h: 2.15, rectRadius: 0.12, fill: { color: MIST } });
    disc(s, x + 0.3, 3.7, n.charAt(0), i % 2 ? SLATE : NAVY, 0.52);
    s.addText(n, { x: x + 1.0, y: 3.68, w: 2.6, h: 0.4, fontSize: 19, bold: true, color: INK,
      fontFace: SANS, isTextBox: true, margin: 0 });
    s.addText(role, { x: x + 0.3, y: 4.35, w: 3.25, h: 0.35, fontSize: 14, bold: true, color: NAVY,
      fontFace: SANS, isTextBox: true, margin: 0 });
    s.addText(scope, { x: x + 0.3, y: 4.72, w: 3.25, h: 0.7, fontSize: 11.5, color: BODY, fontFace: SANS,
      isTextBox: true, margin: 0 });
  });
  s.addText('Pillar ownership across the five workstreams is assigned from this group at kick-off; every artefact carries one name.',
    { x: 0.7, y: 5.8, w: 11.9, h: 0.4, fontSize: 13, italic: true, color: SLATE, fontFace: SANS,
      isTextBox: true, margin: 0 });
  s.addText('Wider contributors — account partners, alliance leads, resourcing and finance — are pulled in per workstream rather than sitting on the core team.',
    { x: 0.7, y: 6.3, w: 11.9, h: 0.4, fontSize: 12, color: MUTE, fontFace: SANS, isTextBox: true, margin: 0 });
  s.addNotes(`Core team as agreed: Dany leads overall; Tanay leads Industry Solutions; Rachel leads AI and Data; Clifford leads Engineering.

Note the shape difference: the three leads represent the three practices, while the workstreams are cut by pillar, which is deliberate — the strategy is one portfolio strategy, not three practice strategies stapled together. Each pillar therefore needs a named owner drawn from this group, with the other two contributing their practice view into it.

A workable split to propose at kick-off: Dany takes the ambition and the operating discipline; the offerings catalogue is owned jointly but chaired by one lead, since integrated propositions cut across all three practices; market and clients sits with whoever carries the strongest regional relationships; delivery and economics sits with the lead whose practice has the largest delivery footprint. Confirm rather than assume.

Everything else — account partners, alliance leads, resourcing, finance — is pulled in per workstream and not standing members.`);
}

/* 9 — Delivery and pricing economics */
/* 9 — Pricing and delivery economics */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, 'Delivery and pricing economics', 'Pillar 3 · delivery model, price and margin are one decision, not three.');
  const chain = [
    ['Target price', 'What the market pays for this offering, in bands'],
    ['Delivery construct', 'The onshore, nearshore and offshore mix it is priced on'],
    ['Blended ADR', 'The average daily rate that construct implies'],
    ['Target margin', 'The gross margin we commit to for this offering'],
  ];
  chain.forEach(([t, d], i) => {
    const x = 0.7 + i * 3.05;
    s.addShape(pres.ShapeType.roundRect, { x, y: 2.0, w: 2.7, h: 1.75, rectRadius: 0.12,
      fill: { color: i % 2 ? MIST : 'E4EBF7' } });
    disc(s, x + 0.25, 2.22, String(i + 1), NAVY, 0.4);
    s.addText(t, { x: x + 0.25, y: 2.72, w: 2.2, h: 0.35, fontSize: 15, bold: true, color: INK,
      fontFace: SANS, isTextBox: true, margin: 0 });
    s.addText(d, { x: x + 0.25, y: 3.1, w: 2.25, h: 0.6, fontSize: 11.5, color: BODY, fontFace: SANS,
      isTextBox: true, margin: 0 });
    if (i < 3) s.addText('→', { x: x + 2.72, y: 2.65, w: 0.3, h: 0.4, align: 'center', fontSize: 18,
      color: MUTE, fontFace: SANS, isTextBox: true, margin: 0 });
  });
  s.addShape(pres.ShapeType.roundRect, { x: 0.7, y: 3.95, w: 11.9, h: 0.75, rectRadius: 0.1, fill: { color: NAVY } });
  s.addText('Set per offering and per segment: a core banking programme, a fund platform build and an AI activation carry different constructs, rates and margins.',
    { x: 1.0, y: 3.95, w: 11.3, h: 0.75, valign: 'middle', fontSize: 13, color: ICE, fontFace: SANS,
      isTextBox: true, margin: 0 });
  const qs = [
    ['What must stay onshore', 'Client-facing, regulatory and data-residency constrained roles'],
    ['What can move', 'Which roles run nearshore or offshore, and at what ratio'],
    ['What margin we hold', 'The target gross margin per offering, and the floor'],
    ['What happens below it', 'Approval route when a deal prices under the floor'],
  ];
  qs.forEach(([h, b], i) => {
    const x = 0.7 + (i % 2) * 6.15, y = 5.0 + Math.floor(i / 2) * 0.9;
    disc(s, x, y, String.fromCharCode(65 + i), i < 2 ? SLATE : NAVY, 0.36);
    s.addText(h, { x: x + 0.52, y: y - 0.04, w: 2.6, h: 0.35, fontSize: 13, bold: true, color: INK,
      fontFace: SANS, isTextBox: true, margin: 0 });
    s.addText(b, { x: x + 3.15, y: y - 0.04, w: 2.7, h: 0.55, fontSize: 11.5, color: BODY, fontFace: SANS,
      isTextBox: true, margin: 0 });
  });
  s.addText('Rates, ratios and margin targets are set in the cycle and recorded in the offering economics template.',
    { x: 0.7, y: 6.85, w: 11.9, h: 0.35, fontSize: 11.5, italic: true, color: MUTE, fontFace: SANS,
      isTextBox: true, margin: 0 });
  s.addNotes(`Costing and delivery model are treated as one topic because they determine each other: the price the market will pay only becomes a business if the delivery construct behind it produces the blended ADR and the gross margin we want.

The chain to work through, per offering: target price band, then the delivery construct it assumes (the split across onshore, nearshore and offshore or global delivery centres), then the blended average daily rate that construct implies, then the gross margin that falls out. If the margin is short, one of the first three has to change — usually the construct.

Nuance to carry: this varies sharply by offering and by segment. A core banking modernisation with heavy onshore functional and governance presence looks nothing like a fund platform implementation, an insurance policy-administration migration, or an AI use-case activation that can run largely offshore. Regulatory and data-residency rules in Saudi and the UAE constrain what can leave the country, and some clients contractually require onshore staffing — both drive the construct before commercial preference does.

The delivery model itself is largely inherited from the firm; what we are setting here is the construct and the economics per offering, plus the margin floor and the approval route when a deal prices below it.`);
}

/* 8 — Team and talent */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, 'Team and talent', 'Pillar 3 · the capacity behind the construct, in three parts.');
  const cols = [
    ['The dedicated team', NAVY, ['Who works on FSI exclusively, across the three practices?', 'Which roles and levels do they hold?', 'What is realistically deliverable with them alone?']],
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
  head(s, 'Eminence and positioning', 'Pillar 4 · how the market \u2014 outside and inside the firm \u2014 comes to know us for this.');
  const ch = [
    ['Published views', ['Points of view', 'Articles', 'Benchmark reports']],
    ['Stages', ['Industry conferences', 'Flagship FSI events', 'Panels and keynotes']],
    ['Client formats', ['Executive workshops', 'Seminars', 'Roundtables']],
    ['Digital', ['Podcasts', 'Webinars', 'Social presence']],
    ['Partner-led', ['Joint content', 'Alliance events', 'Co-branded assets']],
  ];
  ch.forEach(([t, items], i) => {
    const x = 0.7 + i * 2.42;
    s.addShape(pres.ShapeType.roundRect, { x, y: 1.95, w: 2.22, h: 2.35, rectRadius: 0.12,
      fill: { color: i % 2 ? MIST : 'E4EBF7' } });
    s.addText(t, { x: x + 0.22, y: 2.08, w: 1.9, h: 0.4, fontSize: 14, bold: true, color: INK,
      fontFace: SANS, isTextBox: true, margin: 0 });
    s.addText(items.map((q, j) => ({ text: q, options: { bullet: true, breakLine: j < items.length - 1 } })),
      { x: x + 0.22, y: 2.55, w: 1.85, h: 1.6, fontSize: 11, color: BODY, fontFace: SANS,
        isTextBox: true, margin: 0, paraSpaceAfter: 5 });
  });
  s.addShape(pres.ShapeType.roundRect, { x: 0.7, y: 4.5, w: 11.9, h: 1.3, rectRadius: 0.12, fill: { color: MIST } });
  s.addShape(pres.ShapeType.rect, { x: 0.7, y: 4.5, w: 2.55, h: 1.3, fill: { color: SLATE } });
  s.addText('Internal\neminence', { x: 0.92, y: 4.5, w: 2.2, h: 1.3, valign: 'middle', fontSize: 15, bold: true,
    color: WHITE, fontFace: SANS, isTextBox: true, margin: 0, lineSpacing: 17 });
  s.addText('The FS partners and directors in other portfolios are a market in their own right \u2014 they are in front of our clients every week. Catalogue briefings, enablement on the proposition one-pagers, deal clinics and internally circulated wins belong in the plan alongside the external channels.',
    { x: 3.5, y: 4.5, w: 8.9, h: 1.3, valign: 'middle', fontSize: 12.5, color: BODY, fontFace: SANS,
      isTextBox: true, margin: 0 });
  s.addShape(pres.ShapeType.roundRect, { x: 0.7, y: 6.0, w: 11.9, h: 0.9, rectRadius: 0.1, fill: { color: NAVY } });
  s.addText('For each activity: internal or external, which channel, how often, who fronts it, which accounts or audiences it targets, and how we measure the return.',
    { x: 1.0, y: 6.0, w: 11.3, h: 0.9, valign: 'middle', fontSize: 13, color: ICE,
      fontFace: SANS, isTextBox: true, margin: 0 });
  s.addText('Channel mix and cadence ' + TBD + ' in the cycle.', { x: 0.7, y: 7.0, w: 11.9, h: 0.35,
    fontSize: 11, italic: true, color: MUTE, fontFace: SANS, isTextBox: true, margin: 0 });
  s.addNotes(`Raised as a distinct part of the plan: how we enhance our positioning in the market, specifically in FSI, around what we do.

Channels mentioned: workshops, seminars, podcasts, articles, and participation in key industry events — Money20/20 was named as an example of the class of event we should consider. The plan should decide which of these are relevant to us rather than assume all of them.

For each channel the plan answers: do we commit, how often, who fronts it, which accounts or markets it targets, and how we measure return — inbound enquiries, meetings created, inclusion on shortlists and RFP invitations.

Worth linking eminence to the account list: eminence activity should point at the accounts and markets the strategy has prioritised, not at general brand awareness.

Internal eminence matters as much as external here. The FS partners and directors sitting in other portfolios are in front of our target clients constantly, and they can only pull us in if they know what we sell. That means scheduled briefings on the catalogue, enablement on the proposition one-pagers, deal clinics where they bring a live situation and we shape the play, and our wins circulated internally rather than only externally. It is the mechanism behind the integrated-market-team principle — the principle sets the expectation, the internal eminence plan is how it is actually resourced and scheduled.`);
}

/* 9 — Alliances */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, 'Alliances', 'Pillar 4 \u00b7 two categories, used for different things.');
  s.addShape(pres.ShapeType.roundRect, { x: 0.7, y: 1.95, w: 11.9, h: 2.05, rectRadius: 0.12, fill: { color: MIST } });
  s.addShape(pres.ShapeType.rect, { x: 0.7, y: 1.95, w: 2.75, h: 2.05, fill: { color: NAVY } });
  s.addText('Broad technology\nalliances', { x: 0.92, y: 2.05, w: 2.4, h: 0.85, fontSize: 14.5, bold: true,
    color: WHITE, fontFace: SANS, isTextBox: true, margin: 0, lineSpacing: 17 });
  s.addText('Relevant to every proposition', { x: 0.92, y: 3.0, w: 2.4, h: 0.6, fontSize: 11,
    color: ICE, fontFace: SANS, isTextBox: true, margin: 0 });
  ['Google', 'AWS', 'Microsoft', 'Oracle'].forEach((n, i) => {
    const x = 3.75 + i * 2.15;
    s.addShape(pres.ShapeType.roundRect, { x, y: 2.2, w: 1.95, h: 0.75, rectRadius: 0.1, fill: { color: SLATE } });
    s.addText(n, { x, y: 2.2, w: 1.95, h: 0.75, align: 'center', valign: 'middle', fontSize: 15, bold: true,
      color: WHITE, fontFace: SANS, isTextBox: true, margin: 0 });
  });
  s.addText('Used for: co-sell into their accounts \u00b7 funding and credits to activate use cases on their platform \u00b7 platform credibility and certifications',
    { x: 3.75, y: 3.1, w: 8.5, h: 0.75, fontSize: 12, color: BODY, fontFace: SANS, isTextBox: true, margin: 0 });
  s.addShape(pres.ShapeType.roundRect, { x: 0.7, y: 4.15, w: 11.9, h: 2.05, rectRadius: 0.12, fill: { color: 'E4EBF7' } });
  s.addShape(pres.ShapeType.rect, { x: 0.7, y: 4.15, w: 2.75, h: 2.05, fill: { color: SLATE } });
  s.addText('Specialised\nproduct vendors', { x: 0.92, y: 4.25, w: 2.4, h: 0.85, fontSize: 14.5, bold: true,
    color: WHITE, fontFace: SANS, isTextBox: true, margin: 0, lineSpacing: 17 });
  s.addText('Relevant to specific offerings only', { x: 0.92, y: 5.2, w: 2.4, h: 0.6, fontSize: 11,
    color: ICE, fontFace: SANS, isTextBox: true, margin: 0 });
  const spec = [
    ['Core banking', 'Temenos \u00b7 Intellect'],
    ['Data platforms', 'Informatica \u00b7 others'],
    ['Origination', 'to be defined'],
    ['Insurance and fund core', 'to be defined'],
  ];
  spec.forEach(([t, v], i) => {
    const x = 3.75 + (i % 2) * 4.3, y = 4.35 + Math.floor(i / 2) * 0.85;
    s.addShape(pres.ShapeType.roundRect, { x, y, w: 4.1, h: 0.72, rectRadius: 0.1, fill: { color: WHITE } });
    s.addText(t, { x: x + 0.2, y, w: 2.0, h: 0.72, valign: 'middle', fontSize: 12, bold: true, color: NAVY,
      fontFace: SANS, isTextBox: true, margin: 0 });
    s.addText(v, { x: x + 2.2, y, w: 1.8, h: 0.72, valign: 'middle', fontSize: 11.5, color: BODY,
      fontFace: SANS, isTextBox: true, margin: 0 });
  });
  s.addShape(pres.ShapeType.roundRect, { x: 0.7, y: 6.35, w: 11.9, h: 0.8, rectRadius: 0.1, fill: { color: NAVY } });
  s.addText('Per alliance the plan answers: what we leverage it for, the relationship type, which offerings and segments it sits behind, which accounts it opens, and what it costs us back.',
    { x: 1.0, y: 6.35, w: 11.3, h: 0.8, valign: 'middle', fontSize: 12.5, color: ICE, fontFace: SANS,
      isTextBox: true, margin: 0 });
  s.addNotes(`Alliances split into two categories that behave differently, and the plan should treat them differently.

Broad technology alliances — the hyperscalers and platform players: Google, AWS, Microsoft and Oracle. They are relevant to every proposition in the catalogue, so they are a portfolio-level relationship rather than an offering-level one. Three things we take from them: co-sell into their own account relationships, funding and credits to activate use cases on their platforms (a real commercial lever, not a nicety — it can fund the first proof of value), and the certifications and platform credibility clients increasingly ask for.

Specialised product vendors — relevant only where the offering matches. Core banking and origination platforms (Temenos and Intellect today), specialised data vendors such as Informatica, and, as the segment priorities imply, the insurance core and fund platform vendors where we currently hold nothing. These are offering-level and segment-level relationships, so the map must show which propositions each sits behind rather than treating them as general alliances.

The gap to close in the cycle: our specialised alliances are banking-weighted. Funds and insurance are priority two and three with no platform partner named, which is a decision the strategy has to make rather than defer.

Two links to the rest of the strategy: the alliance map must reconcile with the account coverage map (an alliance that opens an account should appear against that account), and with the capacity plan, because certifications and accredited people are a real training and hiring cost. Funding available from the broad alliances should also appear in the commercial view — it changes what a first engagement can be priced at.`);
}

/* 10 — Three-week journey */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, 'The three-week journey', 'Three weeks, three gates, one decision.');
  const weeks = [
    ['Week 1', 'Frame', ['Agree the vision and the ambition', 'Baseline current position and credentials', 'Prioritise markets and segments', 'Build the account longlist', 'Draft the proposition catalogue', 'Baseline the team we have today'], 'Gate: agreed market and segment priorities'],
    ['Week 2', 'Define', ['Lock propositions and permitted variants', 'Benchmark deal sizes and pricing', 'Set delivery construct, ADR and margin', 'Shortlist target accounts', 'Map alliances to offerings', 'Size the capacity each target implies'], 'Gate: agreed catalogue and shortlist'],
    ['Week 3', 'Commit', ['Set revenue targets by account, offering and market', 'Assign an owner to every account', 'Agree the team model and recruitment plan', 'Set the eminence calendar', 'Build the activation plan and sign off'], 'Gate: strategy signed off'],
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

/* 11 — Execution governance */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, 'Governing the execution', 'The discipline · what happens after sign-off, so the strategy does not drift.');
  const layers = [
    ['Weekly', NAVY, ['Pursuit review: live deals, blockers, next actions', 'Owned by the account owners']],
    ['Monthly', SLATE, ['Pipeline and coverage against target by market and offering', 'Talent and alliance actions tracked to date']],
    ['Quarterly', NAVY, ['Strategy review: are the choices still right?', 'Reset targets, accounts and eminence calendar']],
  ];
  layers.forEach(([t, c, items], i) => {
    const x = 0.7 + i * 4.07;
    s.addShape(pres.ShapeType.roundRect, { x, y: 2.0, w: 3.75, h: 2.5, rectRadius: 0.12, fill: { color: MIST } });
    s.addShape(pres.ShapeType.rect, { x, y: 2.0, w: 3.75, h: 0.68, fill: { color: c } });
    s.addText(t, { x: x + 0.3, y: 2.0, w: 3.15, h: 0.68, valign: 'middle', fontSize: 17, bold: true,
      color: WHITE, fontFace: SANS, isTextBox: true, margin: 0 });
    s.addText(items.map((q, j) => ({ text: q, options: { bullet: true, breakLine: j < items.length - 1 } })),
      { x: x + 0.3, y: 2.9, w: 3.15, h: 1.5, fontSize: 12.5, color: BODY, fontFace: SANS,
        isTextBox: true, margin: 0, paraSpaceAfter: 9 });
  });
  s.addText('The scorecard behind the rhythm', { x: 0.7, y: 4.75, w: 11.9, h: 0.35, fontSize: 15,
    bold: true, color: INK, fontFace: SANS, isTextBox: true, margin: 0 });
  const kpi = [
    ['Commercial', 'Pipeline, wins, revenue against target'],
    ['Mix', 'Share of revenue from integrated propositions'],
    ['Coverage', 'Target accounts with an active owner and contact'],
    ['Capacity', 'Recruitment against plan, utilisation'],
    ['Alliances', 'Partner-sourced pipeline and joint pursuits'],
    ['Eminence', 'Activities delivered, meetings they created'],
  ];
  kpi.forEach(([h, b], i) => {
    const x = 0.7 + (i % 3) * 4.07, y = 5.2 + Math.floor(i / 3) * 0.72;
    s.addShape(pres.ShapeType.roundRect, { x, y, w: 3.75, h: 0.6, rectRadius: 0.08, fill: { color: 'F2F6FC' } });
    s.addText(h, { x: x + 0.18, y, w: 1.35, h: 0.6, valign: 'middle', fontSize: 11.5, bold: true, color: NAVY,
      fontFace: SANS, isTextBox: true, margin: 0 });
    s.addText(b, { x: x + 1.55, y, w: 2.05, h: 0.6, valign: 'middle', fontSize: 10.5, color: BODY,
      fontFace: SANS, isTextBox: true, margin: 0 });
  });
  s.addText('Decision rights, chair and escalation route to be set at sign-off.',
    { x: 0.7, y: 6.75, w: 11.9, h: 0.35, fontSize: 11.5, italic: true, color: MUTE, fontFace: SANS,
      isTextBox: true, margin: 0 });
  s.addNotes(`Added because a strategy that is signed off and then unmanaged reverts to whatever the pipeline happens to offer. This is the management-system layer: the rhythm, the measures and the decision rights that keep the choices alive.

Three loops, deliberately different in purpose: weekly is deal-level and run by the account owners; monthly is coverage and progress against target, including talent and alliance commitments; quarterly re-tests the choices themselves — markets, accounts, propositions — rather than just the numbers.

The scorecard is designed so that the strategy's own principles are measurable. The mix KPI is the important one: if the share of revenue from integrated propositions does not move, we are back to selling in silos whatever the catalogue says.

Still to set at sign-off: who chairs each forum, what decisions each can take without escalation, and where the strategy log lives so it stays current rather than being rebuilt each quarter.`);
}

/* 12 — Cycle governance */
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addShape(pres.ShapeType.ellipse, { x: -2.2, y: 4.4, w: 6.2, h: 6.2, fill: { color: SLATE, transparency: 55 } });
  s.addText('How we run the cycle', { x: 0.8, y: 0.6, w: 11, h: 0.8, fontSize: 36, bold: true,
    color: WHITE, fontFace: HEAD, isTextBox: true, margin: 0 });
  s.addText('Cadence, inputs and the decisions we are asking for.', { x: 0.8, y: 1.42, w: 11, h: 0.45,
    fontSize: 14.5, color: ICE, fontFace: SANS, isTextBox: true, margin: 0 });
  const cols = [
    ['Cadence', ['Kick-off to confirm scope and principles', 'Weekly working session per workstream', 'Checkpoint at the end of each week', 'Sign-off session in week three']],
    ['Inputs we need', ['Current pipeline and revenue by pillar', 'Existing account relationships', 'Alliance and partnership status', 'Delivery capacity and skills view', 'Historical deal values']],
    ['Decisions requested', ['The vision, then the ambition in numbers', 'Priority markets and segments', 'Target account list', 'Revenue and margin targets by account', 'Account ownership', 'Alliance priorities', 'Investment and hiring asks']],
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
  s.addText('Dates and participants beyond the core team are set at kick-off.', { x: 0.8, y: 6.4, w: 7, h: 0.4,
    fontSize: 12, italic: true, color: '9AA7C4', fontFace: SANS, isTextBox: true, margin: 0 });
  s.addNotes(`Close on two asks: confirm the principles, and nominate the workstream owners.

Inputs are the long pole — pipeline, revenue by pillar and market, existing relationships, alliance status, delivery capacity and historical deal values all sit with different people. Chase them in the first 48 hours or week 1 slips.

Decisions requested at sign-off: priority markets and segments, the target account list, revenue and margin targets, account ownership, and the investment and hiring asks coming out of the talent and eminence workstreams.

Open items to resolve at kick-off: audience for this pack, the definition of partner for account ownership, and the start date.`);
}

/* Appendix divider */
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addShape(pres.ShapeType.ellipse, { x: 9.6, y: 3.6, w: 5.2, h: 5.2, fill: { color: SLATE, transparency: 50 } });
  s.addText('Appendix', { x: 0.9, y: 2.6, w: 8, h: 0.9, fontSize: 44, bold: true, color: WHITE,
    fontFace: HEAD, isTextBox: true, margin: 0 });
  s.addText('Blank templates for every artefact the cycle produces', { x: 0.95, y: 3.6, w: 8.5, h: 0.5,
    fontSize: 17, color: ICE, fontFace: SANS, isTextBox: true, margin: 0 });
  s.addText('These are the working formats. They are filled during the three weeks and become the strategy record.',
    { x: 0.95, y: 4.25, w: 8.5, h: 0.6, fontSize: 13.5, italic: true, color: '9AA7C4',
      fontFace: SANS, isTextBox: true, margin: 0 });
  s.addNotes(`The appendix is the strategy log: the templates we fill in as the cycle runs. Each one maps to a workstream artefact, so at sign-off the completed appendix is the strategy itself, not a separate write-up.

Keep them as living tables — one file, versioned, updated at each weekly checkpoint rather than rebuilt at the end.`);
}

/* Appendix — the strategy on a page */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, 'Template — the strategy on a page', 'The completed strategy on one sheet. Filled at sign-off, from the artefacts that follow.');
  s.addShape(pres.ShapeType.roundRect, { x: 0.7, y: 2.0, w: 11.9, h: 0.9, rectRadius: 0.1,
    fill: { color: WHITE }, line: { color: NAVY, width: 1.25 } });
  s.addText('AMBITION', { x: 1.0, y: 2.0, w: 2.0, h: 0.9, valign: 'middle', fontSize: 12, bold: true,
    color: NAVY, charSpacing: 1.5, fontFace: SANS, isTextBox: true, margin: 0 });
  s.addShape(pres.ShapeType.line, { x: 3.2, y: 2.62, w: 9.1, h: 0, line: { color: 'C9D4E6', width: 1 } });
  const cols = [
    ['Market and clients', ['Markets and segments', 'Top accounts and offerings', 'Revenue: gross / net']],
    ['Offerings', ['Propositions', 'Use cases', 'Flavours']],
    ['Delivery and economics', ['Construct and ADR', 'Margin', 'Team and gap']],
    ['Route to market', ['Alliances', 'Eminence', 'Differentiation']],
  ];
  cols.forEach(([t, labels], i) => {
    const x = 0.7 + i * 3.02, w = 2.84;
    s.addShape(pres.ShapeType.roundRect, { x, y: 3.1, w, h: 2.75, rectRadius: 0.1,
      fill: { color: WHITE }, line: { color: 'C9D4E6', width: 1 } });
    s.addShape(pres.ShapeType.rect, { x, y: 3.1, w, h: 0.55, fill: { color: i % 2 ? SLATE : NAVY } });
    s.addText(t, { x: x + 0.16, y: 3.1, w: w - 0.32, h: 0.55, valign: 'middle', fontSize: 11.5, bold: true,
      color: WHITE, fontFace: SANS, isTextBox: true, margin: 0 });
    labels.forEach((l, j) => {
      const ly = 3.9 + j * 0.62;
      s.addText(l, { x: x + 0.22, y: ly - 0.02, w: 1.55, h: 0.3, fontSize: 10.5, color: MUTE,
        fontFace: SANS, isTextBox: true, margin: 0 });
      s.addShape(pres.ShapeType.line, { x: x + 0.22, y: ly + 0.33, w: w - 0.44, h: 0,
        line: { color: 'E1E8F2', width: 1 } });
    });
  });
  s.addShape(pres.ShapeType.roundRect, { x: 0.7, y: 6.05, w: 11.9, h: 0.85, rectRadius: 0.1,
    fill: { color: WHITE }, line: { color: NAVY, width: 1.25 } });
  s.addText('DISCIPLINE', { x: 1.0, y: 6.05, w: 2.0, h: 0.85, valign: 'middle', fontSize: 12, bold: true,
    color: NAVY, charSpacing: 1.5, fontFace: SANS, isTextBox: true, margin: 0 });
  s.addShape(pres.ShapeType.line, { x: 3.2, y: 6.63, w: 9.1, h: 0, line: { color: 'C9D4E6', width: 1 } });
  s.addText('One sheet, one version, dated. Everything behind it lives in the artefacts that follow.',
    { x: 0.7, y: 6.98, w: 11.9, h: 0.35, fontSize: 11, italic: true, color: MUTE, fontFace: SANS,
      isTextBox: true, margin: 0 });
  s.addNotes(`This is the actual strategy, in the form it will be presented and carried: one sheet.

It mirrors the architecture slide earlier in the pack, but where that one names the questions, this one holds the answers: the ambition in numbers, the three pillars filled in, and the operating discipline underneath.

Rule of thumb: if something cannot be said on this sheet, it belongs in one of the artefacts behind it, not in the strategy. The sheet is what leadership signs and what the quarterly review is held against; the artefacts are the working record.

Keep it dated and versioned. When the quarterly review changes a choice, this sheet is reissued rather than annotated.`);
}

/* Appendix — proposition one-pager */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, 'Template — proposition one-pager', 'One sheet per proposition in the catalogue. Every dimension on a single page.');
  s.addShape(pres.ShapeType.roundRect, { x: 0.7, y: 1.95, w: 11.9, h: 0.72, rectRadius: 0.1,
    fill: { color: NAVY } });
  s.addText('PROPOSITION', { x: 1.0, y: 1.95, w: 2.1, h: 0.72, valign: 'middle', fontSize: 11.5, bold: true,
    color: ICE, charSpacing: 1.5, fontFace: SANS, isTextBox: true, margin: 0 });
  s.addShape(pres.ShapeType.line, { x: 3.2, y: 2.42, w: 5.4, h: 0, line: { color: '5C6E96', width: 1 } });
  s.addText('LEAD PILLAR / PRACTICE', { x: 8.9, y: 1.95, w: 2.4, h: 0.72, valign: 'middle', fontSize: 9.5,
    bold: true, color: '9AA7C4', charSpacing: 1, fontFace: SANS, isTextBox: true, margin: 0 });
  s.addShape(pres.ShapeType.line, { x: 11.3, y: 2.42, w: 1.1, h: 0, line: { color: '5C6E96', width: 1 } });
  const boxes = [
    'What it is, and the use cases we lead with',
    'Segments it serves, and the buyer in each',
    'Permitted flavours and variants',
    'Target markets and accounts',
    'Deal size and price band',
    'Delivery construct, ADR and target margin',
    'AI and innovation embedded in it',
    'Alliance behind it, and what they bring',
    'Credentials, eminence hooks and key roles',
  ];
  boxes.forEach((label, i) => {
    const x = 0.7 + (i % 3) * 4.03, y = 2.85 + Math.floor(i / 3) * 1.4;
    s.addShape(pres.ShapeType.roundRect, { x, y, w: 3.85, h: 1.25, rectRadius: 0.1,
      fill: { color: WHITE }, line: { color: 'C9D4E6', width: 1 } });
    s.addText(label, { x: x + 0.2, y: y + 0.08, w: 3.45, h: 0.45, fontSize: 10.5, bold: true, color: NAVY,
      fontFace: SANS, isTextBox: true, margin: 0 });
    [0, 1].forEach(k => {
      s.addShape(pres.ShapeType.line, { x: x + 0.2, y: y + 0.72 + k * 0.32, w: 3.45, h: 0,
        line: { color: 'E1E8F2', width: 1 } });
    });
  });
  s.addText('One page per proposition. If a box cannot be filled, the proposition is not ready for the catalogue.',
    { x: 0.7, y: 7.0, w: 11.9, h: 0.35, fontSize: 11, italic: true, color: MUTE, fontFace: SANS,
      isTextBox: true, margin: 0 });
  s.addNotes(`The summary sheet for each proposition, pulling every dimension we have discussed onto one page: what it is and the use cases, the buyer, the permitted flavours, where and to whom we sell it, the price band, the delivery construct with ADR and margin, the AI and innovation content, the alliance behind it, and the credentials and roles needed to sell and deliver it.

It doubles as the readiness test. The principles bite here: if the AI and innovation box is empty, the proposition breaches the embedded-AI principle. If the delivery construct and margin box is empty, it is priced without knowing whether it makes money. If the credentials box is empty, we cannot yet sell it in this domain. An unfillable box is a gap, not a formatting problem.

Keep one sheet per proposition rather than a long catalogue document — it is what a partner takes into an account conversation.`);
}

function tmplSlide(title, sub, cols, colW, rowCount, footer, notes) {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, title, sub);
  const rows = [cols.map(c => ({ text: c, options: { bold: true, color: WHITE } }))];
  for (let r = 0; r < rowCount; r++) rows.push(cols.map(() => ' '));
  const rowH = [0.46];
  for (let r = 0; r < rowCount; r++) rowH.push(0.6);
  s.addTable(rows, {
    x: 0.7, y: 2.15, w: 11.9, colW, rowH, fontSize: 12, fontFace: SANS, color: BODY,
    border: { type: 'solid', color: 'D8E0EE', pt: 1 }, align: 'left', valign: 'middle',
    margin: 7, fill: { color: WHITE },
  });
  s.addShape(pres.ShapeType.rect, { x: 0.7, y: 2.15, w: 11.9, h: 0.46, fill: { color: NAVY } });
  let cx = 0.7;
  cols.forEach((c, i) => {
    s.addText([{ text: c, options: { bold: true } }], { x: cx + 0.12, y: 2.15, w: colW[i] - 0.2, h: 0.46,
      valign: 'middle', fontSize: 11.5, color: WHITE, fontFace: SANS, isTextBox: true, margin: 0 });
    cx += colW[i];
  });
  s.addText(footer, { x: 0.7, y: 2.3 + 0.46 + rowCount * 0.6, w: 11.9, h: 0.4, fontSize: 12,
    italic: true, color: MUTE, fontFace: SANS, isTextBox: true, margin: 0 });
  s.addNotes(notes);
  return s;
}

tmplSlide('Template — market prioritisation', 'Pillar 1 · Market and clients · one row per market.',
  ['Market', 'Segments in focus', 'Why this market', 'Tier', 'Owner'],
  [2.4, 3.5, 3.6, 1.2, 1.2], 6,
  'Tier one or tier two, per segment where they differ. Rationale should stand on evidence, not familiarity.',
  `Fill one row per market under consideration, including the ones we decide against — the rejected rows are as useful as the accepted ones when this is revisited.

Context: Saudi Arabia, UAE and Qatar are the expected primary markets; Kuwait, Oman and Jordan the secondary set where our key relationships sit. Both to be confirmed against evidence in week 1.`);

tmplSlide('Template — target accounts', 'Pillar 1 · Market and clients · one row per named account.',
  ['Account', 'Segment', 'Market', 'Tier', 'Offerings we target there', 'Alliance', 'Owner', 'Stage'],
  [2.1, 1.3, 1.2, 0.8, 2.8, 1.3, 1.3, 1.1], 6,
  'Segment: banking, funds or insurance. Tier one accounts are named individually; tier two may be grouped.',
  `This is the core artefact of the whole exercise — the named account list with account mapping.

Rules: one owner per account, never two. The propositions column names what we lead with, not everything we could sell. The alliance column links to the alliance map, so we can see which accounts a partner opens.`);

tmplSlide('Template — account revenue plan', 'Pillar 1 · Market and clients · one row per account and offering.',
  ['Account', 'Offering targeted', 'Timeframe', 'Gross revenue', 'Net revenue', 'Confidence'],
  [2.5, 3.0, 1.5, 1.7, 1.7, 1.5], 6,
  'Gross and net on every line. The account total and the market total must reconcile to the portfolio target.',
  `This is what turns a target account list into a plan: for each account, which offerings we are going there with, and how much revenue each is expected to carry.

Both numbers are recorded. Gross is the full contract value we bill, including pass-through, third-party software and partner or subcontract content. Net is the revenue delivered by our own people — what the capacity plan and the margin calculation actually run on. Fix the exact definition at kick-off; different parts of the firm use these words differently, and mixing them makes the roll-up meaningless.

Timeframe matters because a core platform programme books over years while an AI activation books in a quarter; without it, an account total is unreadable.

Confidence keeps the plan honest: relationship-only ambition and a qualified pursuit should not be added together as if they were the same thing.

Three reconciliations to hold: account lines roll up to the market total, market totals roll up to the portfolio ambition, and net revenue reconciles to the capacity we have or plan to hire.`);

tmplSlide('Template — coverage and ownership', 'Pillar 1 · Market and clients · one row per priority account.',
  ['Account', 'Partner owner', 'Alliance lead', 'Delivery lead', 'Contact cadence'],
  [2.9, 2.4, 2.3, 2.3, 2.0], 6,
  'Partner owner means our accountable partner. Alliance lead is the vendor relationship behind the account.',
  `Two different mappings live side by side here: our own consulting partners, who own the client relationship, and the alliance or vendor partner we go to market with on that account.

Cadence keeps this honest — an owner with no contact rhythm is a name on a page.`);

tmplSlide('Template — proposition catalogue', 'Pillar 2 · Offerings · one row per integrated proposition.',
  ['Proposition', 'Segments it serves', 'Practices involved', 'Buyer per segment', 'Permitted variants'],
  [2.6, 2.1, 2.2, 2.4, 2.6], 6,
  'Integrated propositions first. A single-pillar offer is recorded as a variant, not as its own row.',
  `The principle bites hardest here: the catalogue is integrated propositions across Industry Solutions, Data and AI, and Engineering and Cloud. Focused offers such as enterprise AI are recorded as permitted variants of an integrated proposition and sold as entry points.

Buyer means the actual role that signs, and it differs by segment: COO, CIO or head of retail banking in a bank; COO or head of investment operations in a fund; chief underwriting or claims officer in an insurer.`);

tmplSlide('Template — sizing and pricing', 'Pillar 3 · Delivery and economics · one row per offering.',
  ['Offering', 'What drives scope', 'Typical deal size', 'Market price band', 'Notes'],
  [2.6, 3.2, 2.0, 2.1, 2.0], 6,
  'Bands, not point prices. Name the drivers that move the number.',
  `Pricing will not reduce to one number per offering. For core modernisation and migration the drivers include segment and scope (retail, corporate or SME in banking; life or general in insurance; fund accounting or front-office in funds), single-country versus multi-country, and whether full data migration is in scope. Record the drivers alongside the band so the number can be reconstructed.

Same discipline for data platform work, AI activation, and cloud and integration engagements.`);

tmplSlide('Template — offering economics', 'Pillar 3 · Delivery and economics · one row per offering.',
  ['Offering', 'Price basis', 'Onshore %', 'Nearshore %', 'Offshore %', 'Blended ADR', 'Target GM'],
  [2.6, 2.3, 1.4, 1.5, 1.4, 1.4, 1.3], 6,
  'The construct drives the rate, the rate drives the margin. Record the floor as well as the target.',
  `The bridge between what the market pays and what we keep. One row per offering, because the construct differs sharply between a core platform programme, a data platform build and an AI activation \u2014 and again between segments.

Price basis: fixed price, time and materials, outcome-based or a managed service — it changes how margin is earned and where the risk sits.

Keep the onshore split honest. Regulatory and data-residency requirements in Saudi and the UAE, and client contractual terms, often set the floor for onshore presence before commercial preference does.

Record both the target gross margin and the floor, and note the approval route for anything below it. A margin floor that nobody enforces is not a floor.`);

tmplSlide('Template — revenue targets', 'Pillar 3 · Delivery and economics · one row per market.',
  ['Market', 'Segment', 'Industry Solutions', 'Data & AI', 'Engineering & Cloud', 'Integrated', 'Total'],
  [1.9, 1.6, 2.0, 1.7, 2.1, 1.4, 1.2], 6,
  'One row per market and segment. Reconciles to the account revenue plan and to capacity.',
  `The integrated column is deliberate: if most revenue sits in single-pillar columns, the strategy has not been followed.

Reconcile every target against the talent template before sign-off. A number we cannot staff is not a target.`);

tmplSlide('Template — team and talent', 'Pillar 3 · Delivery and economics · one row per role and level.',
  ['Role and level', 'Dedicated today', 'Shared available', 'Gap', 'Build, borrow or buy', 'By when'],
  [2.6, 2.0, 2.0, 1.4, 2.1, 1.8], 6,
  'Dedicated means FSI-exclusive. Shared means time we can genuinely count on, not headcount that exists.',
  `Three numbers per row: who is ours exclusively, what we can draw on from the wider portfolio, and what is missing once both are counted.

The gap is expressed by level — partners, directors, managers, specialists — because that is how recruitment is approved. Each gap then gets a route (build from within, borrow from another portfolio, or hire) and a date tied to the targets it supports.`);

tmplSlide('Template — alliance map', 'Pillar 4 · Route to market · one row per alliance.',
  ['Alliance', 'Type', 'Role', 'Offerings and segments it sits behind', 'Accounts it opens', 'What we owe them'],
  [1.7, 1.5, 1.6, 3.1, 2.1, 1.9], 6,
  'Type: broad technology alliance, or specialised product vendor. Role: co-sell, resell, delivery or referral.',
  `Start from the alliances we already hold: Oracle, Google and Amazon as primary; Temenos and Intellect as specialised platform alliances, both weighted to banking. The cycle should decide what we do about funds and insurance platforms, where we currently hold no specialised alliance.

The last column matters as much as the rest — certifications, accredited staff and pipeline commitments are a cost that lands in the talent plan and the budget.`);

tmplSlide('Template — eminence calendar', 'Pillar 4 · Route to market · one row per planned activity.',
  ['Activity', 'Audience', 'Channel', 'Target accounts or audience', 'Owner', 'Date', 'Measure'],
  [2.3, 1.4, 1.7, 2.6, 1.4, 1.2, 1.3], 6,
  'Audience: external client market, or internal \u2014 FS partners and directors in other portfolios.',
  `Channels in scope: published views and articles, conference stages and flagship FSI events, executive workshops, seminars and roundtables, podcasts and webinars, and joint activity with alliance partners.

Measurement should be concrete: inbound enquiries, meetings created, shortlist and RFP invitations — not impressions.`);

tmplSlide('Template — execution scorecard', 'The discipline · one row per measure.',
  ['Measure', 'What it tells us', 'Target', 'Frequency', 'Owner', 'Forum'],
  [2.4, 3.2, 1.6, 1.5, 1.6, 1.6], 6,
  'Every measure has an owner and a forum. A measure nobody presents is a measure nobody acts on.',
  `The scorecard behind the operating rhythm. Keep it short — six to eight measures — so the review stays a decision meeting rather than a reporting one.

The measure to fight for is the mix: share of revenue coming from integrated propositions. It is the only number that tells us whether the integrated-not-siloed principle survived contact with the market.`);

pres.writeFile({ fileName: '/home/user/bae/FSI strategy/EAIND-FSI-Strategy-Development-Plan.pptx' })
  .then(f => console.log('wrote', f));
