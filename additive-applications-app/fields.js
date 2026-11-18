/* Field definitions, AI prompt, positional checks and Teams card builder.
   Works in the browser (window.AAFields) and in Node (module.exports) for testing. */
(function (root) {
  'use strict';

  const REQUIRED = [
    { key: 'appName', label: 'Application Name' },
    { key: 'appDesc', label: 'Application Description' },
    { key: 'department', label: 'Department/area/product line' },
    { key: 'issue', label: 'Current Issue/challenge' },
    { key: 'downtime', label: 'Downtime impact' },
    { key: 'urgency', label: 'Urgency/Importance Scale (1-10)', type: 'scale' },
    { key: 'partSize', label: 'Approximate part size' },
    { key: 'currentMethod', label: 'Current material/manufacturing method' },
    { key: 'confidence', label: 'Confidence Scale (1-10)', type: 'scale' },
    { key: 'photo', label: 'Photo', type: 'photo' },
    { key: 'suggestedAM', label: 'Suggested AM material/process' }
  ].map(f => Object.assign(f, { section: 'required' }));

  const SUPPORTING = [
    { key: 'eau', label: 'EAU (Estimated Annual Usage)' },
    { key: 'partCost', label: 'Current part cost' },
    { key: 'leadTime', label: 'Lead time' },
    { key: 'eol', label: 'EOL concern' },
    { key: 'safety', label: 'Safety concern' },
    { key: 'customerInterest', label: 'Customer buy-in/interest' },
    { key: 'customerQuote', label: 'Customer quote/input' }
  ].map(f => Object.assign(f, { section: 'supporting' }));

  const ALL = REQUIRED.concat(SUPPORTING);
  const KEYS = ALL.map(f => f.key);
  const byKey = Object.fromEntries(ALL.map(f => [f.key, f]));

  /* ---------- small helpers ---------- */
  const PLACEHOLDER = /^(n\/?a|—|–|-+|✓|✔|☑|\?)$/i;
  function clean(v) {
    v = (v == null ? '' : String(v)).trim();
    return PLACEHOLDER.test(v) ? '' : v;
  }
  function scaleValue(v) {
    const m = String(v == null ? '' : v).match(/\b(10|[1-9])\b/);
    if (!m) return null;
    const n = parseInt(m[1], 10);
    return n >= 1 && n <= 10 ? n : null;
  }
  // 1 -> "!"   10 -> "!?!?!?!?"
  function confidenceMarks(n) {
    n = scaleValue(n);
    if (!n) return '';
    const full = '!?!?!?!?';
    return full.slice(0, Math.round(1 + (n - 1) * 7 / 9));
  }
  function urgencyHue(n) { return 120 * (10 - n) / 9; } // 1 green -> 10 red
  function hslToHex(h, s, l) {
    s /= 100; l /= 100;
    const k = n => (n + h / 30) % 12;
    const a = s * Math.min(l, 1 - l);
    const f = n => l - a * Math.max(-1, Math.min(k(n) - 3, Math.min(9 - k(n), 1)));
    return [f(0), f(8), f(4)].map(x => Math.round(x * 255).toString(16).padStart(2, '0')).join('').toUpperCase();
  }
  function urgencyHex(n, lightness) {
    n = scaleValue(n) || 5;
    return hslToHex(urgencyHue(n), 70, lightness == null ? 45 : lightness);
  }

  /* ---------- AI prompt ---------- */
  function buildPrompt() {
    const list = (arr) => arr.map((f, i) => `${i + 1}. ${f.key} — ${f.label}`).join('\n');
    const shape = '{"rotateToFix":[0],"listStyle":"numbered|bulleted|mixed|unclear","fields":{' +
      KEYS.map(k => `"${k}":{"value":"","uncertain":false}`).join(',') +
      '},"warnings":[],"extraText":""}';
    return `You are reading photos of a handwritten notebook page. The writer records one additive-manufacturing (3D printing) application idea per entry using a fixed list of fields, always in the same order. They usually do NOT write the field names; they write only the values, one per bullet point or numbered line.

The list has two sections.

REQUIRED (in this order):
${list(REQUIRED)}

SUPPORTING (in this order):
${list(SUPPORTING)}

Page orientation:
- A photo may be sideways or upside down, or the notebook may have been written in landscape. Read the text in its correct reading orientation no matter how the photo is turned.
- In "rotateToFix", give one number per photo, in order: how many degrees the photo must be turned clockwise so the handwriting reads normally (left to right, top to bottom). Use only 0, 90, 180 or 270. Use 0 when the text already reads normally.

How to read the page:
- Any mark at the start of a line starts the next field. Marks include bullets and dots (• · ●), dashes (- – —), asterisks (*), plus signs (+), check marks and crosses (✓ ✔ ☑ ✗ x), boxes (☐ □ ■), arrows (→ > ➤ -> =>), stars (★ ☆), circles (o ○), tildes (~), letters such as "a)" or "b.", and numbers (1.  1)  (1)  ①). The writer may use any of these and may mix different marks on the same page.
- The mark itself is not part of the value: return only the text after it.
- A line holding only a mark, or a mark followed by "N/A", "—" or "-", is a field left blank on purpose. For photo, a line with only a mark (especially a check mark) means a photo was taken: return "✓".
- Lines without a mark continue the previous field. If the page uses no marks at all, treat each new line as the next field unless it is clearly the wrapped continuation of the line above (it starts in lowercase or finishes the previous sentence).
- If the items are numbered, use the number to place the value (number 7 under REQUIRED is partSize), even if some numbers are skipped. If the items use other marks, use their order.
- A heading such as "Required", "Req", "R", "Supporting", "Support", "S", or a drawn divider line starts a section and restarts the count. If there is no supporting heading, items after the 11th go to SUPPORTING in order.
- If the writer did write a field label (for example "EAU: 500"), use the label to place the value and leave the label out of the value.
- When the text after the mark is just "N/A", "—", "-" or "x", the field was left blank on purpose: return "".
- Transcribe the handwriting as written. Invent nothing. Keep units, $ signs and abbreviations. Fix only obvious misspellings you are sure about.
- urgency and confidence should be numbers from 1 to 10. Return just the number when you can read one.
- If a value does not fit its field (for example words where a 1-10 number is expected), a bullet was probably skipped or added. Still place values by position, mark the affected fields uncertain, and explain in warnings.
- Mark a field uncertain when the handwriting is hard to read or its placement is doubtful.
- Put writing that belongs to no field (dates, side notes, page titles) in extraText.
- Several photos are consecutive pages of the same entry. If more than one entry appears, read only the first and say so in warnings.

Reply with JSON only, no other text, in exactly this shape:
${shape}`;
  }

  /* ---------- part photo orientation ---------- */
  function buildPartOrientationPrompt() {
    return `This is a photo of a manufactured part, possibly lying next to a ruler or tape measure. Decide how many degrees the photo must be turned clockwise so it appears the right way up.
Judge from, in order: readable text or numbers (ruler markings, labels, stamped or printed text on the part), the direction of the ruler's numbers, the table or floor surface and background, and lighting or shadows.
If there is no clear cue (for example a plain part on a plain background), answer 0.
Reply with JSON only: {"rotateToFix":0,"confident":true} where rotateToFix is 0, 90, 180 or 270.`;
  }
  function parseOrientation(text) {
    const s = String(text || ''), a = s.indexOf('{'), b = s.lastIndexOf('}');
    if (a < 0 || b <= a) return 0;
    try {
      const o = JSON.parse(s.slice(a, b + 1));
      if (o.confident === false) return 0;
      const n = ((Math.round(Number(o.rotateToFix) / 90) * 90) % 360 + 360) % 360;
      return isNaN(n) ? 0 : n;
    } catch (e) { return 0; }
  }

  /* ---------- AI response parsing ---------- */
  function parseAIResponse(text) {
    const s = String(text || '');
    const a = s.indexOf('{'), b = s.lastIndexOf('}');
    if (a < 0 || b <= a) throw new Error('The AI reply did not contain any data.');
    let obj;
    try { obj = JSON.parse(s.slice(a, b + 1)); }
    catch (e) { throw new Error('The AI reply could not be read as data.'); }
    const fields = {}, uncertain = [];
    const src = obj.fields || {};
    for (const k of KEYS) {
      let v = src[k];
      let unc = false;
      if (v && typeof v === 'object') { unc = !!v.uncertain; v = v.value; }
      v = k === 'photo' ? String(v == null ? '' : v).trim() : clean(v);
      fields[k] = v;
      if (unc) uncertain.push(k);
    }
    return {
      listStyle: obj.listStyle || 'unclear',
      fields,
      uncertain,
      warnings: Array.isArray(obj.warnings) ? obj.warnings.map(String).filter(Boolean) : [],
      extraText: clean(obj.extraText),
      rotateToFix: (Array.isArray(obj.rotateToFix) ? obj.rotateToFix : [obj.rotateToFix])
        .map(n => { n = ((Math.round(Number(n) / 90) * 90) % 360 + 360) % 360; return isNaN(n) ? 0 : n; })
    };
  }

  /* ---------- checks that catch a skipped bullet ---------- */
  function validate(fields) {
    const warnings = [], flagged = [];
    for (const k of ['urgency', 'confidence']) {
      const raw = (fields[k] || '').trim();
      if (raw && !scaleValue(raw)) {
        flagged.push(k);
        warnings.push(`${byKey[k].label} reads "${raw}", which isn't a 1–10 number. A bullet above it may have been skipped or added.`);
      } else if (!raw) {
        flagged.push(k);
        warnings.push(`${byKey[k].label} wasn't found. Set it with the slider.`);
      }
    }
    const needsDigit = { partCost: 'a cost', partSize: 'a size', eau: 'a quantity' };
    for (const k in needsDigit) {
      const raw = (fields[k] || '').trim();
      if (raw && !/\d/.test(raw)) {
        flagged.push(k);
        warnings.push(`${byKey[k].label} reads "${raw}", which doesn't look like ${needsDigit[k]}. Check that it's in the right slot.`);
      }
    }
    if (!(fields.appName || '').trim()) { flagged.push('appName'); warnings.push('Application Name is empty.'); }
    return { warnings, flagged };
  }

  /* ---------- fixing a shift ---------- */
  function sectionKeys(key) {
    return (byKey[key].section === 'required' ? REQUIRED : SUPPORTING).map(f => f.key);
  }
  // A bullet was skipped above `key`: move this and following values down one slot.
  function shiftDown(fields, key) {
    const ks = sectionKeys(key), i = ks.indexOf(key), out = Object.assign({}, fields);
    const last = ks.length - 1;
    const overflow = out[ks[last]];
    for (let j = last; j > i; j--) out[ks[j]] = fields[ks[j - 1]] || '';
    out[key] = '';
    if (overflow && last > i) out[ks[last]] = [out[ks[last]], overflow].filter(Boolean).join('; ');
    return out;
  }
  // This slot holds an extra line: remove it and move following values up one slot.
  function shiftUp(fields, key) {
    const ks = sectionKeys(key), i = ks.indexOf(key), out = Object.assign({}, fields);
    const removed = fields[key] || '';
    for (let j = i; j < ks.length - 1; j++) out[ks[j]] = fields[ks[j + 1]] || '';
    out[ks[ks.length - 1]] = '';
    return { fields: out, removed };
  }

  /* ---------- Excel template rows ---------- */
  function labeled(pairs) {
    return pairs.filter(p => clean(p[1])).map(p => p[0] ? `${p[0]}: ${clean(p[1])}` : clean(p[1])).join('\n');
  }
  function derive(entry) {
    const f = entry.fields || {};
    const u = scaleValue(f.urgency), c = scaleValue(f.confidence);
    return {
      originator: entry.originator || '',
      department: clean(f.department),
      workstreamOwner: entry.workstreamOwner || '',
      appName: clean(f.appName),
      appDesc: clean(f.appDesc),
      issue: labeled([['', f.issue], ['Downtime impact', f.downtime], ['EOL concern', f.eol], ['Safety concern', f.safety]]),
      partInfo: labeled([['Size', f.partSize], ['Current material/method', f.currentMethod], ['Current part cost', f.partCost], ['Lead time', f.leadTime]]),
      additiveNotes: clean(f.suggestedAM),
      customerUsage: labeled([['EAU', f.eau]]),
      urgency: u,
      confidence: c,
      confidenceText: c ? `${c}  ${confidenceMarks(c)}` : '',
      customerInput: labeled([['Buy-in/interest', f.customerInterest], ['Quote/input', f.customerQuote]])
    };
  }

  const TEMPLATE_ROWS = [
    ['originator', 'Originator', ''],
    ['department', 'Department/Area/Product Line', ''],
    ['workstreamOwner', 'Workstream Owner', ''],
    ['appName', 'Application Name', ''],
    ['appDesc', 'Application Description', ''],
    ['issue', 'Current Issue/Challenge', ' (Safety, EOL, Downtime, inventory, quality, etc.)'],
    ['partInfo', 'Part Info', ' (Size, Material, MFG Method, etc.)'],
    ['additiveNotes', 'Notes for Additive Solution', ' (Print method, material, onsite vs outsource, etc.)'],
    ['customerUsage', 'Customer Usage', ' (Frequency and/or EAU)'],
    ['urgency', 'Urgency/Importance Scale', ' (1-10)'],
    ['confidenceText', 'Confidence Scale', ' (1-10)'],
    ['customerInput', 'Customer input/interest', ''],
    ['image', 'Image', '']
  ];

  /* ---------- Teams Adaptive Card ---------- */
  function buildTeamsPayload(entry, opts) {
    opts = opts || {};
    const d = derive(entry);
    const name = entry.originator || 'Someone';
    const urgColor = !d.urgency ? 'Default' : d.urgency >= 8 ? 'Attention' : d.urgency >= 5 ? 'Warning' : 'Good';
    const tb = (text, extra) => Object.assign({ type: 'TextBlock', text: String(text), wrap: true }, extra || {});
    const md = s => String(s).split('\n').join('\n\n');
    const body = [
      tb(`**${name}** logged a new additive application`, { size: 'Medium' }),
      tb(d.appName || '(no application name)', { size: 'ExtraLarge', weight: 'Bolder', spacing: 'Small' }),
      {
        type: 'ColumnSet', spacing: 'Medium', columns: [
          { type: 'Column', width: 'stretch', items: [tb('Urgency / importance', { isSubtle: true, size: 'Small' }), tb(d.urgency ? `${d.urgency}/10` : 'Not set', { size: 'Large', weight: 'Bolder', color: urgColor, spacing: 'None' })] },
          { type: 'Column', width: 'stretch', items: [tb('Confidence', { isSubtle: true, size: 'Small' }), tb(d.confidence ? `${d.confidence}/10  ${confidenceMarks(d.confidence)}` : 'Not set', { size: 'Large', weight: 'Bolder', spacing: 'None' })] }
        ]
      }
    ];
    const facts = [['Department/area/product line', d.department], ['Workstream owner', d.workstreamOwner], ['Customer usage', d.customerUsage]]
      .filter(p => p[1]).map(p => ({ title: p[0], value: p[1] }));
    if (facts.length) body.push({ type: 'FactSet', spacing: 'Medium', facts });
    const sections = [
      ['Application description', d.appDesc],
      ['Current issue/challenge', d.issue],
      ['Part info', d.partInfo],
      ['Notes for additive solution', d.additiveNotes],
      ['Customer input/interest', d.customerInput]
    ];
    for (const [title, text] of sections) {
      if (!text) continue;
      body.push(tb(title, { weight: 'Bolder', spacing: 'Medium' }));
      body.push(tb(md(text), { spacing: 'None' }));
    }
    if (opts.thumbnailDataUrl) body.push({ type: 'Image', url: opts.thumbnailDataUrl, altText: 'Part photo', size: 'Large', spacing: 'Medium' });
    return {
      type: 'message',
      attachments: [{
        contentType: 'application/vnd.microsoft.card.adaptive',
        contentUrl: null,
        content: { $schema: 'http://adaptivecards.io/schemas/adaptive-card.json', type: 'AdaptiveCard', version: '1.4', msteams: { width: 'Full' }, body }
      }]
    };
  }

  function buildTestPayload(name) {
    return {
      type: 'message',
      attachments: [{
        contentType: 'application/vnd.microsoft.card.adaptive', contentUrl: null,
        content: {
          $schema: 'http://adaptivecards.io/schemas/adaptive-card.json', type: 'AdaptiveCard', version: '1.4',
          body: [{ type: 'TextBlock', wrap: true, text: `**${name || 'Someone'}** connected the Additive Applications app to this chat. New entries will appear here.` }]
        }
      }]
    };
  }

  const api = {
    REQUIRED, SUPPORTING, ALL, KEYS, byKey, TEMPLATE_ROWS,
    clean, scaleValue, confidenceMarks, urgencyHue, urgencyHex,
    buildPrompt, parseAIResponse, buildPartOrientationPrompt, parseOrientation, validate, shiftDown, shiftUp, derive,
    buildTeamsPayload, buildTestPayload
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.AAFields = api;
})(typeof window !== 'undefined' ? window : globalThis);
