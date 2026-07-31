/* بنّاء ملف Word لعقد غين — يعمل في المتصفح (window.buildGainContractDoc) وفي Node (module.exports).
   يستقبل مكتبة docx وبيانات الصفحة + نموذج المحتوى الحي (model) ويعيد Document بالتنسيق المعتمد.
   نموذج المحتوى: قائمة كتل بالترتيب —
     {k:'h2', n, runs} {k:'h3'|'h4'|'p', runs, muted?} {k:'ul'|'ol', items:[runs]}
     {k:'table', head, rows} {k:'parties'} {k:'calc'} {k:'bank'} {k:'sign'}
   حيث runs = [{t, b}] */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.buildGainContractDoc = factory();
})(typeof self !== 'undefined' ? self : this, function () {

  return function buildGainContractDoc(docx, data) {
    const {
      Document, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
      WidthType, AlignmentType, HeadingLevel, BorderStyle, LevelFormat, ShadingType,
    } = docx;

    const NAVY = '1A2C5C', SKY = '1D8CBE', SOFT = '475569', LINE = 'E2E8F0', BGL = 'F1F5F9';
    const PAGE_W = 11906, MARG = 1200, CONTENT = PAGE_W - 2 * MARG;
    const D = data;
    const half = Math.floor(CONTENT / 2);

    const rtl = (text, opts = {}) => new TextRun({ text, rightToLeft: true, font: opts.font || 'Tajawal', size: opts.size || 22, bold: !!opts.bold, color: opts.color });
    const P = (children, opts = {}) => new Paragraph({
      bidirectional: true, alignment: opts.align || AlignmentType.START,
      spacing: { after: opts.after != null ? opts.after : 120, before: opts.before || 0, line: 320 },
      children: Array.isArray(children) ? children : [rtl(children, opts.run || {})],
    });
    const H2 = (t) => new Paragraph({
      bidirectional: true, heading: HeadingLevel.HEADING_2, spacing: { before: 320, after: 140 },
      children: [rtl(t, { font: 'Cairo', size: 27, bold: true, color: NAVY })],
    });
    const H3run = (runs, size, color) => new Paragraph({
      bidirectional: true, spacing: { before: 200, after: 100 },
      children: runs.map(r => rtl(r.t, { font: 'Cairo', size, bold: true, color })),
    });
    const runsToTextRuns = (runs, opts = {}) => runs.map(r => rtl(r.t, Object.assign({}, opts, { bold: r.b || opts.bold })));

    const noBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
    const thin = { style: BorderStyle.SINGLE, size: 4, color: LINE };
    const cell = (children, o = {}) => new TableCell({
      width: { size: o.w, type: WidthType.DXA },
      shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill } : undefined,
      borders: o.borders || { top: thin, bottom: thin, left: thin, right: thin },
      margins: o.margins || { top: 90, bottom: 90, left: 130, right: 130 },
      children,
    });

    const numberingConfigs = [];
    let numIdx = 0;
    const newRef = () => { const r = 'n' + (numIdx++); numberingConfigs.push({
      reference: r, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.START,
        style: { paragraph: { indent: { start: 460, hanging: 320 } } } }] }); return r; };
    numberingConfigs.push({ reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.START,
      style: { paragraph: { indent: { start: 460, hanging: 260 } } } }] });

    const body = [];

    // ---------- الترويسة ----------
    body.push(new Paragraph({ bidirectional: true, alignment: AlignmentType.CENTER, spacing: { after: 160 },
      children: [new ImageRun({ type: 'png', data: D.images.logo, transformation: { width: 132, height: 56 } })] }));
    body.push(P('غين لتقنية المعلومات', { align: AlignmentType.CENTER, run: { color: SOFT, size: 20 }, after: 40 }));
    body.push(new Paragraph({ bidirectional: true, alignment: AlignmentType.CENTER, spacing: { after: 60 },
      children: [rtl('عقد تقديم خدمات نظام GAIN لإدارة وتشغيل العيادات', { font: 'Cairo', size: 34, bold: true, color: NAVY })] }));
    body.push(P([rtl('رقم العقد: ', { color: SOFT, size: 20 }), rtl(D.no, { bold: true, size: 20 }),
                 rtl('        تاريخ الإبرام: ', { color: SOFT, size: 20 }), rtl(D.date, { bold: true, size: 20 })],
               { align: AlignmentType.CENTER, after: 200 }));

    // ---------- الكتل الخاصة ----------
    function partiesBlock() {
      body.push(H2('المادة (1): أطراف العقد'));
      body.push(P([rtl('تم إبرام هذا العقد بتاريخ '), rtl(D.date, { bold: true }), rtl(' بين كلٍّ من:')]));
      const partyRow = (k, v1, v2) => new TableRow({ children: [
        cell([P([rtl(k + ': ', { color: SOFT, size: 20 }), rtl(v1, { bold: true, size: 21 })], { after: 20 })], { w: half, borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder } }),
        cell([P([rtl(k + ': ', { color: SOFT, size: 20 }), rtl(v2, { bold: true, size: 21 })], { after: 20 })], { w: half, borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder } }),
      ]});
      body.push(new Table({
        visuallyRightToLeft: true, columnWidths: [half, half], width: { size: CONTENT, type: WidthType.DXA },
        rows: [
          new TableRow({ children: [
            cell([P('الطرف الأول — المزود', { run: { font: 'Cairo', bold: true, color: SKY, size: 22 }, after: 60 })], { w: half, fill: 'EAF4FA', borders: { top: thin, bottom: noBorder, left: thin, right: thin } }),
            cell([P('الطرف الثاني — العميل', { run: { font: 'Cairo', bold: true, color: '15803D', size: 22 }, after: 60 })], { w: half, fill: 'ECF9F0', borders: { top: thin, bottom: noBorder, left: thin, right: thin } }),
          ]}),
          partyRow('الاسم', 'شركة غين لتقنية المعلومات', D.cliName),
          partyRow('الرقم الوطني الموحد', '7025868212', D.cliCR),
          partyRow('العنوان', 'أبي معاذ الأنصاري، حي الربيع', D.cliAddr),
          new TableRow({ children: [
            cell([P([rtl('يمثلها: ', { color: SOFT, size: 20 }), rtl('محمد بن علي الضويلع', { bold: true, size: 21 })], { after: 40 })], { w: half, borders: { top: noBorder, bottom: thin, left: thin, right: thin } }),
            cell([P([rtl('يمثلها: ', { color: SOFT, size: 20 }), rtl(D.cliRep, { bold: true, size: 21 })], { after: 40 })], { w: half, borders: { top: noBorder, bottom: thin, left: thin, right: thin } }),
          ]}),
        ],
      }));
      body.push(P('ويُشار إلى الطرفين مجتمعين بـ «الطرفين».', { before: 120, run: { color: SOFT } }));
      body.push(P('ويقر الطرفان بأن هذا العقد وملاحقه يمثلان كامل الاتفاق بينهما، وأن تنفيذهما يكون وفق مبادئ حسن النية والتعاون بما يحقق المصلحة المشتركة، دون الإخلال بحقوق والتزامات كل طرف الواردة فيهما.'));
    }

    function calcBlock() {
      const GREEN = '15803D';
      const sumRow = (k, v, strong, color) => new TableRow({ children: [
        cell([P(k, { run: { size: 20, color: strong ? 'FFFFFF' : (color || SOFT), bold: !!strong || !!color }, after: 20 })], { w: 6200, fill: strong ? '20346A' : undefined }),
        cell([P(v, { run: { size: 21, bold: true, color: strong ? 'FFFFFF' : (color || NAVY) }, after: 20 })], { w: 3306, fill: strong ? '20346A' : undefined }),
      ]});
      const rows = [
        sumRow('عدد الفروع', D.calc.branches),
        sumRow('الاشتراك الأساسي (لكل فرع/شهر)', D.prices.base),
        sumRow('تطبيق غين (لكل فرع/شهر)', D.calc.app),
        sumRow('الإجمالي الشهري (قبل الضريبة)', D.calc.monthly),
        sumRow('دورية الفوترة', D.calc.cycle),
        sumRow('قيمة الدورة (قبل الضريبة)', D.calc.cycleAmt),
      ];
      if (D.calc.disc) {
        rows.push(sumRow(D.calc.disc.label, D.calc.disc.amount, false, GREEN));
        if (D.calc.disc.reason) rows.push(sumRow('سبب الخصم', D.calc.disc.reason, false, GREEN));
        rows.push(sumRow('قيمة الدورة بعد الخصم (قبل الضريبة)', D.calc.disc.after));
      }
      rows.push(
        sumRow('ضريبة القيمة المضافة (' + D.prices.vatPct + '%)', D.calc.vat),
        sumRow('الإجمالي المستحق لكل دورة (شامل الضريبة)', D.calc.total, true),
      );
      body.push(new Table({
        visuallyRightToLeft: true, columnWidths: [6200, 3306], width: { size: CONTENT, type: WidthType.DXA },
        rows,
      }));
    }

    function bankBlock() {
      body.push(H2('المادة (20): الحساب البنكي'));
      body.push(new Table({
        visuallyRightToLeft: true, columnWidths: [3000, 6506], width: { size: CONTENT, type: WidthType.DXA },
        rows: [
          ['اسم الحساب', 'غين لتقنية المعلومات'],
          ['البنك', 'بنك الإنماء'],
          ['رقم الحساب', '68207210157000'],
          ['الآيبان (IBAN)', 'SA5205000068207210157000'],
        ].map((kv) => new TableRow({ children: [
          cell([P(kv[0], { run: { color: SOFT, size: 20 }, after: 20 })], { w: 3000, fill: BGL }),
          cell([P(kv[1], { run: { bold: true, size: 21 }, after: 20 })], { w: 6506 }),
        ]})),
      }));
    }

    function signBlock() {
      body.push(H2('المادة (21): التوقيع'));
      const sigCellPar = (label, value) => P([rtl(label + ': ', { color: SOFT, size: 20 }), rtl(value, { bold: true, size: 21 })], { after: 60 });
      const provSigChildren = [
        P('عن غين لتقنية المعلومات (المزود)', { run: { font: 'Cairo', bold: true, size: 22, color: SKY }, after: 100 }),
        sigCellPar('الاسم', 'محمد بن علي الضويلع'),
        sigCellPar('الصفة', 'الرئيس التنفيذي'),
        P('التوقيع:', { run: { color: SOFT, size: 20 }, after: 40 }),
        new Paragraph({ bidirectional: true, alignment: AlignmentType.CENTER, spacing: { after: 60 },
          children: [new ImageRun({ type: 'png', data: D.images.gainSig, transformation: { width: 150, height: 92 } })] }),
        sigCellPar('التاريخ', D.date),
      ];
      const cliSigChildren = [
        P('عن ' + D.cliPartyName + ' (العميل)', { run: { font: 'Cairo', bold: true, size: 22, color: '15803D' }, after: 100 }),
        sigCellPar('الاسم', D.signName),
        sigCellPar('الصفة', D.signTitle),
        P('التوقيع:', { run: { color: SOFT, size: 20 }, after: 40 }),
      ];
      if (D.images.clientSig) {
        cliSigChildren.push(new Paragraph({ bidirectional: true, alignment: AlignmentType.CENTER, spacing: { after: 40 },
          children: [new ImageRun({ type: 'png', data: D.images.clientSig, transformation: { width: 170, height: D.clientSigH || 65 } })] }));
      } else {
        cliSigChildren.push(P('_______________________________', { align: AlignmentType.CENTER, run: { color: SOFT }, after: 40, before: 260 }));
      }
      cliSigChildren.push(sigCellPar('التاريخ', D.date));
      body.push(new Table({
        visuallyRightToLeft: true, columnWidths: [half, half], width: { size: CONTENT, type: WidthType.DXA },
        rows: [new TableRow({ children: [
          cell(provSigChildren, { w: half, fill: 'F7FBFD', margins: { top: 140, bottom: 140, left: 160, right: 160 } }),
          cell(cliSigChildren, { w: half, fill: 'F8FDF9', margins: { top: 140, bottom: 140, left: 160, right: 160 } }),
        ]})],
      }));
      body.push(P([rtl((D.agree ? '☑' : '☐') + '  ', { size: 26, color: '15803D', bold: true }),
        rtl('أقر بأنني اطّلعت على كامل بنود هذا العقد وملحقاته (الملحق الفني (أ) وملحق مستوى الخدمة SLA)، وأوافق عليها، وأن البيانات المُدخلة صحيحة.', { bold: true, size: 20 })],
        { before: 200 }));
    }

    // ---------- تحويل نموذج المحتوى ----------
    (D.model || []).forEach(blk => {
      switch (blk.k) {
        case 'parties': partiesBlock(); break;
        case 'calc': calcBlock(); break;
        case 'bank': bankBlock(); break;
        case 'sign': signBlock(); break;
        case 'h2': {
          const title = blk.runs.map(r => r.t).join('').trim();
          const isNum = /^\d+$/.test(blk.n);
          body.push(H2(isNum ? 'المادة (' + blk.n + '): ' + title : title));
          break;
        }
        case 'h3': body.push(H3run(blk.runs, 23, '20346A')); break;
        case 'h4': body.push(new Paragraph({ bidirectional: true, spacing: { before: 140, after: 60 },
          children: runsToTextRuns(blk.runs, { bold: true, size: 21 }) })); break;
        case 'p': body.push(new Paragraph({ bidirectional: true, spacing: { after: 120, line: 320 },
          children: runsToTextRuns(blk.runs, blk.muted ? { color: SOFT, size: 20 } : {}) })); break;
        case 'ol': {
          const ref = newRef();
          blk.items.forEach(runs => body.push(new Paragraph({
            bidirectional: true, numbering: { reference: ref, level: 0 }, spacing: { after: 100, line: 320 },
            children: runsToTextRuns(runs) })));
          break;
        }
        case 'ul':
          blk.items.forEach(runs => body.push(new Paragraph({
            bidirectional: true, numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 320 },
            children: runsToTextRuns(runs) })));
          break;
        case 'table': {
          const cols = Math.max(blk.head.length, ...(blk.rows.map(r => r.length)), 1);
          const widths = cols === 3 ? [2100, 4900, 2506] : Array(cols).fill(Math.floor(CONTENT / cols));
          const rows = [];
          if (blk.head.length) rows.push(new TableRow({ tableHeader: true, children: blk.head.map((t, i) =>
            cell([P(t, { run: { font: 'Cairo', bold: true, size: 20, color: NAVY }, after: 20 })], { w: widths[i], fill: BGL })) }));
          blk.rows.forEach(rw => rows.push(new TableRow({ children: rw.map((t, i) =>
            cell([P(t, { run: { size: 20, bold: cols === 3 && (i === 0 || i === cols - 1) }, after: 20 })], { w: widths[i] })) })));
          body.push(new Table({ visuallyRightToLeft: true, columnWidths: widths, width: { size: CONTENT, type: WidthType.DXA }, rows }));
          break;
        }
      }
    });

    body.push(P('www.gain.sa  ✧  info@gain.sa', { align: AlignmentType.CENTER, before: 300, run: { color: SKY, size: 20, bold: true } }));

    return new Document({
      styles: { default: { document: { run: { font: 'Tajawal', size: 22 } } } },
      numbering: { config: numberingConfigs },
      sections: [{
        properties: { page: { margin: { top: MARG, bottom: MARG, left: MARG, right: MARG } } },
        children: body,
      }],
    });
  };
});
