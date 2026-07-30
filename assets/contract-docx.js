/* بنّاء ملف Word لعقد غين — يعمل في المتصفح (window.buildGainContractDoc) وفي Node (module.exports).
   يستقبل مكتبة docx وبيانات الصفحة ويعيد Document بنفس التنسيق المعتمد. */
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
    const H3 = (t) => new Paragraph({
      bidirectional: true, spacing: { before: 200, after: 100 },
      children: [rtl(t, { font: 'Cairo', size: 23, bold: true, color: '20346A' })],
    });
    const seg = (parts) => parts.map(([t, b]) => rtl(t, { bold: b }));
    const NUM = (t, ref) => new Paragraph({
      bidirectional: true, numbering: { reference: ref, level: 0 }, spacing: { after: 100, line: 320 },
      children: Array.isArray(t) ? t : [rtl(t)],
    });
    const BUL = (t) => new Paragraph({
      bidirectional: true, numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 320 },
      children: Array.isArray(t) ? t : [rtl(t)],
    });

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
    const half = Math.floor(CONTENT / 2);

    // ---------- header ----------
    body.push(new Paragraph({ bidirectional: true, alignment: AlignmentType.CENTER, spacing: { after: 160 },
      children: [new ImageRun({ type: 'png', data: D.images.logo, transformation: { width: 132, height: 56 } })] }));
    body.push(P('غين لتقنية المعلومات', { align: AlignmentType.CENTER, run: { color: SOFT, size: 20 }, after: 40 }));
    body.push(new Paragraph({ bidirectional: true, alignment: AlignmentType.CENTER, spacing: { after: 60 },
      children: [rtl('عقد تقديم خدمات نظام GAIN لإدارة وتشغيل العيادات', { font: 'Cairo', size: 34, bold: true, color: NAVY })] }));
    body.push(P([rtl('رقم العقد: ', { color: SOFT, size: 20 }), rtl(D.no, { bold: true, size: 20 }),
                 rtl('        تاريخ الإبرام: ', { color: SOFT, size: 20 }), rtl(D.date, { bold: true, size: 20 })],
               { align: AlignmentType.CENTER, after: 200 }));

    // ---------- 1 ----------
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

    // ---------- 2 ----------
    body.push(H2('المادة (2): التعريفات'));
    [['الفرع:', ' موقع تشغيل مستقل للعميل ضمن النظام.'],
     ['الاشتراك:', ' المقابل المالي المستحق لقاء توفير الخدمة وتشغيل النظام وفق هذا العقد.'],
     ['التكاملات:', ' الربط الفني مع الأنظمة الحكومية/بوابات الدفع/الأنظمة المساندة وفق ما ورد في الملحقات.'],
     ['SLA:', ' اتفاقية مستوى الخدمة الموضحة في ملحق (SLA).'],
     ['يوم عمل:', ' من الأحد إلى الخميس (باستثناء الإجازات الرسمية).'],
     ['عدد لا محدود من المستخدمين:', ' يُقصد به عدد حسابات المستخدمين داخل نظام GAIN، ولا يشمل ذلك أي رسوم أو حدود تخص أنظمة أو خدمات خارجية.'],
     ['النظام:', ' منصة GAIN وجميع مكوناتها البرمجية والخدمات المرتبطة بها محل هذا العقد.'],
    ].forEach(([b, t]) => body.push(BUL(seg([[b, true], [t, false]]))));

    // ---------- 3 ----------
    body.push(H2('المادة (3): موضوع العقد ونطاق الخدمة'));
    let r = newRef();
    body.push(NUM(seg([['يلتزم المزود بتقديم خدمة تشغيل نظام GAIN للعميل وفق نطاق الخدمة الموضح في ', false], ['الملحق الفني (أ)', true], ['، مع الاستضافة السحابية المُدارة والنسخ الاحتياطي الدوري والتحديثات الدورية والدعم الفني وفق ', false], ['ملحق (SLA)', true], ['.', false]]), r));
    body.push(NUM(seg([['تشمل الخدمة خطة الإطلاق والتفعيل حسب ما يرد في الملحق الفني (أ)، ولا توجد ', false], ['رسوم تأسيس', true], ['.', false]]), r));
    body.push(NUM(seg([['تشمل الخدمة ', false], ['التحديثات الدورية وإصلاح الأعطال ضمن النسخة القياسية', true], ['، ولا تشمل ', false], ['تطوير وظائف جديدة أو تخصيصات خاصة', true], [' إلا بعرض مستقل.', false]]), r));

    // ---------- 4 ----------
    body.push(H2('المادة (4): مدة العقد والتجديد'));
    r = newRef();
    body.push(NUM(seg([['مدة هذا العقد ', false], ['سنة واحدة', true], [' تبدأ من تاريخ العقد.', false]]), r));
    body.push(NUM(seg([['يتجدد العقد تلقائيًا لمدة مماثلة ما لم يُخطر أحد الطرفين الطرف الآخر برغبته في عدم التجديد قبل ', false], ['30 يومًا', true], [' من نهاية المدة.', false]]), r));

    // ---------- 5 ----------
    body.push(H2('المادة (5): الأسعار وآلية الدفع'));
    body.push(H3('5.1 الأسعار'));
    r = newRef();
    body.push(NUM(seg([['الاشتراك الأساسي: 2,499 ريال سعودي شهريًا لكل فرع', true], ['، ويشمل ', false], ['عددًا لا محدودًا من المستخدمين', true], ['.', false]]), r));
    body.push(NUM(seg([['تطبيق الجوال للموظفين «تطبيق غين» (نظام موارد بشرية متكامل) +500 ريال سعودي شهريًا لكل فرع', true], ['، ويشمل ', false], ['عددًا لا محدودًا من المستخدمين', true], ['.', false]]), r));
    body.push(NUM(seg([['جميع الأسعار ', false], ['لا تشمل ضريبة القيمة المضافة', true], [' وتُضاف وفق الأنظمة المعمول بها.', false]]), r));
    body.push(NUM(seg([['يشمل الاشتراك', true], [' التفعيل المبدئي، والتدريب، ودعم الإطلاق ضمن الخطة المعتمدة.', false]]), r));

    body.push(P('ملخص الاشتراك وفق اختيارات العميل:', { before: 140, run: { color: SOFT, size: 20 } }));
    const sumRow = (k, v, strong) => new TableRow({ children: [
      cell([P(k, { run: { size: 20, color: strong ? 'FFFFFF' : SOFT, bold: !!strong }, after: 20 })], { w: 6200, fill: strong ? '20346A' : undefined }),
      cell([P(v, { run: { size: 21, bold: true, color: strong ? 'FFFFFF' : NAVY }, after: 20 })], { w: 3306, fill: strong ? '20346A' : undefined }),
    ]});
    body.push(new Table({
      visuallyRightToLeft: true, columnWidths: [6200, 3306], width: { size: CONTENT, type: WidthType.DXA },
      rows: [
        sumRow('عدد الفروع', D.calc.branches),
        sumRow('الاشتراك الأساسي (لكل فرع/شهر)', '2,499 ر.س'),
        sumRow('تطبيق غين (لكل فرع/شهر)', D.calc.app),
        sumRow('الإجمالي الشهري (قبل الضريبة)', D.calc.monthly),
        sumRow('دورية الفوترة', D.calc.cycle),
        sumRow('قيمة الدورة (قبل الضريبة)', D.calc.cycleAmt),
        sumRow('ضريبة القيمة المضافة (15%)', D.calc.vat),
        sumRow('الإجمالي المستحق لكل دورة (شامل الضريبة)', D.calc.total, true),
      ],
    }));

    body.push(H3('5.2 طريقة الدفع ودورية الفوترة'));
    r = newRef();
    body.push(NUM(seg([['يتم الدفع ', false], ['شهريًا', true], [' أو ', false], ['ربع سنوي', true], [' أو ', false], ['نصف سنوي', true], [' أو ', false], ['سنوي', true], [' حسب اختيار العميل.', false]]), r));
    body.push(NUM(seg([['يصدر المزود فاتورة في بداية كل دورة اشتراك ويتم السداد ', false], ['مقدمًا', true], [' عن الدورة.', false]]), r));
    body.push(H3('5.3 إشعار الفواتير'));
    body.push(P(seg([['يقوم المزود بإشعار العميل بإصدار الفواتير وإرسالها عبر ', false], ['الوسيلة التي يحددها العميل', true], [' كتابيًا عند التعاقد، ويُعد الإشعار صحيحًا ومنتجًا لآثاره متى ما تم الإرسال عبر تلك الوسيلة.', false]])));
    body.push(P(seg([['وفي حال تغيير العميل للوسيلة المعتمدة، يلتزم بإبلاغ المزود كتابيًا قبل ', false], ['(5) أيام عمل', true], [' من تاريخ الاستحقاق.', false]])));
    body.push(H3('5.4 استحقاق الدفعة ومدة السداد'));
    r = newRef();
    body.push(NUM('تستحق قيمة الفاتورة فور إصدارها.', r));
    body.push(NUM(seg([['يلتزم العميل بسداد الفاتورة خلال ', false], ['(5) أيام عمل', true], [' من تاريخ الإصدار، ما لم يتفق الطرفان كتابيًا على خلاف ذلك.', false]]), r));
    body.push(H3('5.5 التأخر في السداد'));
    r = newRef();
    body.push(NUM(seg([['في حال تأخر العميل عن السداد لأكثر من ', false], ['(10) أيام عمل', true], [' من تاريخ الاستحقاق، يحق للمزود ', false], ['تعليق الخدمة مؤقتًا', true], [' بعد إشعار العميل، لحين السداد.', false]]), r));
    body.push(NUM(seg([['في حال استمرار التأخر لأكثر من ', false], ['(30) يومًا', true], [' من تاريخ الاستحقاق، يحق للمزود إنهاء العقد وفق المادة (18)، مع احتفاظه بحقه في المطالبة بالمبالغ المستحقة حتى تاريخ الإنهاء.', false]]), r));
    body.push(NUM('لا يترتب على تعليق الخدمة أي تمديد تلقائي لمدة الاشتراك ما لم يتفق الطرفان كتابيًا.', r));
    body.push(NUM(seg([['تُستأنف الخدمة خلال ', false], ['(1) يوم عمل', true], [' من تأكيد استلام المبلغ المستحق، ما لم توجد التزامات أخرى قائمة.', false]]), r));
    body.push(H3('5.6 تعديل الأسعار'));
    body.push(P(seg([['يجوز للمزود مراجعة أسعار الاشتراك من وقت لآخر بما يتناسب مع تطوير الخدمات أو تغيّر التكاليف التشغيلية، على أن يسري أي تعديل اعتبارًا من ', false], ['دورة التجديد التالية فقط', true], ['، وبعد إشعار العميل قبل موعد التجديد بما لا يقل عن ', false], ['(30) يومًا', true], ['، ويحق للعميل عندئذٍ قبول التجديد بالأسعار الجديدة أو عدم تجديد الاشتراك دون أي التزام إضافي.', false]])));

    // ---------- 6..21 ----------
    body.push(H2('المادة (6): الاستضافة والنسخ الاحتياطي'));
    r = newRef();
    body.push(NUM(seg([['تُقدَّم الخدمة عبر ', false], ['استضافة سحابية مُدارة من المزود', true], ['.', false]]), r));
    body.push(NUM(seg([['يلتزم المزود بتنفيذ ', false], ['نسخ احتياطي دوري', true], [' وفق سياسات التشغيل المعتمدة لديه.', false]]), r));
    body.push(NUM('يلتزم العميل بإدارة المستخدمين والصلاحيات داخل منشأته بما يضمن حماية بيانات الدخول.', r));

    body.push(H2('المادة (7): التحديثات والصيانة'));
    r = newRef();
    body.push(NUM('يحق للمزود إجراء التحديثات الفنية والأمنية والتطويرية وأعمال الصيانة المجدولة التي يراها ضرورية لتحسين أداء النظام أو المحافظة على أمنه واستقراره.', r));
    body.push(NUM('يحرص المزود — متى كان ذلك ممكنًا — على تنفيذ أعمال الصيانة خارج أوقات الذروة وإشعار العميل مسبقًا إذا كان لها أثر جوهري على الخدمة.', r));
    body.push(NUM('لا تعد التحديثات أو أعمال الصيانة المجدولة التي تتم وفق أحكام هذه المادة إخلالًا بالتزامات المزود التعاقدية، ويجوز أن تتضمن التحديثات تحسينات أو تطويرات أو خصائص جديدة، ويخضع توفيرها للباقات والخطط المعتمدة لدى المزود وقت إطلاقها.', r));

    body.push(H2('المادة (8): التكاملات والرسوم المتعلقة بالجهات والأنظمة'));
    r = newRef();
    body.push(NUM(seg([['يشمل الاشتراك ', false], ['التفعيل والربط الفني', true], [' للتكاملات وفق المتطلبات الفنية وتوفر الصلاحيات.', false]]), r));
    body.push(NUM(seg([['في حال ترتّب على استخدام أي من هذه الأنظمة أو الخدمات ', false], ['رسوم اشتراك أو رسوم تشغيل أو رسوم ربط أو أي رسوم مماثلة', true], [' تفرضها الجهة المالكة للنظام أو مزود الخدمة (حاليًا أو مستقبلًا)، فإن ', false], ['العميل يتحمّل تلك الرسوم', true], ['.', false]]), r));
    body.push(NUM('يلتزم العميل بتوفير التفويضات والصلاحيات والمتطلبات اللازمة للتفعيل لدى الجهات ذات العلاقة عند الحاجة.', r));

    body.push(H2('المادة (9): نقل البيانات من النظام السابق'));
    r = newRef();
    body.push(NUM(seg([['نقل البيانات يعتمد على ', false], ['تقييم فني', true], [' للنظام السابق (نوع البيانات، إمكانية التصدير، الجودة، حجم البيانات).', false]]), r));
    body.push(NUM('بعد التقييم يزوّد المزود العميل بنتيجة وخطة عمل، ويتم توضيح نطاق النقل والمدة المتوقعة والتكلفة واعتمادها قبل التنفيذ.', r));

    body.push(H2('المادة (10): الدعم الفني واتفاقية مستوى الخدمة (SLA)'));
    r = newRef();
    body.push(NUM(seg([['يقدّم المزود الدعم الفني وفق ', false], ['ملحق (SLA)', true], ['.', false]]), r));
    body.push(NUM(seg([['زمن الاستجابة', true], [' يعني بدء التعامل مع البلاغ وتأكيد استلامه وتشخيصه، ولا يعني بالضرورة حلّه فورًا.', false]]), r));
    body.push(NUM('قنوات الدعم يتم الاتفاق عليها عند التعاقد بما يضمن سهولة التواصل وسرعة معالجة البلاغات.', r));
    body.push(NUM('لا يشمل الدعم الفني تطوير خصائص جديدة أو إجراء تعديلات خاصة بطلب العميل أو معالجة الأعطال الناتجة عن أنظمة أو أجهزة أو خدمات خارجية غير معتمدة من المزود، ما لم يتم الاتفاق على ذلك كتابة.', r));

    body.push(H2('المادة (11): التزامات العميل'));
    body.push(P('يلتزم العميل بما يلي:'));
    r = newRef();
    body.push(NUM('توفير البيانات الأساسية المطلوبة للتهيئة وتمكين الوصول الفني عند الحاجة خلال المدة المتفق عليها.', r));
    body.push(NUM('تعيين ممثل/منسق مشروع للتنسيق والاعتمادات.', r));
    body.push(NUM('المحافظة على سرية بيانات الدخول وعدم مشاركتها، واعتماد مستخدمين وصلاحيات مناسبة.', r));
    body.push(NUM(seg([['يلتزم العميل بتوفير البيانات الأساسية المطلوبة للتهيئة خلال ', false], ['(10) أيام عمل', true], [' من تاريخ طلبها، وفي حال التأخر يحق للمزود تعديل جدول التفعيل بما يتناسب مع ذلك.', false]]), r));
    body.push(NUM('يلتزم العميل باستخدام النظام استخدامًا مشروعًا وعدم القيام بأي ممارسة قد تؤثر على أمن النظام أو استقراره أو حقوق المستخدمين الآخرين.', r));

    body.push(H2('المادة (12): التغيير في نطاق العمل وحق الرفض'));
    r = newRef();
    body.push(NUM('أي طلب خارج نطاق هذا العقد وملحقاته يتم توثيقه كنطاق إضافي، ويقدّم المزود للعميل عرضًا بالسعر والمدة، ولا يتم التنفيذ إلا بعد اعتماد العميل.', r));
    body.push(NUM(seg([['للمزود الحق في رفض أي طلب إضافي', true], [' إذا كان غير متاح ضمن منتجاته أو غير ممكن فنيًا أو يتعارض مع سياسات التشغيل أو المتطلبات النظامية.', false]]), r));

    body.push(H2('المادة (13): السرية وحماية البيانات'));
    r = newRef();
    body.push(NUM('يلتزم الطرفان بالمحافظة على سرية جميع المعلومات والبيانات المتبادلة وعدم استخدامها أو الإفصاح عنها إلا بالقدر اللازم لتنفيذ هذا العقد أو تنفيذًا لالتزام نظامي.', r));
    body.push(NUM('يلتزم المزود بتطبيق التدابير الفنية والتنظيمية المعقولة لحماية بيانات العميل داخل بيئة التشغيل السحابية من الوصول غير المصرح به أو الفقد أو التعديل أو الإفصاح.', r));
    body.push(NUM(seg([['تبقى جميع بيانات العميل وسجلاته ومحتوياته ', false], ['مملوكة للعميل', true], ['، ولا يكتسب المزود أي حق ملكية عليها، ويقتصر دوره على معالجتها واستضافتها وفق أحكام هذا العقد.', false]]), r));
    body.push(NUM(seg([['يلتزم الطرفان بمعالجة البيانات الشخصية والبيانات الصحية — متى كانت محلًا للخدمة — وفق الأنظمة واللوائح والتعليمات السارية في المملكة العربية السعودية، بما في ذلك ', false], ['نظام حماية البيانات الشخصية', true], [' وأي متطلبات تنظيمية تصدر عن الجهات المختصة.', false]]), r));
    body.push(NUM(seg([['عند انتهاء أو إنهاء هذا العقد يحتفظ المزود ببيانات العميل لمدة ', false], ['(30) يومًا', true], [' من تاريخ انتهاء العقد، ويحق للعميل خلال هذه المدة طلب تصدير بياناته بصيغة إلكترونية معتمدة، وبعد انتهاء هذه المدة يجوز للمزود حذف البيانات أو إتلافها وفق سياسات الاحتفاظ بالبيانات، ما لم يوجد التزام نظامي بخلاف ذلك.', false]]), r));
    body.push(NUM('يلتزم العميل بالحصول على جميع الموافقات أو الأسس النظامية اللازمة لمعالجة البيانات الشخصية أو الصحية الخاصة بمستفيديه داخل النظام، ويتحمل مسؤولية صحة البيانات المدخلة ومشروعية جمعها.', r));

    body.push(H2('المادة (14): قبول التسليم وتفعيل النظام'));
    r = newRef();
    body.push(NUM(seg([['يُعد النظام ', false], ['مقبولًا للتشغيل', true], [' عند اكتمال التهيئة وإتاحة الوصول للمستخدمين وتنفيذ تجربة تشغيل أولية (UAT).', false]]), r));
    body.push(NUM(seg([['يلتزم العميل بمراجعة النظام وتقديم ملاحظاته الجوهرية خلال ', false], ['(5) أيام عمل', true], [' من تاريخ بدء تجربة التشغيل الأولية (UAT).', false]]), r));
    body.push(NUM('يُعد النظام مقبولًا ومفعّلًا عند اعتماد العميل له، أو عند انتهاء المدة المشار إليها في الفقرة السابقة دون تقديم ملاحظات جوهرية.', r));

    body.push(H2('المادة (15): حدود المسؤولية'));
    r = newRef();
    body.push(NUM(seg([['مسؤولية المزود عن أي أضرار أو خسائر مباشرة ناشئة عن هذا العقد تكون في حدود ما يثبت أنه نتج عن خطأ جسيم من المزود، وبحد أقصى لا يتجاوز إجمالي المبالغ المدفوعة من العميل للمزود خلال ', false], ['(3) أشهر', true], [' السابقة لسبب المطالبة.', false]]), r));
    body.push(NUM('لا يتحمل المزود أي خسائر غير مباشرة أو تبعية مثل فوات الفرص أو خسارة الأرباح أو فقدان البيانات الناتج عن طرف خارجي أو استخدام غير صحيح أو انقطاع خدمات خارج نطاق سيطرة المزود.', r));
    body.push(NUM(seg([['لا يتحمل المزود أي مسؤولية عن تعطل أو فشل أي تكامل خارجي أو خدمات طرف ثالث خارجة عن سيطرته', true], ['، بما في ذلك الانقطاعات أو الأعطال أو التغييرات التي تطرأ من مزودي الخدمات أو الجهات المالكة للأنظمة.', false]]), r));
    body.push(NUM('كما لا يتحمل المزود المسؤولية عن أي انقطاع أو خلل ناتج عن خدمات الإنترنت لدى العميل، أو أعطال أجهزته، أو البرامج أو الخدمات المقدمة من أطراف ثالثة، أو أي استخدام مخالف لتعليمات التشغيل.', r));

    body.push(H2('المادة (16): الملكية الفكرية'));
    r = newRef();
    body.push(NUM(seg([['جميع حقوق الملكية الفكرية المتعلقة بمنصة GAIN، بما في ذلك البرامج، والأكواد المصدرية والتنفيذية، وقواعد البيانات، والتصاميم، وواجهات المستخدم، والخوارزميات، والوثائق الفنية، والعلامات التجارية، وأي تحديثات أو تطويرات أو تحسينات مستقبلية، هي ', false], ['ملك حصري للمزود', true], [' وتبقى جميع حقوقها محفوظة له.', false]]), r));
    body.push(NUM(seg([['يمنح هذا العقد العميل حقًا ', false], ['غير حصري وغير قابل للتنازل أو الترخيص من الباطن', true], [' لاستخدام النظام طوال مدة الاشتراك ووفقًا لأحكام هذا العقد، ولا يترتب على ذلك نقل أو التنازل عن أي من حقوق الملكية الفكرية إلى العميل، وتنتهي حقوق الاستخدام بانتهاء أو إنهاء العقد.', false]]), r));
    body.push(NUM('لا يجوز للعميل أو لأي طرف يعمل لصالحه نسخ النظام أو إعادة إنتاجه أو تعديله أو إجراء الهندسة العكسية عليه أو محاولة استخراج الكود المصدري أو إعادة بيعه أو تأجيره أو إتاحته للغير أو إنشاء أعمال مشتقة منه أو إزالة أي إشعارات تتعلق بحقوق الملكية الفكرية، إلا بموافقة كتابية مسبقة من المزود.', r));

    body.push(H2('المادة (17): القوة القاهرة وتعذر الخدمة'));
    r = newRef();
    body.push(NUM('لا يعد أي طرف مسؤولًا عن التأخير أو عدم التنفيذ الناتج عن ظروف خارجة عن إرادته (القوة القاهرة) مثل تعطل خدمات الاتصالات أو انقطاع مزودي الخدمات السحابية أو قرارات الجهات التنظيمية أو الأعطال العامة أو الكوارث.', r));
    body.push(NUM('يتم إخطار الطرف الآخر، وتُمنح مدة مناسبة لاستئناف الخدمة.', r));
    body.push(NUM(seg([['إذا استمر التعذر لمدة تتجاوز ', false], ['(30) يومًا', true], [' جاز لأي طرف إنهاء العقد دون مسؤولية، مع تسوية المستحقات حتى تاريخ الإنهاء.', false]]), r));

    body.push(H2('المادة (18): الإنهاء'));
    r = newRef();
    body.push(NUM(seg([['يجوز لأي من الطرفين إنهاء هذا العقد إذا أخل الطرف الآخر إخلالًا جوهريًا بأي من التزاماته ولم يقم بمعالجة الإخلال خلال ', false], ['(15) يومًا', true], [' من تاريخ إشعاره كتابة.', false]]), r));
    body.push(NUM('لا يؤثر إنهاء العقد على الحقوق أو الالتزامات أو المستحقات المالية الناشئة قبل تاريخ الإنهاء.', r));
    body.push(NUM('يجوز للمزود تعليق الخدمة مؤقتًا، كليًا أو جزئيًا، إذا تبين وجود: (أ) استخدام غير مشروع للنظام؛ (ب) إساءة استخدام تؤثر على أمن أو استقرار المنصة؛ (ج) محاولات اختراق أو هجمات إلكترونية أو مخاطر أمنية تستوجب ذلك؛ (د) مخالفة جوهرية لأحكام هذا العقد.', r));
    body.push(NUM('ويحرص المزود — متى كان ذلك ممكنًا — على إشعار العميل قبل التعليق أو فور اتخاذ الإجراء، ومنحه فرصة لمعالجة المخالفة إذا كانت طبيعتها تسمح بذلك.', r));
    body.push(NUM('لا يعد تعليق الخدمة إنهاءً للعقد، وتستأنف الخدمة بعد زوال سبب التعليق واستيفاء المتطلبات اللازمة لإعادة التشغيل.', r));

    body.push(H2('المادة (19): القانون والاختصاص'));
    body.push(P(seg([['يخضع هذا العقد ويُفسّر وفق ', false], ['القانون السعودي', true], ['، ويُفصل في أي نزاع وفق الأنظمة النافذة في المملكة العربية السعودية.', false]])));

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

    // ---------- 21 ----------
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

    // ---------- Annex A ----------
    body.push(H2('الملحق الفني (أ): نطاق منصة GAIN وخطة التفعيل'));
    body.push(H3('أولًا: نطاق منصة GAIN لإدارة وتشغيل العيادات يشمل'));
    ['إدارة المرضى والمواعيد والسجلات الطبية (EMR / PMS)', 'إدارة الفوترة والمحاسبة', 'إدارة المخزون والمشتريات', 'إدارة الموارد البشرية', 'إدارة العمليات اليومية والتشغيل', 'التقارير ولوحات المتابعة', 'التكامل مع الأنظمة الحكومية والمالية', 'تطبيق الجوال «تطبيق غين» كامتداد تشغيلي للنظام (اختياري)'].forEach(t => body.push(BUL(t)));
    body.push(H3('ثانيًا: عناصر النطاق التفصيلية'));
    r = newRef();
    ['نظام إدارة شؤون المرضى (PMS).', 'إدارة المواعيد.', 'السجلات الطبية الإلكترونية (EMR).', 'الفوترة والمحاسبة.', 'المخزون والمشتريات.', 'الموارد البشرية (HR).', 'التقارير ولوحات المتابعة.', 'التهيئة والإعدادات (ضمن الاشتراك).', 'تدريب المستخدمين وإدارة النظام ودعم الإطلاق مشمولة ضمن الاشتراك وفق خطة التفعيل المعتمدة.', 'الاستضافة السحابية المُدارة والنسخ الاحتياطي الدوري.'].forEach(t => body.push(NUM(t, r)));
    body.push(H3('ثالثًا: خطة التفعيل المبدئية'));
    body.push(P(seg([['يتفق الطرفان على أن خطة التفعيل التالية ', false], ['مبدئية', true], ['، ويتم اعتمادها نهائيًا بعد مراجعة العميل وتأكيد الجاهزية.', false]])));
    [['المرحلة (1): المواءمة وتجهيز البيانات', ['اجتماع مواءمة لتثبيت نطاق التهيئة وتحديد ممثل العميل والمتطلبات.', 'استلام بيانات العيادة الأساسية (الخدمات، الأطباء، جداول العمل، الصلاحيات، القوالب).']],
     ['المرحلة (2): مراجعة البيانات السابقة وتقييم الترحيل', ['يقوم المزود بمراجعة عينة/تصدير من البيانات المتوفرة لدى العميل من النظام السابق.', 'بناءً على المراجعة، يتم تحديد نطاق الترحيل وإمكانية التنفيذ والمدة التقديرية والجهد المتوقع، وتوضيح التكلفة قبل البدء.']],
     ['المرحلة (3): التهيئة والتجربة الأولية', ['تهيئة النظام وفق البيانات المعتمدة وتفعيل الصلاحيات الأساسية.', 'تجربة تشغيل أولية (UAT) لمعالجة الملاحظات الأساسية.']],
     ['المرحلة (4): التدريب والإطلاق', ['تنفيذ التدريب وإدارة النظام ودعم الإطلاق وفق خطة التفعيل المعتمدة.', 'الإطلاق الرسمي والمتابعة ضمن خطة الإطلاق.']],
    ].forEach((ph) => {
      body.push(new Paragraph({ bidirectional: true, spacing: { before: 140, after: 60 }, children: [rtl(ph[0], { bold: true, size: 21 })] }));
      ph[1].forEach(t => body.push(BUL(t)));
    });
    body.push(P(seg([['ملاحظة اعتماد:', true], [' يعتمد الجدول الزمني التفصيلي وآلية التنفيذ النهائية بعد اعتماد خطة التفعيل من العميل وتوفر البيانات والصلاحيات اللازمة.', false]]), { run: { color: SOFT, size: 20 } }));
    body.push(P('للمزود الحق في تطوير وتحسين الوظائف القياسية للنظام بصورة مستمرة، بما يسهم في رفع جودة الخدمة وكفاءة الأداء، ولا يترتب على ذلك التزام بتوفير خصائص أو تطويرات خاصة ما لم يتم الاتفاق عليها كتابة.', { run: { color: SOFT, size: 20 } }));
    body.push(H3('رابعًا: تطبيق الجوال «تطبيق غين» (اختياري للموظفين)'));
    body.push(P('يوفر تطبيق غين وصولًا متكاملًا لموظفي العميل إلى نظام GAIN عبر الأجهزة الذكية، ويُعد امتدادًا تشغيليًا للنظام يتيح تنفيذ وإدارة العمليات اليومية من أي مكان. ويشمل التطبيق المزايا التالية:'));
    ['إدارة الموارد البشرية (الحضور والانصراف، الإجازات، الرواتب، المصاريف)', 'إدارة المبيعات (إنشاء العروض والطلبات ومتابعة العملاء)', 'إدارة المهام والمشاريع ومتابعة الأداء', 'تنفيذ الإجراءات التشغيلية والموافقات (Workflows)', 'إشعارات وتنبيهات فورية', 'دعم العمل الميداني باستخدام تحديد الموقع الجغرافي'].forEach(t => body.push(BUL(t)));
    body.push(P('كما تشمل مزايا التطبيق أي مزايا قياسية أخرى ضمن إصدار التطبيق المعتمد لدى المزود وقت التفعيل. ويهدف التطبيق إلى تمكين التشغيل الرقمي الكامل للعيادة وتحسين كفاءة الفرق التشغيلية وتقليل الاعتماد على الإجراءات اليدوية.', { run: { color: SOFT, size: 20 } }));

    // ---------- SLA ----------
    body.push(H2('ملحق (SLA): اتفاقية مستوى الخدمة'));
    body.push(H3('1) ساعات الدعم'));
    body.push(P(seg([['الأحد – الخميس: ', false], ['9:00 ص – 6:00 م', true], [' (قابلة للتعديل بالاتفاق).', false]])));
    body.push(H3('2) تصنيف البلاغات وزمن الاستجابة'));
    const SLA_W = [2100, 4900, 2506];
    const slaCell = (t, w, b) => cell([P(t, { run: { size: 20, bold: !!b }, after: 20 })], { w });
    body.push(new Table({
      visuallyRightToLeft: true, columnWidths: SLA_W, width: { size: CONTENT, type: WidthType.DXA },
      rows: [
        new TableRow({ tableHeader: true, children: ['التصنيف', 'الوصف', 'زمن الاستجابة'].map((t, i) =>
          cell([P(t, { run: { font: 'Cairo', bold: true, size: 20, color: NAVY }, after: 20 })], { w: SLA_W[i], fill: BGL })) }),
      ].concat([
        ['حرِج', 'تعطل كامل / توقف خدمة رئيسية', 'خلال ساعتين ضمن ساعات الدعم'],
        ['عالٍ', 'تأثير كبير على التشغيل مع وجود بديل مؤقت', 'خلال 4 ساعات'],
        ['متوسط', 'تأثير محدود على جزء من الوظائف', 'خلال يوم عمل'],
        ['منخفض', 'طلبات تحسين / استفسارات / ملاحظات', 'خلال يومي عمل'],
      ].map(rw => new TableRow({ children: [slaCell(rw[0], SLA_W[0], true), slaCell(rw[1], SLA_W[1]), slaCell(rw[2], SLA_W[2], true)] }))),
    }));
    body.push(P('ملاحظة: زمن الاستجابة لا يعني زمن الحل، وزمن الحل يعتمد على طبيعة البلاغ وتوفر البيانات وتعاون الأطراف.', { before: 120, run: { color: SOFT, size: 20 } }));
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
