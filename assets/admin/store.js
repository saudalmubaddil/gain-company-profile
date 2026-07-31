/* ==================================================================
   GainStore — تخزين عقود لوحة التحكم في متصفح المستخدم (localStorage)
   - فهرس خفيف يُقرأ عند كل عرض + سجل كامل لكل عقد + توقيع منفصل (ثقيل)
   - ترحيل غير هدّام من سجل الروابط القديم (gain-admin-history)
   - حماية من امتلاء مساحة التخزين برسالة عربية واضحة
   المفاتيح:
     gain-dash-index          فهرس العقود {v:1, contracts:[...]}
     gain-dash-rec:<id>       السجل الكامل للعقد
     gain-dash-rec:<id>:sig   توقيع العميل (صورة) — يُحمّل عند الطلب فقط
   ================================================================== */
(function () {
  'use strict';

  const INDEX_KEY = 'gain-dash-index';
  const REC_PREFIX = 'gain-dash-rec:';
  const MIGRATED_KEY = 'gain-dash-migrated';
  const OLD_HIST_KEY = 'gain-admin-history';

  const QUOTA_MSG = 'مساحة التخزين في المتصفح ممتلئة. الحل: من «الإعدادات» صدّر نسخة احتياطية ثم احذف عقودًا قديمة، وأعد المحاولة.';

  /* حالات العقد المخزنة. «منتهي» تُشتق من التاريخ ولا تُخزن إلا عند الإنهاء اليدوي */
  const STATUS = {
    draft:   { key: 'draft',   label: 'مسودة' },
    review:  { key: 'review',  label: 'قيد المراجعة' },
    signed:  { key: 'signed',  label: 'موقّع' },
    expired: { key: 'expired', label: 'منتهي' }, // يدوي فقط (عقد مفسوخ)
  };

  /* ---------- أدوات داخلية ---------- */
  function safeSet(key, value) {
    try { localStorage.setItem(key, value); return { ok: true }; }
    catch (e) { return { ok: false, error: QUOTA_MSG }; }
  }
  function readJSON(key, fallback) {
    try { return JSON.parse(localStorage.getItem(key)) ?? fallback; } catch (e) { return fallback; }
  }
  /* تنسيق تاريخ محلي yyyy-mm-dd (لا نستخدم toISOString لأنه بتوقيت UTC ويُرجع اليوم السابق أحيانًا) */
  const fmtLocalDate = d => d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  const today = () => fmtLocalDate(new Date());

  /* إضافة أشهر لتاريخ ISO (yyyy-mm-dd) */
  function addMonths(iso, months) {
    if (!iso) return '';
    const d = new Date(iso + 'T00:00');
    if (isNaN(d)) return '';
    d.setMonth(d.getMonth() + months);
    return fmtLocalDate(d);
  }

  /* ---------- الفهرس ---------- */
  function readIndex() {
    const idx = readJSON(INDEX_KEY, null);
    return (idx && Array.isArray(idx.contracts)) ? idx : { v: 1, contracts: [] };
  }
  function writeIndex(idx) { return safeSet(INDEX_KEY, JSON.stringify(idx)); }

  function list() { return readIndex().contracts.slice(); }
  function getMeta(id) { return readIndex().contracts.find(c => c.id === id) || null; }
  function getRecord(id) { return readJSON(REC_PREFIX + id, null); }
  function getSignature(id) { return localStorage.getItem(REC_PREFIX + id + ':sig') || null; }

  /* إنشاء/تحديث عقد: meta يدخل الفهرس، rec السجل الكامل */
  function upsert(meta, rec) {
    const idx = readIndex();
    const now = new Date().toISOString();
    const i = idx.contracts.findIndex(c => c.id === meta.id);
    const merged = Object.assign(
      i >= 0 ? idx.contracts[i] : { createdAt: now },
      meta, { updatedAt: now });
    if (i >= 0) idx.contracts[i] = merged; else idx.contracts.unshift(merged);
    if (rec) {
      const r = safeSet(REC_PREFIX + meta.id, JSON.stringify(Object.assign({ id: meta.id }, rec)));
      if (!r.ok) return r;
    }
    return writeIndex(idx);
  }

  function setSignature(id, dataURL) {
    if (!dataURL) { localStorage.removeItem(REC_PREFIX + id + ':sig'); return { ok: true }; }
    return safeSet(REC_PREFIX + id + ':sig', dataURL);
  }

  function remove(id) {
    const idx = readIndex();
    idx.contracts = idx.contracts.filter(c => c.id !== id);
    localStorage.removeItem(REC_PREFIX + id);
    localStorage.removeItem(REC_PREFIX + id + ':sig');
    return writeIndex(idx);
  }

  /* ---------- الحالة الفعلية (منتهي يُشتق من التاريخ) ---------- */
  function effectiveStatus(meta) {
    if (meta.status === 'expired') return 'expired';
    if ((meta.status === 'signed' || meta.status === 'review') && meta.expiryDate && meta.expiryDate < today()) return 'expired';
    return meta.status || 'draft';
  }
  /* ينتهي خلال 30 يومًا (مهلة إشعار عدم التجديد في المادة 4) */
  function isExpiringSoon(meta) {
    if (!meta.expiryDate || meta.status === 'draft' || meta.status === 'expired') return false;
    const t = today();
    if (meta.expiryDate < t) return false;
    const soon = new Date(t + 'T00:00');
    soon.setDate(soon.getDate() + 30);
    return meta.expiryDate <= fmtLocalDate(soon);
  }

  /* ---------- ترحيل سجل الروابط القديم (غير هدّام) ---------- */
  function migrate() {
    if (localStorage.getItem(MIGRATED_KEY)) return { migrated: 0 };
    if (localStorage.getItem(INDEX_KEY)) { localStorage.setItem(MIGRATED_KEY, '1'); return { migrated: 0 }; }
    const hist = readJSON(OLD_HIST_KEY, []);
    if (!hist.length) { localStorage.setItem(MIGRATED_KEY, '1'); return { migrated: 0 }; }
    const now = new Date().toISOString();
    const idx = { v: 1, contracts: [] };
    hist.forEach(item => {
      const cid = window.GainLink ? window.GainLink.newId() : Math.random().toString(36).slice(2, 10);
      idx.contracts.push({
        id: cid, tpl: 'services', status: 'review',
        counterparty: item.cli || '', contractNo: '',
        createdAt: now, updatedAt: now,
        startDate: '', expiryDate: '',
        hasLink: !!item.url, pw: !!item.pw, hasSig: false,
        migratedFrom: 'history', legacyWhen: item.when || '',
      });
      safeSet(REC_PREFIX + cid, JSON.stringify({ id: cid, url: item.url || '', fields: {}, statusLog: [{ to: 'review', at: now }] }));
    });
    const r = writeIndex(idx);
    localStorage.setItem(MIGRATED_KEY, '1'); // المفتاح القديم يبقى كما هو (شبكة أمان)
    return r.ok ? { migrated: idx.contracts.length } : { migrated: 0, error: r.error };
  }

  window.GainStore = {
    STATUS, QUOTA_MSG,
    list, getMeta, getRecord, getSignature,
    upsert, setSignature, remove,
    effectiveStatus, isExpiringSoon, addMonths, today,
    migrate,
  };
})();
