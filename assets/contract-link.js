/* ==================================================================
   GainLink — أدوات مشتركة لروابط العقود (#c=...)
   تُستخدم في admin.html (توليد الروابط) وcontract.html (قراءتها).
   الصيغ:
     p0.<data>            = إعدادات مضغوطة (deflate-raw) بلا حماية
     p1.<salt>.<iv>.<enc> = إعدادات مشفّرة AES-GCM-256 بكلمة مرور (PBKDF2)
   ================================================================== */
(function () {
  'use strict';

  const PBKDF2_ITERATIONS = 150000;

  const b64u = bytes => btoa(String.fromCharCode(...bytes)).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');

  const b64uToBytes = s => {
    s = s.replace(/-/g, '+').replace(/_/g, '/');
    while (s.length % 4) s += '=';
    const bin = atob(s), u = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
    return u;
  };

  async function deflate(str) {
    const stream = new Blob([new TextEncoder().encode(str)]).stream().pipeThrough(new CompressionStream('deflate-raw'));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  async function inflateBytes(bytes) {
    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  const parseCfgBytes = async bytes => JSON.parse(new TextDecoder().decode(await inflateBytes(bytes)));

  async function deriveKey(pw, salt, usage) {
    const keyMat = await crypto.subtle.importKey('raw', new TextEncoder().encode(pw), 'PBKDF2', false, ['deriveKey']);
    return crypto.subtle.deriveKey(
      { name: 'PBKDF2', salt, iterations: PBKDF2_ITERATIONS, hash: 'SHA-256' },
      keyMat, { name: 'AES-GCM', length: 256 }, false, [usage]);
  }

  /* إعدادات → حمولة رابط (p0 أو p1 حسب وجود كلمة مرور) */
  async function buildPayload(cfg, pw) {
    const packed = await deflate(JSON.stringify(cfg));
    if (!pw) return 'p0.' + b64u(packed);
    const salt = crypto.getRandomValues(new Uint8Array(16));
    const iv = crypto.getRandomValues(new Uint8Array(12));
    const key = await deriveKey(pw, salt, 'encrypt');
    const enc = new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, key, packed));
    return 'p1.' + b64u(salt) + '.' + b64u(iv) + '.' + b64u(enc);
  }

  /* فك حمولة p1 بكلمة المرور — يرمي خطأ إن كانت خاطئة */
  async function decryptPayload(parts, pw) {
    const key = await deriveKey(pw, b64uToBytes(parts[1]), 'decrypt');
    const plain = await crypto.subtle.decrypt({ name: 'AES-GCM', iv: b64uToBytes(parts[2]) }, key, b64uToBytes(parts[3]));
    return parseCfgBytes(new Uint8Array(plain));
  }

  /* معرّف قصير عشوائي للعقود */
  const newId = () => b64u(crypto.getRandomValues(new Uint8Array(6)));

  /* نسخ نص للحافظة مع بديل للمتصفحات القديمة */
  async function copyText(t) {
    try { await navigator.clipboard.writeText(t); return true; }
    catch (e) {
      const ta = document.createElement('textarea'); ta.value = t; document.body.appendChild(ta);
      ta.select(); const ok = document.execCommand('copy'); ta.remove(); return ok;
    }
  }

  window.GainLink = { b64u, b64uToBytes, deflate, inflateBytes, parseCfgBytes, buildPayload, decryptPayload, newId, copyText };
})();
