// ==UserScript==
// @name         Wali Session 文件上传增强（手机 X浏览器 / 电脑通用）
// @namespace    local.wali.session.fileupload
// @version      3.0.0
// @description  在 Wali 充值页面的“账号 Token JSON”输入框前增加文件上传；只在浏览器本地解析 TXT、JSON、DOCX，不上传到额外服务器。
// @match        https://pro20x.com/automate*
// @match        https://www.pro20x.com/automate*
// @grant        none
// @run-at       document-idle
// ==/UserScript==

(() => {
  'use strict';
  if (!['pro20x.com', 'www.pro20x.com'].includes(location.hostname) || !location.pathname.startsWith('/automate')) return;
  const MARKER = 'data-wali-local-upload';
  const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const MAX_BYTES = 5 * 1024 * 1024;
  const decoder = new TextDecoder('utf-8');

  // .docx 是 ZIP 容器；只在本机取出 word/document.xml。
  async function readDocx(file) {
    const bytes = await file.arrayBuffer();
    const view = new DataView(bytes);
    let end = -1;
    const min = Math.max(0, view.byteLength - 65557);
    for (let i = view.byteLength - 22; i >= min; --i) {
      if (view.getUint32(i, true) === 0x06054b50) { end = i; break; }
    }
    if (end < 0) throw new Error('DOCX 文件结构不正确');
    const directoryOffset = view.getUint32(end + 16, true);
    const count = view.getUint16(end + 10, true);
    let pos = directoryOffset;
    let xml = null;
    for (let i = 0; i < count; ++i) {
      if (pos + 46 > view.byteLength || view.getUint32(pos, true) !== 0x02014b50) break;
      const method = view.getUint16(pos + 10, true);
      const compressedSize = view.getUint32(pos + 20, true);
      const filenameSize = view.getUint16(pos + 28, true);
      const extraSize = view.getUint16(pos + 30, true);
      const commentSize = view.getUint16(pos + 32, true);
      const localOffset = view.getUint32(pos + 42, true);
      const name = decoder.decode(new Uint8Array(bytes, pos + 46, filenameSize));
      if (name === 'word/document.xml') {
        if (localOffset + 30 > view.byteLength || view.getUint32(localOffset, true) !== 0x04034b50) {
          throw new Error('DOCX 正文入口损坏');
        }
        const localNameSize = view.getUint16(localOffset + 26, true);
        const localExtraSize = view.getUint16(localOffset + 28, true);
        const start = localOffset + 30 + localNameSize + localExtraSize;
        if (start + compressedSize > view.byteLength) throw new Error('DOCX 正文数据不完整');
        const compressed = new Uint8Array(bytes, start, compressedSize);
        if (method === 0) {
          xml = decoder.decode(compressed);
        } else if (method === 8 && typeof DecompressionStream !== 'undefined') {
          const stream = new Blob([compressed]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
          xml = await new Response(stream).text();
        } else {
          throw new Error('当前浏览器不支持该 Word 压缩格式，请改发 TXT/JSON');
        }
        break;
      }
      pos += 46 + filenameSize + extraSize + commentSize;
    }
    if (xml === null) throw new Error('无法从 DOCX 找到正文，请改发 TXT/JSON');
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) throw new Error('Word 正文无法解析');
    const paras = doc.getElementsByTagNameNS(WORD_NS, 'p');
    const lines = [];
    for (const p of paras) {
      let line = '';
      // Word 把一段 JSON 分拆成许多 <w:t> 时必须拼回原始顺序。
      for (const el of p.getElementsByTagName('*')) {
        if (el.namespaceURI !== WORD_NS) continue;
        if (el.localName === 't') line += el.textContent || '';
        else if (el.localName === 'tab') line += '\t';
        else if (el.localName === 'br' || el.localName === 'cr') line += '\n';
      }
      lines.push(line);
    }
    return lines.join('\n');
  }

  async function fileToJson(file) {
    if (!file || file.size > MAX_BYTES) throw new Error('文件为空或超过 5MB');
    const lower = file.name.toLowerCase();
    if (/\.doc$/.test(lower)) throw new Error('旧版 .doc 不支持，请另存为 .docx、.txt 或 .json');
    if (!/\.(json|txt|text|docx)$/.test(lower)) throw new Error('只支持 JSON、TXT、TEXT、DOCX 文件');
    let content = lower.endsWith('.docx') ? await readDocx(file) : await file.text();
    content = content.replace(/^\uFEFF/, '').trim();
    const fence = content.match(/^```(?:json)?\s*\n([\s\S]*?)\n```$/i);
    if (fence) content = fence[1].trim();
    let data;
    try { data = JSON.parse(content); }
    catch { throw new Error('文件内容不是有效 JSON，请检查是否复制了完整的 /api/auth/session 内容'); }
    if (!data || typeof data !== 'object' || Array.isArray(data)) throw new Error('Session 必须是 JSON 对象');
    if (typeof data.accessToken !== 'string' || !data.accessToken) {
      throw new Error('文件缺少 accessToken；请确认它是完整的 Session JSON');
    }
    return content;
  }

  function findTarget() {
    const areas = Array.from(document.querySelectorAll('textarea'));
    return areas.find(a => (a.getAttribute('placeholder') || '').includes('/api/auth/session'))
      || areas.find(a => (a.getAttribute('placeholder') || '').includes('Token JSON'))
      || (areas.length === 1 && document.body.textContent.includes('账号 Token JSON') ? areas[0] : null);
  }

  function setReactCompatibleValue(el, value) {
    const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')?.set;
    if (setter) setter.call(el, value);
    else el.value = value;
    // 兼容 React 受控 textarea 的 onChange / input 监听。
    try { el.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText' })); }
    catch { el.dispatchEvent(new Event('input', { bubbles: true })); }
    el.dispatchEvent(new Event('change', { bubbles: true }));
  }

  function mount() {
    const textarea = findTarget();
    if (!textarea || !textarea.parentElement || textarea.parentElement.querySelector('[' + MARKER + ']')) return;

    const bar = document.createElement('div');
    bar.setAttribute(MARKER, '1');
    Object.assign(bar.style, {
      display: 'flex', alignItems: 'center', flexWrap: 'wrap', gap: '8px',
      margin: '8px 0', fontSize: '13px', lineHeight: '1.5'
    });
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'wali-install-upload-btn';
    btn.textContent = '📂 上传 Session 文件';
    Object.assign(btn.style, {
      border: '1px solid #94a3b8', background: '#eff6ff', color: '#0f172a',
      borderRadius: '9px', padding: '9px 12px', cursor: 'pointer',
      fontSize: '14px', fontWeight: '600', touchAction: 'manipulation'
    });
    const status = document.createElement('span');
    status.setAttribute('aria-live', 'polite');
    status.textContent = '支持 JSON / TXT / DOCX，本地读取';
    status.style.color = '#475569';
    const picker = document.createElement('input');
    picker.type = 'file';
    picker.accept = '.json,.txt,.text,.docx';
    picker.style.display = 'none';
    btn.addEventListener('click', () => picker.click());
    picker.addEventListener('change', async () => {
      if (!picker.files?.length) return;
      btn.disabled = true;
      status.textContent = '正在本地解析文件…';
      status.style.color = '#475569';
      try {
        const text = await fileToJson(picker.files[0]);
        // Wali 原有前端继续完成验证、套餐选择和支付，本脚本不会发请求。
        setReactCompatibleValue(textarea, text);
        status.textContent = '✓ 已填入 Token JSON，请等待 Wali 验证通过后再提交';
        if (window.matchMedia('(max-width: 820px)').matches) textarea.scrollIntoView({behavior:'smooth',block:'center'});
        status.style.color = '#047857';
      } catch (e) {
        status.textContent = '✕ ' + (e?.message || '读取失败');
        status.style.color = '#b91c1c';
      } finally {
        picker.value = '';
        btn.disabled = false;
      }
    });
    if (!document.getElementById('wali-mobile-upload-style')) {
      const st = document.createElement('style');
      st.id = 'wali-mobile-upload-style';
      st.textContent = '@media(max-width:820px){.wali-install-upload-btn{position:fixed!important;right:14px!important;bottom:70px!important;z-index:2147483600!important;background:#08745a!important;color:white!important;border-color:#08745a!important;border-radius:26px!important;min-height:48px!important;box-shadow:0 4px 17px #0005!important}}';
      document.head.appendChild(st);
    }
    bar.append(btn, status, picker);
    textarea.parentElement.insertBefore(bar, textarea);
  }
  let scheduled = false;
  const observer = new MutationObserver(() => {
    if (scheduled) return;
    scheduled = true;
    requestAnimationFrame(() => { scheduled = false; mount(); });
  });
  observer.observe(document.documentElement, { childList: true, subtree: true });
  mount();
})();