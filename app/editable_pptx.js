// Runs inside Chromium. Measures one slide at a time and splits it into
// PowerPoint objects: text boxes, native shapes and pictures. Whatever cannot
// be rebuilt natively (gradients, patterns, rotated or clipped decoration)
// stays in a background picture, which is captured after the rebuilt parts are
// hidden from the page.
(() => {
  if (window.__pptx) return;

  const IMPORTANT = 'important';
  const AR = /[؀-ۿݐ-ݿࢠ-ࣿﭐ-﷿ﹰ-﻿]/;
  const NO_TEXT_TAGS = new Set(['SCRIPT', 'STYLE', 'NOSCRIPT', 'TEMPLATE', 'svg', 'SVG', 'CANVAS', 'VIDEO', 'IFRAME', 'OBJECT', 'TEXTAREA', 'SELECT', 'OPTION']);
  const REPLACED_TAGS = new Set(['IMG', 'svg', 'SVG', 'CANVAS', 'VIDEO', 'IFRAME', 'OBJECT', 'EMBED', 'PICTURE', 'INPUT']);

  let slides = [];

  const num = (v) => parseFloat(v) || 0;
  const setImportant = (el, prop, value) => el.style.setProperty(prop, value, IMPORTANT);

  function color(value) {
    const m = /rgba?\(([^)]+)\)/.exec(value || '');
    if (!m) return null;
    const parts = m[1].split(/[\s,/]+/).filter(Boolean).map(Number);
    const alpha = parts.length > 3 ? parts[3] : 1;
    if (!(alpha > 0.004)) return null;
    const hex = parts.slice(0, 3).map((n) => Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, '0')).join('').toUpperCase();
    return { hex, alpha: Math.round(alpha * 1000) / 1000 };
  }

  function effectiveOpacity(el, slide) {
    let o = 1;
    for (let n = el; n && n !== slide.parentElement; n = n.parentElement) o *= num(getComputedStyle(n).opacity === '' ? 1 : getComputedStyle(n).opacity);
    return o;
  }

  function transformedOrFiltered(el, slide) {
    // A rotated/skewed/filtered ancestor means measured boxes are not what is drawn.
    for (let n = el; n && n !== slide; n = n.parentElement) {
      const cs = getComputedStyle(n);
      if ((cs.transform && cs.transform !== 'none') || (cs.filter && cs.filter !== 'none') || (cs.clipPath && cs.clipPath !== 'none') ||
          (cs.mask && cs.mask !== 'none') || (cs.webkitMaskImage && cs.webkitMaskImage !== 'none') || (cs.mixBlendMode && cs.mixBlendMode !== 'normal')) return true;
    }
    return false;
  }

  function inSvg(el) {
    return !!(el.closest && el.closest('svg'));
  }

  function isInlineDisplay(cs) {
    return cs.display === 'inline' || cs.display === 'contents';
  }

  function visible(el, cs) {
    if (cs.display === 'none' || cs.visibility === 'hidden' || cs.visibility === 'collapse') return false;
    const r = el.getBoundingClientRect();
    return r.width > 0 || r.height > 0 || cs.display === 'contents';
  }

  function relRect(r, s) {
    return { x: r.left - s.left, y: r.top - s.top, w: r.width, h: r.height };
  }

  function intersect(a, b) {
    const x = Math.max(a.x, b.x), y = Math.max(a.y, b.y);
    const r = Math.min(a.x + a.w, b.x + b.w), btm = Math.min(a.y + a.h, b.y + b.h);
    return r > x && btm > y ? { x, y, w: r - x, h: btm - y } : null;
  }

  // Ancestors with overflow other than visible clip their descendants.
  function clipRect(el, slide, sRect) {
    let clip = { x: 0, y: 0, w: sRect.width, h: sRect.height };
    for (let n = el.parentElement; n && n !== slide.parentElement; n = n.parentElement) {
      const cs = getComputedStyle(n);
      if (cs.overflowX !== 'visible' || cs.overflowY !== 'visible' || (cs.contain || '').includes('paint')) {
        const r = n.getBoundingClientRect();
        const pad = relRect({
          left: r.left + num(cs.borderLeftWidth), top: r.top + num(cs.borderTopWidth),
          width: r.width - num(cs.borderLeftWidth) - num(cs.borderRightWidth), height: r.height - num(cs.borderTopWidth) - num(cs.borderBottomWidth),
        }, sRect);
        clip = intersect(clip, pad);
        if (!clip) return null;
      }
    }
    return clip;
  }

  /* ---------- pseudo-elements: turn ::before/::after text into real spans ---------- */

  function counterText(value, style) {
    const n = Math.round(value);
    const toAlpha = (k, base) => { let s = ''; let x = k; while (x > 0) { x--; s = String.fromCharCode(base + (x % 26)) + s; x = Math.floor(x / 26); } return s; };
    const toRoman = (k) => {
      const map = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
      let s = '';
      let x = k;
      for (const [v, r] of map) {
        while (x >= v) {
          s += r;
          x -= v;
        }
      }
      return s;
    };
    switch ((style || 'decimal').trim()) {
      case 'decimal-leading-zero': return (n < 0 ? '-' : '') + String(Math.abs(n)).padStart(2, '0');
      case 'arabic-indic': return String(n).replace(/\d/g, (d) => '٠١٢٣٤٥٦٧٨٩'[d]);
      case 'lower-alpha': case 'lower-latin': return toAlpha(n, 97);
      case 'upper-alpha': case 'upper-latin': return toAlpha(n, 65);
      case 'lower-roman': return toRoman(n);
      case 'upper-roman': return toRoman(n).toUpperCase();
      case 'none': return '';
      default: return String(n);
    }
  }

  function parseCounterList(value, defaultValue) {
    const out = [];
    if (!value || value === 'none') return out;
    const tokens = value.trim().split(/\s+/);
    for (let i = 0; i < tokens.length; i++) {
      const name = tokens[i];
      const next = tokens[i + 1];
      if (next !== undefined && /^-?\d+$/.test(next)) { out.push([name, Number(next)]); i++; } else out.push([name, defaultValue]);
    }
    return out;
  }

  // Parses the computed `content` value. Returns null when it holds something
  // other than text (an image, quotes), so the pseudo-element is left alone.
  function contentText(content, el, counters) {
    if (!content || content === 'none' || content === 'normal') return null;
    if (/url\(|gradient\(|image-set\(|open-quote|close-quote/.test(content)) return null;
    let out = '';
    const re = /"((?:[^"\\]|\\.)*)"|'((?:[^'\\]|\\.)*)'|counter\(\s*([\w-]+)\s*(?:,\s*([\w-]+)\s*)?\)|counters\(\s*([\w-]+)\s*,\s*"([^"]*)"\s*(?:,\s*([\w-]+)\s*)?\)|attr\(\s*([\w-]+)\s*\)|(\S+)/g;
    let m;
    while ((m = re.exec(content))) {
      if (m[1] !== undefined || m[2] !== undefined) {
        out += (m[1] !== undefined ? m[1] : m[2]).replace(/\\([0-9a-fA-F]{1,6})\s?/g, (_, h) => String.fromCodePoint(parseInt(h, 16))).replace(/\\(.)/g, '$1');
      } else if (m[3]) out += counterText(counters.get(m[3]) || 0, m[4]);
      else if (m[5]) out += counterText(counters.get(m[5]) || 0, m[7]);
      else if (m[8]) out += el.getAttribute(m[8]) || '';
      else return null;
    }
    return out;
  }

  function snapshotLayout(slide) {
    return Array.from(slide.querySelectorAll('*')).map((el) => [el, el.getBoundingClientRect()]);
  }

  function layoutMoved(before) {
    for (const [el, r] of before) {
      const n = el.getBoundingClientRect();
      if (Math.abs(n.left - r.left) > 1 || Math.abs(n.top - r.top) > 1 || Math.abs(n.width - r.width) > 1 || Math.abs(n.height - r.height) > 1) return true;
    }
    return false;
  }

  function materializePseudos(slide) {
    const counters = new Map();
    const created = [];
    const before = snapshotLayout(slide);
    const walk = (el) => {
      const cs = getComputedStyle(el);
      if (cs.display === 'none') return;
      for (const [name, v] of parseCounterList(cs.counterReset, 0)) counters.set(name, v);
      for (const [name, v] of parseCounterList(cs.counterSet, 0)) counters.set(name, v);
      for (const [name, v] of parseCounterList(cs.counterIncrement, 1)) counters.set(name, (counters.get(name) || 0) + v);
      if (!inSvg(el)) {
        for (const which of ['before', 'after']) {
          const pcs = getComputedStyle(el, '::' + which);
          const text = contentText(pcs.content, el, counters);
          if (text === null || pcs.display === 'none') continue;
          const span = document.createElement('span');
          span.setAttribute('data-pptx-pseudo', which);
          for (let i = 0; i < pcs.length; i++) {
            const prop = pcs[i];
            if (prop === 'content' || prop.startsWith('counter-')) continue;
            setImportant(span, prop, pcs.getPropertyValue(prop));
          }
          span.textContent = text;
          if (which === 'before') el.insertBefore(span, el.firstChild); else el.appendChild(span);
          el.setAttribute('data-pptx-no' + which, '');
          created.push([el, span, which]);
        }
      }
      for (const child of Array.from(el.children)) if (!child.hasAttribute('data-pptx-pseudo')) walk(child);
    };
    walk(slide);
    if (created.length && layoutMoved(before)) {
      // The page relied on structure we changed (for example :first-child); keep the originals.
      for (const [el, span, which] of created) { span.remove(); el.removeAttribute('data-pptx-no' + which); }
      return false;
    }
    return true;
  }

  /* ---------- text ---------- */

  function textTransform(text, tt) {
    if (tt === 'uppercase') return text.toUpperCase();
    if (tt === 'lowercase') return text.toLowerCase();
    if (tt === 'capitalize') return text.replace(/(^|\s)(\S)/g, (_, a, b) => a + b.toUpperCase());
    return text;
  }

  function fontName(family) {
    const first = String(family || '').split(',')[0].trim().replace(/^["']|["']$/g, '');
    if (!first || /^(sans-serif|system-ui|-apple-system|BlinkMacSystemFont|ui-sans-serif)$/i.test(first)) return 'Arial';
    if (/^serif$/i.test(first)) return 'Times New Roman';
    if (/^monospace$/i.test(first)) return 'Courier New';
    if (/^(Helvetica Neue|Helvetica)$/i.test(first)) return 'Arial';
    return first;
  }

  // The nearest box that lays out a text node's line: its closest ancestor
  // that is not display:inline.
  function textContainer(node, slide) {
    let el = node.parentElement;
    while (el && el !== slide && isInlineDisplay(getComputedStyle(el))) el = el.parentElement;
    return el;
  }

  function collectRuns(container, slide, opacityOf) {
    const paragraphs = [[]];
    const textNodes = [];
    const inlineEls = [];
    const walk = (node) => {
      if (node.nodeType === 3) {
        const parent = node.parentElement;
        const cs = getComputedStyle(parent);
        let text = node.nodeValue;
        const pre = /^(pre|pre-wrap|break-spaces)$/.test(cs.whiteSpace);
        if (!pre) text = text.replace(/[\t\n\r\f ]+/g, ' ');
        if (!text) return;
        textNodes.push(node);
        const lines = pre ? text.split('\n') : [text];
        lines.forEach((line, i) => {
          if (i > 0) paragraphs.push([]);
          if (!line) return;
          const c = color(cs.color);
          const underline = /underline/.test(cs.textDecorationLine || cs.textDecoration || '');
          paragraphs[paragraphs.length - 1].push({
            text: textTransform(line, cs.textTransform),
            font: fontName(cs.fontFamily),
            size: num(cs.fontSize),
            bold: num(cs.fontWeight) >= 600 || cs.fontWeight === 'bold',
            italic: cs.fontStyle === 'italic' || cs.fontStyle === 'oblique',
            underline,
            color: c ? c.hex : '000000',
            alpha: Math.round((c ? c.alpha : 1) * opacityOf(parent) * 1000) / 1000,
            letterSpacing: cs.letterSpacing === 'normal' ? 0 : num(cs.letterSpacing),
            lang: AR.test(line) ? 'ar-AE' : 'en-US',
          });
        });
        return;
      }
      if (node.nodeType !== 1) return;
      const el = node;
      if (NO_TEXT_TAGS.has(el.tagName)) return;
      const cs = getComputedStyle(el);
      if (cs.display === 'none') return;
      if (el.tagName === 'BR') { paragraphs.push([]); return; }
      if (el !== container && !isInlineDisplay(cs)) return;
      if (el !== container) inlineEls.push(el);
      if (cs.visibility === 'hidden') return;
      for (const child of el.childNodes) walk(child);
    };
    walk(container);
    // Collapse spaces the way CSS does across run boundaries and at line ends.
    const out = [];
    for (const runs of paragraphs) {
      const cleaned = [];
      let prevSpace = true;
      for (const r of runs) {
        let t = r.text;
        if (prevSpace) t = t.replace(/^ +/, '');
        if (!t) continue;
        prevSpace = / $/.test(t);
        cleaned.push({ ...r, text: t });
      }
      while (cleaned.length) {
        const last = cleaned[cleaned.length - 1];
        while (last.text.endsWith(' ')) last.text = last.text.slice(0, -1);
        if (last.text) break;
        cleaned.pop();
      }
      out.push(cleaned);
    }
    while (out.length && !out[out.length - 1].length) out.pop();
    while (out.length && !out[0].length) out.shift();
    return { paragraphs: out, textNodes, inlineEls };
  }

  function textBox(container, slide, sRect, opacityOf) {
    const cs = getComputedStyle(container);
    const { paragraphs, textNodes, inlineEls } = collectRuns(container, slide, opacityOf);
    if (!paragraphs.some((p) => p.length)) return null;
    const range = document.createRange();
    const rects = [];
    for (const t of textNodes) {
      range.selectNodeContents(t);
      for (const r of range.getClientRects()) if (r.width > 0 && r.height > 0) rects.push(r);
    }
    if (!rects.length) return null;
    let left = Infinity, top = Infinity, right = -Infinity, bottom = -Infinity;
    for (const r of rects) { left = Math.min(left, r.left); top = Math.min(top, r.top); right = Math.max(right, r.right); bottom = Math.max(bottom, r.bottom); }

    // Where the browser broke the lines, word by word.
    const words = [];
    for (const t of textNodes) {
      const re = /\S+/g;
      let m;
      while ((m = re.exec(t.nodeValue))) {
        range.setStart(t, m.index);
        range.setEnd(t, m.index + m[0].length);
        const wr = Array.from(range.getClientRects()).filter((r) => r.width > 0 && r.height > 0);
        // A word broken across lines (at a hyphen) gives one rect per line.
        wr.forEach((r, i) => words.push({ l: r.left, r: r.right, cy: (r.top + r.bottom) / 2, w: r.width, split: i > 0 }));
      }
    }
    const lineList = [];
    for (const w of words) {
      const cur = lineList[lineList.length - 1];
      if (cur && Math.abs(w.cy - cur.cy) <= 4) { cur.l = Math.min(cur.l, w.l); cur.r = Math.max(cur.r, w.r); cur.words.push(w); }
      else lineList.push({ cy: w.cy, l: w.l, r: w.r, words: [w] });
    }
    const lines = Math.max(1, lineList.length);
    const fontSize = num(cs.fontSize);
    let gapSum = 0, gapCount = 0;
    for (const ln of lineList) {
      const ws = ln.words.slice().sort((a, b) => a.l - b.l);
      for (let i = 1; i < ws.length; i++) { const g = ws[i].l - ws[i - 1].r; if (g > 0 && g < fontSize) { gapSum += g; gapCount++; } }
    }
    const space = gapCount ? gapSum / gapCount : fontSize * 0.28;
    const maxLine = lineList.reduce((a, ln) => Math.max(a, ln.r - ln.l), 0) || (right - left);
    // The next word did not fit on each line: the box must stay narrower than that.
    let breakLimit = Infinity;
    for (let i = 0; i + 1 < lineList.length; i++) {
      const next = lineList[i + 1].words[0];
      if (next.split) continue;
      breakLimit = Math.min(breakLimit, (lineList[i].r - lineList[i].l) + space + next.w);
    }

    const box = container.getBoundingClientRect();
    const contentLeft = box.left + num(cs.borderLeftWidth) + num(cs.paddingLeft);
    const contentRight = box.right - num(cs.borderRightWidth) - num(cs.paddingRight);
    const rtl = cs.direction === 'rtl';
    let align = cs.textAlign;
    if (align === 'start' || align === '-webkit-auto' || align === 'auto') align = rtl ? 'right' : 'left';
    else if (align === 'end') align = rtl ? 'left' : 'right';
    else if (align === '-webkit-center') align = 'center';
    else if (align === 'match-parent') align = rtl ? 'right' : 'left';

    const wrap = lines > 1;
    let width;
    if (!wrap) width = maxLine + Math.max(4, maxLine * 0.03);
    else if (align === 'justify') width = contentRight - contentLeft;
    else if (breakLimit === Infinity) width = Math.max(maxLine + 4, contentRight - contentLeft);
    // Halfway between the longest line and the first word that did not fit:
    // the same breaks survive small differences in glyph widths.
    else width = breakLimit > maxLine + 2 ? (maxLine + breakLimit) / 2 : maxLine + 2;
    let x0;
    if (align === 'left' || align === 'justify') x0 = wrap && align === 'justify' ? contentLeft : left;
    else if (align === 'right') x0 = right - width;
    else x0 = (wrap ? (contentLeft + contentRight) / 2 : (left + right) / 2) - width / 2;
    const x1 = x0 + width;

    // Line pitch as the browser laid it out.
    const firstCy = lineList.length ? lineList[0].cy : (top + bottom) / 2;
    const contentH = rects.filter((r) => Math.abs((r.top + r.bottom) / 2 - firstCy) <= 4).reduce((a, r) => Math.max(a, r.height), 0) || (bottom - top);
    let lineHeight = num(cs.lineHeight);
    if (lines > 1) lineHeight = (lineList[lineList.length - 1].cy - lineList[0].cy) / (lines - 1);
    else if (!lineHeight || cs.lineHeight === 'normal') lineHeight = contentH;
    // Browsers split extra line spacing above and below the glyphs; PowerPoint
    // and LibreOffice put it all above. Align the first baseline accordingly.
    const firstTop = rects.filter((r) => Math.abs((r.top + r.bottom) / 2 - firstCy) <= 4).reduce((a, r) => Math.min(a, r.top), Infinity);
    const boxTop = (Number.isFinite(firstTop) ? firstTop : top) + contentH - lineHeight;
    const boxHeight = Math.max(bottom - boxTop, lineHeight * lines);

    const slot = container.closest('[data-slot]');
    return {
      node: { container, inlineEls },
      box: {
        x: x0 - sRect.left, y: boxTop - sRect.top, w: x1 - x0, h: boxHeight,
        wrap, rtl, align, lineHeight, lines,
        name: slot ? slot.getAttribute('data-slot') : (container.getAttribute('data-pptx-pseudo') ? 'Number' : container.tagName.toLowerCase()),
        paragraphs,
      },
    };
  }

  /* ---------- shapes and pictures ---------- */

  function hasBorderPaint(cs, side) {
    return num(cs['border' + side + 'Width']) > 0 && !/^(none|hidden)$/.test(cs['border' + side + 'Style']) && !!color(cs['border' + side + 'Color']);
  }

  function paints(el, cs) {
    if (REPLACED_TAGS.has(el.tagName)) return true;
    if (color(cs.backgroundColor)) return true;
    if (cs.backgroundImage && cs.backgroundImage !== 'none') return true;
    if (['Top', 'Right', 'Bottom', 'Left'].some((s) => hasBorderPaint(cs, s))) return true;
    if (cs.boxShadow && cs.boxShadow !== 'none') return true;
    if (cs.outlineStyle !== 'none' && num(cs.outlineWidth) > 0 && color(cs.outlineColor)) return true;
    for (const which of ['::before', '::after']) {
      const p = getComputedStyle(el, which).content;
      if (p && p !== 'none' && p !== 'normal') return true;
    }
    return false;
  }

  function radii(cs) {
    return ['borderTopLeftRadius', 'borderTopRightRadius', 'borderBottomRightRadius', 'borderBottomLeftRadius'].map((p) => num(cs[p]));
  }

  function backgroundUrl(cs) {
    const bi = cs.backgroundImage;
    if (!bi || bi === 'none') return null;
    const m = /^url\("?(.*?)"?\)$/.exec(bi.trim());
    return m ? m[1] : null;
  }

  function percent(value, fallback) {
    const m = /^(-?[\d.]+)%$/.exec(String(value || '').trim());
    return m ? Number(m[1]) / 100 : fallback;
  }

  async function naturalSize(src) {
    const img = new Image();
    img.src = src;
    try { await img.decode(); } catch (e) { return null; }
    return img.naturalWidth && img.naturalHeight ? { w: img.naturalWidth, h: img.naturalHeight } : null;
  }

  // Where a picture of natural size n lands in box b for a CSS fit mode.
  function fitPicture(n, b, fit, px, py) {
    if (fit === 'fill') return { ...b, crop: { l: 0, t: 0, r: 0, b: 0 } };
    const scale = fit === 'cover' ? Math.max(b.w / n.w, b.h / n.h) : Math.min(b.w / n.w, b.h / n.h);
    const rw = n.w * scale, rh = n.h * scale;
    if (fit === 'cover') {
      const ex = Math.max(0, (rw - b.w) / rw), ey = Math.max(0, (rh - b.h) / rh);
      return { ...b, crop: { l: ex * px, r: ex * (1 - px), t: ey * py, b: ey * (1 - py) } };
    }
    return { x: b.x + (b.w - rw) * px, y: b.y + (b.h - rh) * py, w: rw, h: rh, crop: { l: 0, t: 0, r: 0, b: 0 } };
  }

  async function pictureFor(el, cs, rect, usable) {
    if (el.tagName === 'IMG') {
      const src = el.currentSrc || el.src;
      if (!src || !usable.has(src) || !el.complete || !el.naturalWidth) return null;
      if (radii(cs).some((r) => r > 0)) return null;
      const fit = cs.objectFit || 'fill';
      if (!/^(fill|contain|cover)$/.test(fit)) return null;
      const pos = (cs.objectPosition || '50% 50%').split(/\s+/);
      const content = {
        x: rect.x + num(cs.borderLeftWidth) + num(cs.paddingLeft), y: rect.y + num(cs.borderTopWidth) + num(cs.paddingTop),
        w: rect.w - num(cs.borderLeftWidth) - num(cs.borderRightWidth) - num(cs.paddingLeft) - num(cs.paddingRight),
        h: rect.h - num(cs.borderTopWidth) - num(cs.borderBottomWidth) - num(cs.paddingTop) - num(cs.paddingBottom),
      };
      const placed = fitPicture({ w: el.naturalWidth, h: el.naturalHeight }, content, fit, percent(pos[0], 0.5), percent(pos[1], 0.5));
      return { src, alt: el.getAttribute('alt') || '', ...placed };
    }
    const src = backgroundUrl(cs);
    if (!src || !usable.has(src)) return null;
    if (radii(cs).some((r) => r > 0)) return null;
    if (cs.backgroundImage.includes(',')) return null;
    const size = cs.backgroundSize.trim();
    const fit = size === 'cover' ? 'cover' : size === 'contain' ? 'contain' : size === '100% 100%' ? 'fill' : null;
    if (!fit) return null;
    if (fit !== 'cover' && !/no-repeat/.test(cs.backgroundRepeat)) return null;
    if (fit !== 'cover' && color(cs.backgroundColor)) return null;
    const n = await naturalSize(src);
    if (!n) return null;
    const pos = cs.backgroundPosition.split(',')[0].trim().split(/\s+/);
    const placed = fitPicture(n, rect, fit, percent(pos[0], 0.5), percent(pos[1], 0.5));
    return { src, alt: el.getAttribute('aria-label') || '', ...placed };
  }

  // The painted parts of a plain box: its fill and its border sides.
  function shapeParts(el, cs, rect, clip) {
    if (REPLACED_TAGS.has(el.tagName)) return null;
    if (cs.backgroundImage && cs.backgroundImage !== 'none') return null;
    if (cs.boxShadow && cs.boxShadow !== 'none') return null;
    if (cs.outlineStyle !== 'none' && num(cs.outlineWidth) > 0) return null;
    const fill = color(cs.backgroundColor);
    const sides = ['Top', 'Right', 'Bottom', 'Left'].filter((s) => hasBorderPaint(cs, s));
    if (!fill && !sides.length) return null;
    if (sides.some((s) => !/^solid$/.test(cs['border' + s + 'Style']))) return null;
    const r = radii(cs);
    const rounded = r.some((v) => v > 0.5);
    const uniformBorder = sides.length === 4 && sides.every((s) => num(cs['border' + s + 'Width']) === num(cs.borderTopWidth) && cs['border' + s + 'Color'] === cs.borderTopColor);
    const clipped = !clip || clip.x > rect.x + 0.5 || clip.y > rect.y + 0.5 || clip.x + clip.w < rect.x + rect.w - 0.5 || clip.y + clip.h < rect.y + rect.h - 0.5;
    const parts = [];
    if (rounded) {
      // Rounded boxes become one rounded shape; only uniform borders keep their outline.
      if (clipped || (sides.length && !uniformBorder)) return null;
      if (r.some((v) => Math.abs(v - r[0]) > 0.5)) return null;
      const minSide = Math.min(rect.w, rect.h);
      const radius = Math.min(r[0], minSide / 2);
      const circle = radius >= minSide / 2 - 0.5 && Math.abs(rect.w - rect.h) < 1;
      parts.push({
        kind: circle ? 'ellipse' : 'roundRect', ...rect, radius,
        fill, line: sides.length ? { ...color(cs.borderTopColor), width: num(cs.borderTopWidth) } : null,
      });
      return parts;
    }
    const area = clip ? intersect(rect, clip) : null;
    if (!area) return null;
    if (fill) parts.push({ kind: 'rect', ...area, fill, line: null });
    const bw = (s) => num(cs['border' + s + 'Width']);
    for (const s of sides) {
      const c = color(cs['border' + s + 'Color']);
      let side;
      if (s === 'Top') side = { x: rect.x, y: rect.y, w: rect.w, h: bw('Top') };
      else if (s === 'Bottom') side = { x: rect.x, y: rect.y + rect.h - bw('Bottom'), w: rect.w, h: bw('Bottom') };
      else if (s === 'Left') side = { x: rect.x, y: rect.y, w: bw('Left'), h: rect.h };
      else side = { x: rect.x + rect.w - bw('Right'), y: rect.y, w: bw('Right'), h: rect.h };
      const visiblePart = intersect(side, clip);
      if (visiblePart) parts.push({ kind: 'rect', ...visiblePart, fill: c, line: null });
    }
    return parts.length ? parts : null;
  }

  function samplePoints(part, sRect, clip) {
    const a = intersect(part, clip || { x: 0, y: 0, w: sRect.width, h: sRect.height });
    if (!a) return [];
    const nx = Math.min(5, Math.max(1, Math.round(a.w / 40))), ny = Math.min(5, Math.max(1, Math.round(a.h / 40)));
    const pts = [];
    for (let i = 0; i < nx; i++) for (let j = 0; j < ny; j++) {
      pts.push([sRect.left + a.x + a.w * (i + 0.5) / nx, sRect.top + a.y + a.h * (j + 0.5) / ny]);
    }
    return pts;
  }

  /* ---------- public steps ---------- */

  function prepare() {
    const style = document.createElement('style');
    style.id = 'pptx-export-style';
    style.textContent = [
      '*,*::before,*::after{animation:none!important;transition:none!important;pointer-events:auto!important;caret-color:transparent!important}',
      '[data-pptx-nobefore]::before{content:none!important}',
      '[data-pptx-noafter]::after{content:none!important}',
      'html,body{overflow:hidden!important}',
    ].join('\n');
    document.head.appendChild(style);
    slides = Array.from(document.querySelectorAll('.slide')).filter((s) => !(s.parentElement && s.parentElement.closest('.slide')));
    return slides.map((s) => ({ w: s.offsetWidth, h: s.offsetHeight }));
  }

  // Shows one slide alone, unscaled, at the top-left of the viewport.
  function focus(index) {
    slides.forEach((s, j) => { if (j !== index) setImportant(s, 'visibility', 'hidden'); });
    const s = slides[index];
    const w = s.offsetWidth, h = s.offsetHeight;
    for (let a = s.parentElement; a && a !== document.documentElement; a = a.parentElement) {
      for (const [p, v] of [['transform', 'none'], ['filter', 'none'], ['perspective', 'none'], ['contain', 'none'], ['will-change', 'auto'], ['visibility', 'hidden']]) setImportant(a, p, v);
    }
    for (const [p, v] of [['position', 'fixed'], ['left', '0px'], ['top', '0px'], ['right', 'auto'], ['bottom', 'auto'], ['margin', '0px'],
      ['transform', 'none'], ['width', w + 'px'], ['height', h + 'px'], ['z-index', '2147483647'], ['visibility', 'visible']]) setImportant(s, p, v);
    return { w, h };
  }

  function unfocus(index) {
    const s = slides[index];
    setImportant(s, 'visibility', 'hidden');
  }

  async function extract(index, usableList) {
    const usable = new Set(usableList || []);
    const slide = slides[index];
    const pseudos = materializePseudos(slide);
    const sRect = slide.getBoundingClientRect();
    const opacityCache = new Map();
    const opacityOf = (el) => { if (!opacityCache.has(el)) opacityCache.set(el, effectiveOpacity(el, slide)); return opacityCache.get(el); };
    const all = Array.from(slide.querySelectorAll('*'));
    const order = new Map(all.map((el, i) => [el, i]));

    // Text boxes.
    const containers = new Set();
    const tw = document.createTreeWalker(slide, NodeFilter.SHOW_TEXT);
    for (let t = tw.nextNode(); t; t = tw.nextNode()) {
      if (!t.nodeValue.trim()) continue;
      const parent = t.parentElement;
      if (!parent || inSvg(parent) || parent.closest('script,style,noscript,template')) continue;
      const c = textContainer(t, slide);
      if (c) containers.add(c);
    }
    const texts = [];
    for (const c of containers) {
      const cs = getComputedStyle(c);
      if (!visible(c, cs) || transformedOrFiltered(c, slide)) continue;
      const tb = textBox(c, slide, sRect, opacityOf);
      if (tb) texts.push(tb);
    }

    // Candidates for native shapes and pictures.
    const candidates = new Map();
    for (const el of all) {
      if (inSvg(el) && el.tagName.toLowerCase() !== 'svg') continue;
      const cs = getComputedStyle(el);
      if (!visible(el, cs) || isInlineDisplay(cs) || el.hasAttribute('data-pptx-pseudo') && !paints(el, cs)) continue;
      if (!paints(el, cs) || transformedOrFiltered(el, slide)) continue;
      const opacity = opacityOf(el);
      const rect = relRect(el.getBoundingClientRect(), sRect);
      const clip = clipRect(el, slide, sRect);
      if (!clip || !intersect(rect, clip)) continue;
      // Paint from the element's own pseudo-elements would end up under the shape.
      const ownPseudo = ['::before', '::after'].some((w) => { const p = getComputedStyle(el, w).content; return p && p !== 'none' && p !== 'normal'; });
      if (ownPseudo) continue;
      const pic = (el.tagName === 'IMG' || backgroundUrl(cs)) && opacity > 0.995 ? await pictureFor(el, cs, rect, usable) : null;
      if (pic) {
        const clippedPic = intersect(pic, clip);
        if (!clippedPic || Math.abs(clippedPic.w - pic.w) > 0.5 || Math.abs(clippedPic.h - pic.h) > 0.5) {
          // Clip-aware cropping is only safe for cover-fitted or filled pictures inside a plain clip.
          continue;
        }
        candidates.set(el, { el, type: 'picture', parts: [rect], obj: { kind: 'picture', ...pic } });
        continue;
      }
      if (el.tagName === 'IMG' || backgroundUrl(cs)) continue;
      const parts = shapeParts(el, cs, rect, clip);
      if (!parts) continue;
      for (const p of parts) {
        if (p.fill) p.fill = { ...p.fill, alpha: Math.round(p.fill.alpha * opacity * 1000) / 1000 };
        if (p.line) p.line = { ...p.line, alpha: Math.round(p.line.alpha * opacity * 1000) / 1000 };
      }
      candidates.set(el, { el, type: 'shape', parts, obj: null });
    }

    // Hit-test each candidate. It can only leave the background picture if
    // everything drawn above it is itself rebuilt (or draws nothing).
    const stacks = new Map();
    for (const cand of candidates.values()) {
      const clip = clipRect(cand.el, slide, sRect);
      const pts = [];
      for (const part of cand.parts) for (const p of samplePoints(part, sRect, clip)) pts.push(p);
      stacks.set(cand.el, pts.map(([x, y]) => document.elementsFromPoint(x, y)));
    }
    const paintsCache = new Map();
    const paintsEl = (el) => { if (!paintsCache.has(el)) paintsCache.set(el, paints(el, getComputedStyle(el))); return paintsCache.get(el); };
    let changed = true;
    while (changed) {
      changed = false;
      for (const el of Array.from(candidates.keys())) {
        const ok = stacks.get(el).length > 0 && stacks.get(el).every((stack) => {
          const i = stack.indexOf(el);
          if (i < 0) return false;
          for (let k = 0; k < i; k++) {
            const above = stack[k];
            if (!slide.contains(above)) continue;
            if (candidates.has(above)) continue;
            if (paintsEl(above)) return false;
          }
          return true;
        });
        if (!ok) { candidates.delete(el); changed = true; }
      }
    }

    // Paint order of what was lifted: what the hit tests saw on top goes later.
    const lifted = Array.from(candidates.values());
    const above = new Map(lifted.map((c) => [c.el, new Set()]));
    for (const c of lifted) {
      for (const stack of stacks.get(c.el)) {
        const i = stack.indexOf(c.el);
        for (let k = 0; k < i; k++) if (candidates.has(stack[k])) above.get(c.el).add(stack[k]);
      }
    }
    const sorted = [];
    const state = new Map();
    const visit = (c) => {
      if (state.get(c.el) === 2) return;
      if (state.get(c.el) === 1) return; // cycle: keep document order
      state.set(c.el, 1);
      // Everything below c must come first: those are the elements that have c above them.
      for (const other of lifted) if (other !== c && above.get(other.el).has(c.el)) visit(other);
      state.set(c.el, 2);
      sorted.push(c);
    };
    lifted.sort((a, b) => order.get(a.el) - order.get(b.el)).forEach(visit);

    const objects = [];
    for (const c of sorted) {
      if (c.type === 'picture') objects.push(c.obj);
      else for (const p of c.parts) objects.push(p);
    }

    // Remove what was rebuilt from the page, so the background picture is clean.
    const frozen = new Set();
    const freeze = (el) => {
      if (frozen.has(el)) return;
      frozen.add(el);
      const cs = getComputedStyle(el);
      for (const p of ['color', 'border-top-color', 'border-right-color', 'border-bottom-color', 'border-left-color', 'outline-color', 'background-color', 'box-shadow', 'text-decoration-color', 'column-rule-color']) {
        setImportant(el, p, cs.getPropertyValue(p));
      }
    };
    all.forEach(freeze);
    for (const t of texts) {
      for (const el of [t.node.container, ...t.node.inlineEls]) {
        setImportant(el, 'color', 'transparent');
        setImportant(el, '-webkit-text-fill-color', 'transparent');
        setImportant(el, 'text-shadow', 'none');
        setImportant(el, 'text-decoration-color', 'transparent');
      }
    }
    for (const c of lifted) {
      if (c.type === 'picture') {
        if (c.el.tagName === 'IMG') setImportant(c.el, 'opacity', '0');
        else { setImportant(c.el, 'background-image', 'none'); setImportant(c.el, 'background-color', 'transparent'); }
      } else {
        setImportant(c.el, 'background-color', 'transparent');
        for (const s of ['top', 'right', 'bottom', 'left']) setImportant(c.el, 'border-' + s + '-color', 'transparent');
      }
    }

    return {
      w: sRect.width, h: sRect.height,
      objects,
      texts: texts.map((t) => t.box),
      pseudos,
      lang: slide.getAttribute('lang') || document.documentElement.lang || '',
      rtl: getComputedStyle(slide).direction === 'rtl',
    };
  }

  function sources() {
    // Every picture source a slide uses, so Python can say which ones it can embed.
    const out = new Set();
    for (const s of slides) {
      for (const el of [s, ...s.querySelectorAll('*')]) {
        if (el.tagName === 'IMG' && (el.currentSrc || el.src)) out.add(el.currentSrc || el.src);
        const u = backgroundUrl(getComputedStyle(el));
        if (u) out.add(u);
      }
    }
    return Array.from(out);
  }

  window.__pptx = { prepare, focus, unfocus, extract, sources };
})();
