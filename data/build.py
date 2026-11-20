#!/usr/bin/env python3
"""Превращает текст НПА (txt / docx / html) в data/<id>.json для сайта.

Пример:
  python3 build.py koap.docx ../npa-site/data/koap.json --edition "ред. от 04.08.2026"
"""
import argparse, html, json, re, sys
from pathlib import Path

def read_lines(path: Path):
    ext = path.suffix.lower()
    if ext == '.docx':
        # идём по документу по порядку: абзацы + таблицы-примечания КонсультантПлюс
        import docx
        from docx.table import Table
        from docx.text.paragraph import Paragraph
        d = docx.Document(str(path)); out = []
        for el in d.element.body.iterchildren():
            tag = el.tag.split('}')[1]
            if tag == 'p':
                out.extend(Paragraph(el, d).text.split('\n'))
            elif tag == 'tbl':
                t = Table(el, d)
                txt = ' '.join(c.text for r in t.rows for c in r.cells).strip()
                if txt.startswith('КонсультантПлюс: примечание'):
                    out.append(re.sub(r'\s+', ' ', txt.replace('примечание.', 'примечание.', 1)))
        return out
    raw = path.read_bytes()
    for enc in ('utf-8-sig', 'cp1251'):
        try:
            text = raw.decode(enc); break
        except UnicodeDecodeError:
            continue
    if ext in ('.htm', '.html'):
        text = re.sub(r'(?is)<(script|style).*?</\1>', ' ', text)
        text = re.sub(r'(?i)<br\s*/?>|</(p|div|h\d|li|tr)>', '\n', text)
        text = html.unescape(re.sub(r'<[^>]+>', '', text))
    return text.splitlines()

RE_PART    = re.compile(r'^(ОБЩАЯ|ОСОБЕННАЯ)\s+ЧАСТЬ\s*$', re.I)
RE_SECTION = re.compile(r'^Раздел\s+([IVXLC]+(?:\.\d+)?)\.?\s*(.*)$', re.I)
RE_SECTION_W = re.compile(r'^РАЗДЕЛ\s+(ПЕРВЫЙ|ВТОРОЙ|ТРЕТИЙ|ЧЕТВЕРТЫЙ|ПЯТЫЙ|ШЕСТОЙ)$')
RE_DASH = re.compile(r'^-{10,}$')
RE_CHAPTER = re.compile(r'^Глава\s+(\d+(?:\.\d+)?)\.?\s*(.*)$', re.I)
RE_ARTICLE = re.compile(r'^Статья\s+(\d+(?:\.\d+)*(?:-\d+)?)\.?\s*(.*)$')
NOISE = re.compile(r'^(Документ предоставлен КонсультантПлюс|www\.consultant\.ru|Дата сохранения|Страница \d+ из \d+|Список изменяющих документов|\(см\. Обзор изменений)', re.I)

def clean(s):
    s = s.replace(' ', ' ').replace('\t', ' ')
    return re.sub(r'\s{2,}', ' ', s).strip()

def parse(lines, preamble=None):
    items, cur, pending = [], None, None   # pending — заголовок, продолжающийся на след. строке
    seen = set(); hold = []; skip_paren = False     # примечания Консультанта между заголовком и статьёй
    for raw in lines:
        s = clean(raw)
        if not s or NOISE.match(s) or RE_DASH.match(s):
            continue
        if preamble and not items and s.startswith(preamble):
            cur = {'type': 'article', 'num': 'Преамбула', 'head': 'Преамбула', 'title': '', 'text': []}
            items.append(cur); seen.add(cur['num'])
        m = RE_SECTION_W.match(s)
        if m:
            cur = None
            it = {'type': 'section', 'label': 'Раздел ' + m.group(1).lower(), 'title': 'Раздел ' + m.group(1).lower()}
            items.append(it); pending = it; continue
        if s == 'Президент' and cur is not None:   # подпись в конце закона
            break
        m = RE_PART.match(s)
        if m:
            cur = None; items.append({'type': 'part', 'title': s.capitalize()}); continue
        m = RE_SECTION.match(s)
        if m and (not m.group(2) or m.group(2).isupper() or len(s) < 120):
            cur = None
            it = {'type': 'section', 'label': 'Раздел ' + m.group(1), 'title': (m.group(2) or '').strip() or 'Раздел ' + m.group(1)}
            items.append(it); pending = it if not m.group(2).strip() else None; continue
        m = RE_CHAPTER.match(s)
        if m and len(s) < 220:
            cur = None
            it = {'type': 'chapter', 'label': 'Глава ' + m.group(1), 'title': m.group(2).strip() or 'Глава ' + m.group(1)}
            items.append(it); pending = it if not m.group(2) else None; continue
        m = RE_ARTICLE.match(s)
        if m and len(s) < 400:
            num = m.group(1)
            if num in seen:          # дубль (например, ссылка в тексте) — считаем обычным абзацем
                if cur: cur['text'].append(s)
                continue
            seen.add(num)
            cur = {'type': 'article', 'num': num, 'title': m.group(2).strip(), 'text': hold}; hold = []
            items.append(cur); pending = None; continue
        if pending is not None and pending['title'].startswith(('Раздел', 'Глава')):
            pending['title'] = s; continue
        # продолжение заголовка главы/раздела, перенесённое на новую строку (всё заглавными)
        if cur is None and items and items[-1]['type'] in ('chapter', 'section') \
                and not re.search(r'[а-яё]', s) and re.search(r'[А-ЯЁ]', s):
            items[-1]['title'] = (items[-1]['title'] + ' ' + s).strip(); continue
        pending = None
        if skip_paren:                   # продолжение многострочной «(в ред. ...)»
            skip_paren = not s.endswith(')'); continue
        if cur is None and s.startswith('(') and items and items[-1]['type'] in ('chapter', 'section'):
            skip_paren = not s.endswith(')'); continue   # «(в ред. ...)» к заголовку главы
        if cur is None and s.startswith('КонсультантПлюс: примечание'):
            hold.append(s); continue
        if cur is None and re.match(r'^<(\d+|\*)>', s):   # сноска к заголовку главы — пропускаем
            continue
        if cur is None and items and items[-1]['type'] in ('chapter', 'section'):
            h = items[-1]
            num = ''.join(w[0] for w in re.findall(r'[А-ЯЁа-яё]+', h['title']) if len(w) > 1).upper()[:5] or 'Т'
            cur = {'type': 'article', 'num': num, 'head': h['title'].capitalize(), 'title': '', 'text': []}
            items.append(cur); seen.add(num)
        if cur is not None:
            cur['text'].append(s)
    for it in items:
        if it['type'] != 'article':
            it['title'] = re.sub(r'\s*<(\d+|\*)>', '', it['title']).strip()
        if it['type'] == 'article' and not it['title'] and not it.get('head') and it['text']:
            first = re.sub(r'^\d+\.\s*', '', it['text'][0])
            it['hint'] = first[:90].rsplit(' ', 1)[0] + '…' if len(first) > 90 else first
        if it['type'] == 'article':
            t = it['text']
            if not it['title'] and t and re.match(r'^Утратил[аио]? силу', t[0]):
                it['title'] = t[0]
            if re.match(r'^Утратил[аио]? силу', it['title'] or (t[0] if t else '')):
                it['lost'] = True
                it['title'] = re.split(r'\s+[-–—]\s+', it['title'])[0].rstrip('. ') or 'Утратила силу'
    return items

RE_ROMAN = re.compile(r'^([IVXL]+)\.\s+(.+)$')

def nice(t):
    """ЗАГЛАВНЫЕ → обычный регистр с сохранением имён собственных."""
    t = t.strip().rstrip('.')
    t = t[0] + t[1:].lower()
    for a, b in (('российской федерации', 'Российской Федерации'), ('мвд россии', 'МВД России')):
        t = t.replace(a, b)
    return t

def parse_order(lines):
    """Приказ: «Текст приказа» + приложения, внутри — разделы с римскими номерами (пункты сквозные)."""
    items, cur, mode = [], None, 'head'   # head → order → sign → app_hdr → body
    app_no, title_buf, prefix = 0, [], ''
    for raw in lines:
        s = clean(raw)
        if not s or NOISE.match(s) or RE_DASH.match(s):
            continue
        if mode == 'head':
            if s.startswith('В соответствии') or re.match(r'^1\.\s', s):
                cur = {'type': 'article', 'num': 'Приказ', 'head': 'Текст приказа', 'title': '', 'text': [s]}
                items.append(cur); mode = 'order'
            continue
        if re.match(r'^Приложение( N \d+)?$', s):
            cur = None; mode = 'app_hdr'; title_buf = []; app_no += 1
            prefix = '' if app_no == 1 else f'п{app_no}-'
            continue
        if mode == 'order':
            if s == 'Министр':
                mode = 'sign'; cur = None; continue
            cur['text'].append(s); continue
        if mode == 'sign':
            continue
        if mode == 'app_hdr':
            if re.search(r'[а-яё]', s) and not title_buf:
                continue                                  # «к приказу МВД России», «от ...»
            if not re.search(r'[а-яё]', s):
                title_buf.append(s); continue
            items.append({'type': 'part', 'title': nice(' '.join(title_buf))})
            mode = 'body'                                 # и обрабатываем строку дальше
        m = RE_ROMAN.match(s)
        if m and len(s) < 1200 and not s.endswith(';'):
            cur = {'type': 'article', 'num': prefix + m.group(1), 'head': s, 'title': '', 'text': []}
            items.append(cur); continue
        if cur is None:                                   # приложение без разделов
            cur = {'type': 'article', 'num': f'прил{app_no}', 'head': items[-1]['title'], 'title': '', 'text': []}
            items.append(cur)
        if re.match(r'^(Глава|Раздел) утратил[аи]? силу', s) and not cur['text']:
            cur['lost'] = True
        cur['text'].append(s)
    return items

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('src'); ap.add_argument('out')
    ap.add_argument('--title', default=''); ap.add_argument('--edition', default='')
    ap.add_argument('--preamble', default='', help='начало текста преамбулы (для Конституции)')
    ap.add_argument('--order', action='store_true', help='приказ: разделы I, II… и сквозные пункты')
    a = ap.parse_args()
    lines = read_lines(Path(a.src))
    items = parse_order(lines) if a.order else parse(lines, a.preamble or None)
    arts = [i for i in items if i['type'] == 'article']
    if not arts:
        sys.exit('Статьи не найдены — проверь исходный файл')
    out = Path(a.out); dump = lambda o: json.dumps(o, ensure_ascii=False, separators=(',', ':'))
    for old in out.parent.glob(out.stem + '-*.json'): old.unlink()
    # большие НПА режем на части (~2 МБ): главный файл хранит список частей
    chunks, buf, size = [], [], 0
    for it in items:
        buf.append(it); size += len(dump(it).encode())
        if size > 2_000_000: chunks.append(buf); buf, size = [], 0
    if buf: chunks.append(buf)
    if len(chunks) == 1:
        out.write_text(dump({'title': a.title, 'edition': a.edition, 'items': items}), encoding='utf-8')
    else:
        names = []
        for k, ch in enumerate(chunks[1:], 2):
            n = f'{out.stem}-{k}.json'; names.append(n)
            (out.parent / n).write_text(dump(ch), encoding='utf-8')
        out.write_text(dump({'title': a.title, 'edition': a.edition, 'items': chunks[0], 'parts': names}), encoding='utf-8')
        print('частей:', len(chunks))
    print(f"{a.out}: статей {len(arts)}, глав {sum(i['type']=='chapter' for i in items)}, "
          f"разделов {sum(i['type']=='section' for i in items)}, {Path(a.out).stat().st_size//1024} КБ")

if __name__ == '__main__':
    main()
