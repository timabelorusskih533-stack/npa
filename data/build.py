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
RE_CHAPTER = re.compile(r'^Глава\s+(\d+(?:\.\d+)?)\.?\s*(.*)$', re.I)
RE_ARTICLE = re.compile(r'^Статья\s+(\d+(?:\.\d+)*(?:-\d+)?)\.?\s*(.*)$')
NOISE = re.compile(r'^(Документ предоставлен КонсультантПлюс|www\.consultant\.ru|Дата сохранения|Страница \d+ из \d+|Список изменяющих документов|\(см\. Обзор изменений)', re.I)

def clean(s):
    s = s.replace(' ', ' ').replace('\t', ' ')
    return re.sub(r'\s{2,}', ' ', s).strip()

def parse(lines):
    items, cur, pending = [], None, None   # pending — заголовок, продолжающийся на след. строке
    seen = set()
    for raw in lines:
        s = clean(raw)
        if not s or NOISE.match(s):
            continue
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
            cur = {'type': 'article', 'num': num, 'title': m.group(2).strip(), 'text': []}
            items.append(cur); pending = None; continue
        if pending is not None and pending['title'].startswith(('Раздел', 'Глава')):
            pending['title'] = s; continue
        # продолжение заголовка главы/раздела, перенесённое на новую строку (всё заглавными)
        if cur is None and items and items[-1]['type'] in ('chapter', 'section') \
                and not re.search(r'[а-яё]', s) and re.search(r'[А-ЯЁ]', s):
            items[-1]['title'] = (items[-1]['title'] + ' ' + s).strip(); continue
        pending = None
        if cur is not None:
            cur['text'].append(s)
    for it in items:
        if it['type'] == 'article':
            t = it['text']
            if not it['title'] and t and re.match(r'^Утратил[аио]? силу', t[0]):
                it['title'] = t[0]
            if re.match(r'^Утратил[аио]? силу', it['title'] or (t[0] if t else '')):
                it['lost'] = True
                it['title'] = re.split(r'\s+[-–—]\s+', it['title'])[0].rstrip('. ') or 'Утратила силу'
    return items

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('src'); ap.add_argument('out')
    ap.add_argument('--title', default=''); ap.add_argument('--edition', default='')
    a = ap.parse_args()
    items = parse(read_lines(Path(a.src)))
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
