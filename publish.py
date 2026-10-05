#!/usr/bin/env python3
"""Выкладка готовой презентации для пользователя (Open WebUI + Open Terminal).

    python3 publish.py "<собранный>.pptx" [--root ~/витомэк-презентации]

- копирует .pptx в «<root>/готовые презентации/» (там лежат только презентации);
- кладёт картинку-обзор (…_preview/contact.jpg) в «<root>/служебное/превью/»;
- пересобирает страницу «<root>/служебное/скачать.html»: свежая презентация с превью и кнопкой
  скачивания, ниже — все остальные готовые презентации;
- печатает путь к странице: её нужно показать в чате через display_file(path, inline=true).

Ссылки на странице относительные — просмотрщик Open WebUI отдаёт HTML по пути файла,
поэтому они ведут прямо к файлам в терминале.
"""
import argparse
import html
import shutil
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

READY = 'готовые презентации'
SERVICE = 'служебное'
PREVIEWS = 'превью'
PAGE = 'скачать.html'


def size_label(p):
    mb = p.stat().st_size / 1024 / 1024
    return f'{mb:.1f} МБ'.replace('.', ',') if mb >= 0.1 else f'{p.stat().st_size // 1024} КБ'


def slide_count(p):
    try:
        from pptx import Presentation
        return len(Presentation(str(p)).slides)
    except Exception:
        return None


def href(*parts):
    return '/'.join(quote(x) for x in parts)


def page(root, current):
    ready = sorted((root / READY).glob('*.pptx'), key=lambda p: p.stat().st_mtime, reverse=True)
    others = [p for p in ready if p.name != current.name]

    def meta(p):
        n = slide_count(p)
        bits = ([f'{n} слайд' + ('ов' if n % 10 in (0, 5, 6, 7, 8, 9) or 11 <= n % 100 <= 14
                                  else 'а' if n % 10 in (2, 3, 4) else '')] if n else [])
        bits += [size_label(p), datetime.fromtimestamp(p.stat().st_mtime).strftime('%d.%m.%Y %H:%M')]
        return ' · '.join(bits)

    prev = root / SERVICE / PREVIEWS / (current.stem + '.jpg')
    img = (f'<a href="{href(PREVIEWS, prev.name)}" target="_blank"><img src="{href(PREVIEWS, prev.name)}" '
           f'alt="Обзор слайдов"></a>') if prev.exists() else ''
    rows = ''.join(
        f'<li><a href="{href("..", READY, p.name)}" download="{html.escape(p.name)}">{html.escape(p.stem)}</a>'
        f'<span>{meta(p)}</span></li>' for p in others)
    older = f'<h2>Другие готовые презентации</h2><ul>{rows}</ul>' if others else ''
    return f'''<!DOCTYPE html>
<html lang="ru"><head><meta charset="utf-8"><title>Готовые презентации VITOMEK</title>
<style>
 body{{font-family:Arial,Helvetica,sans-serif;margin:0;padding:20px;color:#0C4C4C;background:#fff}}
 .card{{border:1px solid #e3e3e3;border-radius:12px;padding:18px;max-width:760px}}
 .tag{{font-size:12px;color:#6f9a1e;text-transform:uppercase;letter-spacing:.06em}}
 h1{{font-size:20px;margin:6px 0 4px}} .meta{{color:#7a7a7a;font-size:13px;margin-bottom:14px}}
 .btn{{display:inline-block;background:#0C4C4C;color:#fff;text-decoration:none;font-weight:bold;
   padding:12px 22px;border-radius:8px;font-size:15px}} .btn:hover{{background:#4E7C7C}}
 img{{display:block;width:100%;border:1px solid #eee;border-radius:8px;margin-top:16px}}
 .where{{font-size:12px;color:#7a7a7a;margin-top:12px}}
 h2{{font-size:15px;margin:24px 0 8px}} ul{{list-style:none;padding:0;margin:0;max-width:760px}}
 li{{display:flex;justify-content:space-between;gap:12px;padding:9px 0;border-bottom:1px solid #eee;font-size:14px}}
 li a{{color:#0C4C4C}} li span{{color:#7a7a7a;font-size:12px;white-space:nowrap}}
</style></head><body>
<div class="card">
 <div class="tag">Готовая презентация</div>
 <h1>{html.escape(current.stem)}</h1>
 <div class="meta">{meta(current)}</div>
 <a class="btn" href="{href("..", READY, current.name)}" download="{html.escape(current.name)}">⬇ Скачать .pptx</a>
 {img}
 <div class="where">Файл лежит в папке терминала: витомэк-презентации / {READY}</div>
</div>
{older}
</body></html>
'''


def main():
    ap = argparse.ArgumentParser(description='Выложить готовую презентацию и обновить страницу скачивания')
    ap.add_argument('pptx', help='собранный .pptx')
    ap.add_argument('--root', default=str(Path.home() / 'витомэк-презентации'))
    a = ap.parse_args()

    src = Path(a.pptx).expanduser().resolve()
    root = Path(a.root).expanduser()
    (root / READY).mkdir(parents=True, exist_ok=True)
    (root / SERVICE / PREVIEWS).mkdir(parents=True, exist_ok=True)

    dst = root / READY / src.name
    if src != dst.resolve():
        shutil.copy2(src, dst)
    contact = src.parent / (src.stem + '_preview') / 'contact.jpg'
    if contact.exists():
        shutil.copy2(contact, root / SERVICE / PREVIEWS / (src.stem + '.jpg'))

    out = root / SERVICE / PAGE
    out.write_text(page(root, dst), encoding='utf-8')
    print(f'Готово: {dst}')
    print(f'Страница скачивания: {out}')
    print(f'Покажи её в чате: display_file(path="{out}", inline=true)')


if __name__ == '__main__':
    main()
