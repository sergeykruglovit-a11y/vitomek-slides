#!/usr/bin/env python3
"""Выкладка готовой презентации для пользователя (Open WebUI + Open Terminal).

    python3 publish.py "<собранный>.pptx" [--root ~/витомэк-презентации] [--link]

- копирует .pptx в «<root>/готовые презентации/» (там лежат только презентации);
- кладёт картинку-обзор (…_preview/contact.jpg) в «<root>/служебное/превью/»;
- пересобирает страницу «<root>/служебное/скачать.html»: свежая презентация с превью и кнопкой
  скачивания, ниже — все остальные готовые презентации;
- печатает путь к странице: её нужно показать в чате через display_file(path, inline=true).

По умолчанию файл и превью встроены в страницу (base64): кнопка скачивает без обращения к серверу,
поэтому работает в любом браузере, даже если он не передаёт cookie во встроенный просмотрщик чата.
С --link страница лёгкая: относительные ссылки на файлы терминала.
В обоих режимах на странице есть подсказка, как скачать через панель «Файлы», а в выводе — блок
«ТЕКСТ ДЛЯ ОТВЕТА», который агент копирует в начало ответа.
"""
import argparse
import base64
import html
import shutil
from urllib.parse import quote
from datetime import datetime
from pathlib import Path

READY = 'готовые презентации'
SERVICE = 'служебное'
PREVIEWS = 'превью'
PAGE = 'скачать.html'
PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'


def size_label(p):
    mb = p.stat().st_size / 1024 / 1024
    return f'{mb:.1f} МБ'.replace('.', ',') if mb >= 0.1 else f'{p.stat().st_size // 1024} КБ'


def slide_count(p):
    try:
        from pptx import Presentation
        return len(Presentation(str(p)).slides)
    except Exception:
        return None


def b64(p):
    return base64.b64encode(p.read_bytes()).decode()


def href(*parts):
    return '/'.join(quote(x) for x in parts)


def hint_html(name):
    return f'''<div class="hint"><b>Если вместо скачивания появилась ошибка</b> (некоторые браузеры так защищают вход):
 <ol><li>В правом верхнем углу чата нажмите «Управление» и откройте вкладку «Файлы».</li>
 <li>Перейдите в папку <b>витомэк-презентации → {READY}</b>.</li>
 <li>У файла «{html.escape(name)}» нажмите «⋯» → «Загрузить» — файл скачается на компьютер.</li></ol></div>'''


def answer_text(name):
    """Готовый текст для начала ответа агента — модель копирует его, а не пишет сама."""
    return ('Презентация готова — кнопка «Скачать .pptx» выше.\n'
            'Если браузер вместо скачивания показал ошибку: справа вверху нажмите «Управление» → вкладка «Файлы» → '
            f'витомэк-презентации → {READY} → у файла «{name}» нажмите «⋯» → «Загрузить».')


SCRIPT = '''<script>
// data:-ссылка большого размера может не скачаться в некоторых браузерах — отдаём через Blob
document.getElementById('dl').addEventListener('click', function (e) {
  try {
    var a = this, b = atob(a.href.split(',')[1]), u = new Uint8Array(b.length);
    for (var i = 0; i < b.length; i++) u[i] = b.charCodeAt(i);
    var url = URL.createObjectURL(new Blob([u], {type: '%s'}));
    var t = document.createElement('a'); t.href = url; t.download = a.getAttribute('download');
    document.body.appendChild(t); t.click(); t.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 10000);
    e.preventDefault();
  } catch (err) { /* остаётся обычная data:-ссылка */ }
});
</script>''' % PPTX_MIME


def page(root, current, embed=True):
    ready = sorted((root / READY).glob('*.pptx'), key=lambda p: p.stat().st_mtime, reverse=True)
    others = [p for p in ready if p.name != current.name]

    def meta(p):
        n = slide_count(p)
        bits = ([f'{n} слайд' + ('ов' if n % 10 in (0, 5, 6, 7, 8, 9) or 11 <= n % 100 <= 14
                                  else 'а' if n % 10 in (2, 3, 4) else '')] if n else [])
        bits += [size_label(p), datetime.fromtimestamp(p.stat().st_mtime).strftime('%d.%m.%Y %H:%M')]
        return ' · '.join(bits)

    # embed: файл и превью встроены в страницу — просмотрщик чата открывает её в изолированном iframe,
    # где браузер может не передать cookie входа, и тогда относительные ссылки на файлы терминала не работают.
    prev = root / SERVICE / PREVIEWS / (current.stem + '.jpg')
    if embed:
        src = f'data:image/jpeg;base64,{b64(prev)}' if prev.exists() else ''
        dl = f'data:{PPTX_MIME};base64,{b64(current)}'
        rows = ''.join(f'<li>{html.escape(p.stem)}<span>{meta(p)}</span></li>' for p in others)
        script, hint = SCRIPT, hint_html(current.name)
    else:
        src = href(PREVIEWS, prev.name) if prev.exists() else ''
        dl = href('..', READY, current.name)
        rows = ''.join(
            f'<li><a href="{href("..", READY, p.name)}" download="{html.escape(p.name)}">{html.escape(p.stem)}</a>'
            f'<span>{meta(p)}</span></li>' for p in others)
        script, hint = '', hint_html(current.name)
    img = f'<img src="{src}" alt="Обзор слайдов">' if src else ''
    older = (f'<h2>Другие готовые презентации (в той же папке)</h2><ul>{rows}</ul>') if others else ''
    return f'''<!DOCTYPE html>
<html lang="ru"><head><meta charset="utf-8"><title>Готовые презентации VITOMEK</title>
<style>
 body{{font-family:Arial,Helvetica,sans-serif;margin:0;padding:20px;color:#0C4C4C;background:#fff}}
 .card{{border:1px solid #e3e3e3;border-radius:12px;padding:18px;max-width:760px}}
 .tag{{font-size:12px;color:#6f9a1e;text-transform:uppercase;letter-spacing:.06em}}
 h1{{font-size:20px;margin:6px 0 4px}} .meta{{color:#7a7a7a;font-size:13px;margin-bottom:14px}}
 .btn{{display:inline-block;background:#0C4C4C;color:#fff;text-decoration:none;font-weight:bold;
   padding:12px 22px;border-radius:8px;font-size:15px;cursor:pointer}} .btn:hover{{background:#4E7C7C}}
 img{{display:block;width:100%;border:1px solid #eee;border-radius:8px;margin-top:16px}}
 .where{{font-size:12px;color:#7a7a7a;margin-top:12px}}
 h2{{font-size:15px;margin:24px 0 8px}} ul{{list-style:none;padding:0;margin:0;max-width:760px}}
 li{{display:flex;justify-content:space-between;gap:12px;padding:9px 0;border-bottom:1px solid #eee;font-size:14px}}
 li a{{color:#0C4C4C}} li span{{color:#7a7a7a;font-size:12px;white-space:nowrap}}
 .hint{{font-size:13px;color:#4E7C7C;background:#F3EED0;border-radius:8px;padding:10px 14px;margin-top:14px}}
 .hint ol{{margin:6px 0 0;padding-left:20px}} .hint li{{display:list-item;border:0;padding:2px 0}}
</style></head><body>
<div class="card">
 <div class="tag">Готовая презентация</div>
 <h1>{html.escape(current.stem)}</h1>
 <div class="meta">{meta(current)}</div>
 <a class="btn" id="dl" download="{html.escape(current.name)}"
    href="{dl}">⬇ Скачать .pptx</a>
 {hint}
 {img}
 <div class="where">Файл также лежит в папке терминала: витомэк-презентации / {READY}</div>
</div>
{older}
{script}
</body></html>
'''


def main():
    ap = argparse.ArgumentParser(description='Выложить готовую презентацию и обновить страницу скачивания')
    ap.add_argument('pptx', help='собранный .pptx')
    ap.add_argument('--root', default=str(Path.home() / 'витомэк-презентации'))
    ap.add_argument('--link', action='store_true', help='не встраивать файл: ссылки + подсказка про панель «Файлы»')
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
    out.write_text(page(root, dst, embed=not a.link), encoding='utf-8')
    print(f'Готово: {dst}')
    print(f'Страница скачивания: {out}')
    print(f'Покажи её в чате: display_file(path="{out}", inline=true)')
    print('\n=== ТЕКСТ ДЛЯ ОТВЕТА: скопируй дословно в начало ответа пользователю ===')
    print(answer_text(dst.name))
    print('=== КОНЕЦ ТЕКСТА ===')


if __name__ == '__main__':
    main()
