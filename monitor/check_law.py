# -*- coding: utf-8 -*-
"""Проверка статей УПК РК, на которых построен ProDoc KZ.
Скачивает текст кодекса, считает отпечаток каждой нужной статьи и сравнивает с сохранённым.
Результат пишется в report.md (если есть изменения или ошибка). Только стандартная библиотека."""
import hashlib, html, json, os, re, sys, urllib.request

URL = os.environ.get("LAW_URL", "https://adilet.zan.kz/rus/docs/K1400000231")
WATCH = [30, 48, 49, 50, 64, 65, 71, 78, 99, 100, 102, 105, 106, 179, 180, 181, 208, 259, 261]
BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "baseline.json")
REPORT = "report.md"

def fetch():
    req = urllib.request.Request(URL, headers={"User-Agent": "Mozilla/5.0 (compatible; prodoc-law-monitor)"})
    with urllib.request.urlopen(req, timeout=60) as r:
        return r.read().decode("utf-8", errors="replace")

def to_text(raw):
    raw = re.sub(r"(?is)<(script|style).*?</\1>", "", raw)
    raw = re.sub(r"(?i)<(br|/p|/div|/li|/h\d|/tr|/section)[^>]*>", "\n", raw)
    t = html.unescape(re.sub(r"<[^>]+>", "", raw))
    t = t.replace("\xa0", " ")
    return re.sub(r"[ \t]+", " ", t)

def articles(text):
    """Возвращает {номер: текст статьи} для нужных статей."""
    out = {}
    heads = [(m.start(), m.group(1)) for m in re.finditer(r"(?m)^\s*Статья (\d+(?:-\d+)?)\.\s", text)]
    for i, (pos, num) in enumerate(heads):
        end = heads[i + 1][0] if i + 1 < len(heads) else len(text)
        if num.isdigit() and int(num) in WATCH and num not in out:
            out[num] = text[pos:end]
    return out, len(heads)

def norm(s):
    return re.sub(r"\s+", " ", s).strip()

def digest(s):
    return hashlib.sha256(norm(s).encode("utf-8")).hexdigest()

def write(msg):
    open(REPORT, "w", encoding="utf-8").write(msg)
    print(msg)

def main():
    try:
        text = to_text(fetch())
    except Exception as e:
        write("## Не удалось проверить законы\n\nСайт %s недоступен или не отвечает: `%s`.\n\nЭто не значит, что закон изменился. Проверка повторится по расписанию. Если ошибка держится несколько недель, нужно поправить `monitor/check_law.py`." % (URL, e))
        return
    arts, total = articles(text)
    missing = [n for n in WATCH if str(n) not in arts]
    if total < 100 or missing:
        write("## Не удалось разобрать текст кодекса\n\nНайдено заголовков статей: %d. Не найдены статьи: %s.\n\nВозможно, сайт изменил вёрстку. Нужно поправить `monitor/check_law.py`." % (total, ", ".join(map(str, missing)) or "нет"))
        return
    new = {n: digest(t) for n, t in arts.items()}
    if not os.path.exists(BASE):
        json.dump(new, open(BASE, "w"), indent=1)
        print("Первый запуск: сохранён эталон статей. Сообщений об изменениях нет.")
        open("baseline_created", "w").write("1")
        return
    old = json.load(open(BASE))
    changed = [n for n in new if old.get(n) != new[n]]
    if not changed:
        print("Изменений в проверяемых статьях нет.")
        return
    lines = ["## Изменились статьи УПК РК, на которых построен сайт\n",
             "Статьи: **%s**.\n" % ", ".join(sorted(changed, key=int)),
             "Что сделать: откройте актуальный текст на https://adilet.zan.kz/rus/docs/K1400000231, пришлите эти статьи Claude и попросите обновить сайт. Не меняйте шаблоны вручную без сверки.\n"]
    for n in sorted(changed, key=int):
        t = norm(arts[n])
        lines.append("\n### Статья %s\n\n%s%s\n" % (n, t[:1500], " ..." if len(t) > 1500 else ""))
    write("".join(lines))
    json.dump(new, open(BASE, "w"), indent=1)

if __name__ == "__main__":
    main()
