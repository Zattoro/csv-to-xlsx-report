#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
csv_report.py — собирает огромное количество CSV-файлов в один аккуратный Excel-отчёт.

Что делает:
  1. Читает все CSV в папке (рекурсивно при -r), определяет разделитель (, ; таб).
  2. Складывает данные в один лист «Данные».
  3. Считает сводку: по каждому файлу число строк, по числовым колонкам — суммы.
  4. Строит сводную таблицу «по колонке X → сумма колонки Y» (аргументы --by, --sum).
  5. Сохраняет готовый .xlsx с форматированием (жирная шапка, автоширина, числа).

Пример:
    python csv_report.py --input "D:\\продажи_csv" --by "Менеджер" --sum "Сумма"
"""

from __future__ import annotations

import argparse
import csv
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    raise SystemExit("pip install openpyxl")


def sniff_sample(path: Path) -> dict:
    """Определяем разделитель CSV по образцу файла."""
    with open(path, "r", encoding="utf-8-sig", errors="replace", newline="") as f:
        sample = f.read(4096)
    result = {"delimiter": ","}
    if not sample:
        return result
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
        result["delimiter"] = dialect.delimiter
    except csv.Error:
        pass
    return result


def read_csv_files(root: Path, recursive: bool) -> tuple[list[Path], list[dict], list[str]]:
    files = sorted(p for p in (root.rglob("*.csv") if recursive else root.glob("*.csv"))
                   if p.is_file())
    if not files:
        raise SystemExit(f"CSV-файлы не найдены в {root}")

    rows: list[dict] = []
    skipped: list[str] = []
    header: list[str] = []
    for fp in files:
        with open(fp, "r", encoding="utf-8-sig", errors="replace", newline="") as f:
            reader = csv.DictReader(f, **sniff_sample(fp))
            if not reader.fieldnames:
                skipped.append(fp.name)
                continue
            fns = [c.strip() for c in reader.fieldnames]
            if not header:
                header = fns
            for r in reader:
                if r is None:
                    continue
                rows.append({k.strip(): (v.strip() if isinstance(v, str) else v)
                             for k, v in r.items() if k is not None})
    if skipped:
        print(f"Файлов без шапки пропущено: {len(skipped)}")
    return files, rows, header


def to_num(value: str) -> float:
    """В число: аккуратно с запятой как десятичным разделителем."""
    s = (value or "").replace(" ", "").replace("\u00a0", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def is_numbery(values: list[str]) -> bool:
    vals = [v for v in values if v not in ("", None)]
    return len(vals) > 0 and sum(
        1 for v in vals if to_num(str(v))) >= 0.8 * len(vals)


def fmt_header(ws, ncols: int) -> None:
    fill = PatternFill("solid", fgColor="2F5597")
    font = Font(bold=True, color="FFFFFF")
    for c in range(1, ncols + 1):
        cell = ws.cell(1, c)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center")


def autofit(ws, max_w: int = 60) -> None:
    for col in ws.columns:
        width = max((len(str(c.value)) for c in col if c.value is not None), default=8)
        ws.column_dimensions[get_column_letter(col[0].column)].width = \
            min(max_w, max(10, width + 2))


def build_xlsx(out: Path, files, rows: list[dict], header: list[str],
               numeric: set[str], by: str | None, sum_col: str | None) -> None:
    wb = Workbook()

    # 1) Сводка по файлам
    ws = wb.active
    ws.title = "Сводка"
    ws.append(["Свойство", "Значение"])
    ws.append(["Источник", str(files[0].parent)])
    ws.append(["Файлов CSV", len(files)])
    ws.append(["Строк данных", len(rows)])
    ws.append(["Колонок", len(header)])
    ws.append(["Числовые колонки", ", ".join(sorted(numeric)) or "-"])
    fmt_header(ws, 2)
    autofit(ws)

    # 2) Сводная таблица (если указано)
    if by and sum_col:
        ws2 = wb.create_sheet(f"Свод по {sum_col}")
        agg: dict[str, float] = defaultdict(float)
        cnt: dict[str, int] = defaultdict(int)
        for r in rows:
            key = str(r.get(by, "(пусто)"))
            agg[key] += to_num(r.get(sum_col, ""))
            cnt[key] += 1
        ws2.append([by, f"Сумма {sum_col}", "Кол-во строк"])
        for key in sorted(agg, key=lambda k: -agg[k]):
            ws2.append([key, round(agg[key], 2), cnt[key]])
        ws2.append(["ИТОГО", round(sum(agg.values()), 2), len(rows)])
        fmt_header(ws2, 3)
        autofit(ws2)

    # 3) Данные
    ws3 = wb.create_sheet("Данные")
    ws3.append(header)
    for r in rows:
        ws3.append([to_num(r[h]) if h in numeric else r.get(h, "") for h in header])
    fmt_header(ws3, len(header))
    for c, h in enumerate(header, 1):
        if h in numeric:
            fmt = get_column_letter(c)
            for i in range(2, len(rows) + 2):
                ws3[f"{fmt}{i}"].number_format = "#,##0.00"
    ws3.freeze_panes = "A2"
    ws3.auto_filter.ref = ws3.dimensions
    autofit(ws3)

    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--input", required=True, help="папка с CSV")
    ap.add_argument("-r", "--recursive", action="store_true")
    ap.add_argument("--by", help="колонка для сводной таблицы (категория)")
    ap.add_argument("--sum", dest="sum_col", help="числовая колонка для суммирования")
    ap.add_argument("--out", default=None,
                    help="путь итогового xlsx (по умолчанию <input>\\report.xlsx)")
    args = ap.parse_args()

    root = Path(args.input)
    if not root.is_dir():
        raise SystemExit(f"Папка не найдена: {args.input}")
    files, rows, header = read_csv_files(root, args.recursive)
    if not rows:
        raise SystemExit("Данные не найдены")

    numeric = set()
    for h in header:
        col = [str(r.get(h, "")) for r in rows]
        if is_numbery(col):
            numeric.add(h)

    by = args.by or (header[0] if header else None)
    sum_col = args.sum_col if args.sum_col in (numeric or set()) else None
    if args.sum_col and args.sum_col not in numeric:
        print(f"Внимание: колонка «{args.sum_col}» не числовая — сводная таблица строится "
              f"по «{next(iter(numeric), '—')}»")
        sum_col = next(iter(numeric), None) if numeric else None

    out = Path(args.out) if args.out else root / "report.xlsx"
    build_xlsx(out, files, rows, header, numeric, by, sum_col)

    print(f"Файлов: {len(files)} | строк: {len(rows)} | колонок: {len(header)}")
    print(f"Числовые: {', '.join(sorted(numeric)) or 'нет'}")
    print(f"Готово: {out}")


if __name__ == "__main__":
    main()
