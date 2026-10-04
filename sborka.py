# -*- coding: utf-8 -*-
"""
Сборка ведомости выплат водителям за месяц.

Источник  : 2026_Kalenda_*.xlsx — листы «Wyjazdy» (выезды), «Zjazdy» (съезды),
            «2026» (реестр машин и водителей).
Справочники: LISTA_PRACOWNIKÓW.xlsx, справочник_ставок.xlsx,
            синонимы_имён.xlsx (необязательно), исключения.xlsx (необязательно).
Результат : Wyplaty_ГГГГ-ММ_N.xlsx в папке «Отчеты». Файлы не перезаписываются:
            повторный расчёт того же периода получает следующий индекс.

ЛОГИКА ПОДСЧЁТА ДНЕЙ (подтверждена клиентом, см. ВОПРОСЫ_К_СВЕРКЕ):
  * Первое событие водителя в месяце — ВЫЕЗД → до него он стоял (отпуск/сервис),
    период считается с этого выезда.
  * Первое событие в месяце — СЪЕЗД → водитель был в рейсе с 1-го числа.
  * Событий в месяце нет вовсе, но в реестре «2026» выезд раньше месяца →
    водитель весь месяц в рейсе.
  * Повторный выезд, пока водитель уже в рейсе (сервис, замена авто) — рейс
    НЕ прерывает.
  * Дни считаются включительно, один день не оплачивается дважды.

Запуск: python sborka.py 2026 8
"""

import sys, os, glob, calendar, datetime as dt, unicodedata
from collections import defaultdict

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# ============================ НАСТРОЙКИ ===================================

if getattr(sys, "frozen", False):
    БАЗА = os.path.dirname(sys.executable)
else:
    БАЗА = os.path.dirname(os.path.abspath(__file__))

ПАПКА_ИСТОЧНИКОВ = os.path.join(БАЗА, "Источники")
ПАПКА_ОТЧЕТОВ    = os.path.join(БАЗА, "Отчеты")
ИМЯ_ОТЧЕТА       = "Wyplaty"
ФАЙЛ_СОТРУДНИКОВ = os.path.join(БАЗА, "LISTA_PRACOWNIKÓW.xlsx")
ФАЙЛ_СТАВОК      = os.path.join(БАЗА, "справочник_ставок.xlsx")
ФАЙЛ_СИНОНИМОВ   = os.path.join(БАЗА, "синонимы_имён.xlsx")
ФАЙЛ_ИСКЛЮЧЕНИЙ  = os.path.join(БАЗА, "исключения.xlsx")

ЛИСТ_ВЫЕЗДОВ = "Wyjazdy"
ЛИСТ_СЪЕЗДОВ = "Zjazdy"
ЛИСТ_РЕЕСТРА = "2026"

# Колонки журналов Wyjazdy/Zjazdy (с 1): B=дата, C=машина, D=водитель, E=причина
К_ДАТА, К_МАШИНА, К_ВОДИТЕЛЬ, К_ПРИЧИНА = 2, 3, 4, 5
# Реестр «2026»: F=Nr. rejestracyjny, G=Kierowca, H=Data wyjazdu
Р_МАШИНА, Р_ВОДИТЕЛЬ, Р_ВЫЕЗД = 6, 7, 8

СТАВКА_ПО_УМОЛЧАНИЮ = 320.0
КУРС_EUR = 4.2359            # колонка AB; правится здесь раз в месяц

СУФФИКС_ПРОБЛЕМ = " uwagi"

# ==========================================================================

проблемы = []
показ = {}      # ключ(водитель) -> как фамилия записана в журнале
def проблема(файл, лист, строка, что, детали=""):
    проблемы.append((файл, лист, строка, что, детали))

_ПОЛЬСКИЕ = str.maketrans({"ą":"a","ć":"c","ę":"e","ł":"l","ń":"n","ó":"o",
                           "ś":"s","ź":"z","ż":"z"})
def ключ(s):
    """Имя для сравнения: регистр, пробелы, ё→е, польские диакритики сняты."""
    t = " ".join(str(s or "").strip().lower().replace("ё", "е").split())
    t = t.translate(_ПОЛЬСКИЕ)
    return "".join(c for c in unicodedata.normalize("NFKD", t)
                   if not unicodedata.combining(c))

def дата(v):
    if v is None: return None
    if isinstance(v, dt.datetime): return v.date()
    if isinstance(v, dt.date): return v
    if isinstance(v, (int, float)):
        try: return dt.date(1899, 12, 30) + dt.timedelta(days=int(v))
        except Exception: return None
    t = str(v).strip()
    for f in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try: return dt.datetime.strptime(t, f).date()
        except ValueError: pass
    return None


# ------------------------- чтение источников ------------------------------

def прочитать_журналы():
    """
    Возвращает:
      события[ключ(водитель)] = список (дата, 'out'|'in', машина, причина, строка, файл)
      реестр[ключ(водитель)]  = (дата выезда из реестра, машина)
      показ[ключ]             = как фамилия записана в журнале (для отчёта)
    """
    события = defaultdict(list)
    реестр  = {}
    показ   = {}

    файлы = [f for f in sorted(glob.glob(os.path.join(ПАПКА_ИСТОЧНИКОВ, "*.xlsx")))
             if not os.path.basename(f).startswith("~$")]
    if not файлы:
        raise SystemExit("В папке «Источники» нет файлов .xlsx")

    for путь in файлы:
        имя = os.path.basename(путь)
        try:
            книга = openpyxl.load_workbook(путь, data_only=True, read_only=True)
        except Exception as e:
            проблема(имя, "", "", "Не удалось открыть файл", str(e)); continue
        if ЛИСТ_ВЫЕЗДОВ not in книга.sheetnames:
            книга.close(); continue

        def читать_события(лист_имя, тип):
            if лист_имя not in книга.sheetnames: return
            for н, стр in enumerate(книга[лист_имя].iter_rows(values_only=True), start=1):
                def я(i): return стр[i-1] if стр and i-1 < len(стр) else None
                водитель = str(я(К_ВОДИТЕЛЬ) or "").strip()
                if not водитель: continue
                if ключ(водитель) in ("kierowca", "водитель"): continue
                д = дата(я(К_ДАТА))
                if д is None:
                    проблема(имя, лист_имя, н, "Не читается дата", водитель); continue
                k = ключ(водитель)
                показ.setdefault(k, водитель)
                события[k].append((д, тип, я(К_МАШИНА), я(К_ПРИЧИНА), н, имя))

        читать_события(ЛИСТ_ВЫЕЗДОВ, "out")
        читать_события(ЛИСТ_СЪЕЗДОВ, "in")

        if ЛИСТ_РЕЕСТРА in книга.sheetnames:
            for стр in книга[ЛИСТ_РЕЕСТРА].iter_rows(min_row=2, values_only=True):
                def я(i): return стр[i-1] if стр and i-1 < len(стр) else None
                водитель = str(я(Р_ВОДИТЕЛЬ) or "").strip()
                д = дата(я(Р_ВЫЕЗД))
                if водитель and д:
                    k = ключ(водитель)
                    показ.setdefault(k, водитель)
                    реестр[k] = (д, я(Р_МАШИНА))
        книга.close()
    return события, реестр, показ


def прочитать_сотрудников():
    карта = {}
    if not os.path.exists(ФАЙЛ_СОТРУДНИКОВ): return карта
    книга = openpyxl.load_workbook(ФАЙЛ_СОТРУДНИКОВ, data_only=True)
    for стр in книга.active.iter_rows(min_row=2, values_only=True):
        if not стр or not стр[1]: continue
        имя = str(стр[1]).strip()
        карта[ключ(имя)] = имя
    книга.close()
    return карта


def прочитать_ставки():
    ставки, общая = {}, СТАВКА_ПО_УМОЛЧАНИЮ
    if not os.path.exists(ФАЙЛ_СТАВОК): return ставки, общая
    книга = openpyxl.load_workbook(ФАЙЛ_СТАВОК, data_only=True)
    for стр in книга.active.iter_rows(min_row=2, values_only=True):
        if not стр or стр[1] is None: continue
        if стр[0]: ставки[ключ(стр[0])] = float(стр[1])
        else: общая = float(стр[1])
    книга.close()
    return ставки, общая


def прочитать_справочник(путь, колонок=2):
    """Простая таблица ключ→значение из первого листа; пусто, если файла нет."""
    рез = {}
    if not os.path.exists(путь): return рез
    книга = openpyxl.load_workbook(путь, data_only=True)
    for стр in книга.active.iter_rows(min_row=2, values_only=True):
        if not стр or not стр[0]: continue
        рез[ключ(стр[0])] = (str(стр[1]).strip() if колонок > 1 and len(стр) > 1
                             and стр[1] is not None else "")
    книга.close()
    return рез


# ----------------------------- расчёт -------------------------------------

def посчитать(события, реестр, год, месяц, синонимы, исключения,
              сотрудники, ставки, общая):
    первое = dt.date(год, месяц, 1)
    последнее = dt.date(год, месяц, calendar.monthrange(год, месяц)[1])
    данные = {}

    все = set(события) | set(реестр)
    for k in sorted(все):
        if k in исключения:
            continue
        # каноническое имя: если есть синоним — приводим к нему, затем к LISTA
        kk = ключ(синонимы[k]) if k in синонимы else k

        соб = sorted(события.get(k, []),
                     key=lambda x: (x[0], 0 if x[1] == "in" else 1))
        маш_реестра = реестр[k][1] if k in реестр else None

        aug = [(д, тип, м, прич, н, файл) for (д, тип, м, прич, н, файл) in соб
               if первое <= д <= последнее]

        # периоды: (начало, конец, машина этого периода)
        периоды, flags = [], []
        if aug:
            on = False; start = None; маш = None
            if aug[0][1] == "in":
                on = True; start = первое; маш = aug[0][2]
            for д, тип, м, прич, н, файл in aug:
                if тип == "out":
                    if not on:
                        on = True; start = д; маш = м
                    # иначе — сервисный повторный выезд, рейс не прерываем
                else:
                    if on:
                        периоды.append((start, д, маш)); on = False
                    else:
                        flags.append(f"съезд {д:%d.%m} без выезда")
            if on:
                периоды.append((start, последнее, маш))
            источник = "журналы Wyjazdy/Zjazdy"
        else:
            if k in реестр and реестр[k][0] < первое:
                периоды.append((первое, последнее, маш_реестра))
                источник = f"нет событий; в рейсе с {реестр[k][0]:%d.%m.%Y}"
            else:
                continue  # в этом месяце не ездил

        # отрезки без двойного счёта дня
        отрезки = []; prev = None
        for a, b, м in sorted(периоды, key=lambda x: x[0]):
            if prev and a <= prev:
                a = prev + dt.timedelta(days=1)
            if a > b: continue
            # машину берём из реестра «2026» (текущее закрепление),
            # из журнала — только если водителя в реестре нет
            отрезки.append((a, b, маш_реестра or м)); prev = b
        if not отрезки:
            continue

        если_нет_ставки = kk not in ставки
        имя_вывод = (сотрудники.get(kk) or синонимы.get(k)
                     or показ.get(k) or k.upper())
        if kk not in сотрудники:
            проблема(соб[0][5] if соб else "", ЛИСТ_ВЫЕЗДОВ,
                     соб[0][4] if соб else "",
                     "ФИО нет в списке сотрудников", имя_вывод)

        данные[k] = dict(имя=имя_вывод, отрезки=отрезки,
                         источник=источник, ставка=ставки.get(kk, общая),
                         нет_ставки=если_нет_ставки, flags=flags)
    return данные, первое, последнее


# ------------------- имя файла (по периоду, без перезаписи) ---------------

def следующее_имя_файла(год, месяц):
    шаблон = f"{ИМЯ_ОТЧЕТА}_{год}-{месяц:02d}_"
    есть = glob.glob(os.path.join(ПАПКА_ОТЧЕТОВ, шаблон + "*.xlsx"))
    номера = []
    for п in есть:
        х = os.path.basename(п)[len(шаблон):-5]
        if х.isdigit(): номера.append(int(х))
    n = max(номера) + 1 if номера else 1
    return os.path.join(ПАПКА_ОТЧЕТОВ, f"{шаблон}{n}.xlsx")


# --------------------------- запись отчёта --------------------------------

A = lambda s=10, b=False, i=False: Font(name="Arial", size=s, bold=b, italic=i)
РАМКА = Border(*[Side(style="thin", color="B0B0B0")] * 4)
ШАПКА = PatternFill("solid", fgColor="D9D9D9")
ИТОГ = PatternFill("solid", fgColor="F2F2F2")
ЗАПОЛНИ = PatternFill("solid", fgColor="FFF2CC")

ГРУППЫ = [("I", "M", "PLN"), ("N", "Q", "plus euro"), ("R", "T", "minus euro"),
          ("U", "V", "plus PLN"), ("W", "AA", "minus PLN")]
ПОДПИСИ = {"B": "nr rej.", "C": "Kierowca", "D": "wyjazd", "E": "zjazd",
           "F": "wypłata", "G": "status", "I": "stawka", "J": "dni",
           "K": "urlop", "L": "zaliczka", "M": "zaliczka", "N": "premia",
           "O": "d post", "P": "tanken", "Q": "winieta", "R": "Kaucja",
           "S": "Mandat", "T": "tanken", "U": "wytraty", "V": "premia",
           "W": "prędk/km", "X": "tlf.inter", "Y": "kont", "Z": "kont",
           "AA": "obciążen", "AB": "kurs", "AE": "uwagi (auto)"}
РУЧНЫЕ = ["K", "L", "M", "N", "O", "P", "Q", "R", "S", "T",
          "U", "V", "W", "X", "Y", "Z", "AA"]
ФОРМУЛА = ("=I{r}*J{r}+(K{r}*100)-(L{r}+M{r})+(N{r}+O{r}+P{r}+Q{r})*AB{r}"
           "-(R{r}+S{r}+T{r})*AB{r}+U{r}+V{r}-(W{r}+X{r}+Y{r}+Z{r}+AA{r})")


def записать(данные, первое, последнее, книга, лист_месяца):
    ws = книга.create_sheet(лист_месяца)
    for c1, c2, текст in ГРУППЫ:
        ws.merge_cells(f"{c1}1:{c2}1")
        я = ws[f"{c1}1"]; я.value = текст
        я.font, я.fill, я.alignment = A(10, True), ШАПКА, Alignment(horizontal="center")
    for кол, текст in ПОДПИСИ.items():
        я = ws[f"{кол}2"]; я.value = текст
        я.font, я.fill, я.border = A(9, True), ШАПКА, РАМКА
        я.alignment = Alignment(horizontal="center", wrap_text=True)

    r = 3
    for k in sorted(данные, key=lambda x: данные[x]["имя"]):
        з = данные[k]
        первая = r
        пометки_общие = []
        if з["нет_ставки"]:
            пометки_общие.append(f"stawka domyślna {СТАВКА_ПО_УМОЛЧАНИЮ:.0f}")
        пометки_общие += з["flags"]
        for idx, (a, b, маш) in enumerate(з["отрезки"]):
            пометки = list(пометки_общие) if idx == 0 else []
            if a == первое and "в рейсе" in з["источник"]:
                пометки.append("w trasie od początku miesiąca")
            if b == последнее:
                пометки.append("w trasie na koniec miesiąca")
            ws[f"B{r}"] = маш
            ws[f"C{r}"] = з["имя"]
            ws[f"D{r}"] = a
            ws[f"E{r}"] = b
            ws[f"F{r}"] = ФОРМУЛА.format(r=r)
            ws[f"I{r}"] = з["ставка"]
            ws[f"J{r}"] = f"=E{r}-D{r}+1"
            ws[f"AB{r}"] = КУРС_EUR
            ws[f"AE{r}"] = "; ".join(пометки)
            for кол in РУЧНЫЕ:
                ws[f"{кол}{r}"].fill = ЗАПОЛНИ
            for кол in list(ПОДПИСИ) + ["H"]:
                я = ws[f"{кол}{r}"]; я.font, я.border = A(), РАМКА
            ws[f"D{r}"].number_format = ws[f"E{r}"].number_format = "dd.mm.yyyy"
            ws[f"F{r}"].number_format = "0.00"
            ws[f"AE{r}"].font = A(8, i=True)
            r += 1
        if len(з["отрезки"]) > 1:
            ws[f"C{r}"] = f"{з['имя']} — RAZEM"
            ws[f"J{r}"] = f"=SUM(J{первая}:J{r-1})"
            ws[f"F{r}"] = f"=SUM(F{первая}:F{r-1})"
            ws[f"F{r}"].number_format = "0.00"
            for кол in list(ПОДПИСИ) + ["H"]:
                я = ws[f"{кол}{r}"]; я.font, я.fill, я.border = A(10, True), ИТОГ, РАМКА
            r += 1
        r += 1

    for кол, ш in {"A": 4.25, "B": 10.13, "C": 21.5, "D": 10, "E": 10, "F": 11,
                   "G": 10, "I": 6.88, "J": 6.38, "K": 9.75, "AB": 7.75,
                   "AE": 42}.items():
        ws.column_dimensions[кол].width = ш
    ws.freeze_panes = "D3"

    if not проблемы:
        return
    p = книга.create_sheet(лист_месяца + СУФФИКС_ПРОБЛЕМ)
    p["A1"] = f"Расхождения за {лист_месяца} — проверьте вручную"
    p["A1"].font = A(11, True)
    p["A2"] = ("Жёлтые ячейки на листе Wypłaty заполняются вручную: "
               "urlop, zaliczki, premie, mandaty и прочее.")
    p["A2"].font = A(9, i=True)
    for j, т in enumerate(["Файл", "Лист", "Строка", "Что не так", "Детали"], start=1):
        я = p.cell(row=4, column=j, value=т); я.font, я.fill, я.border = A(10, True), ШАПКА, РАМКА
    for i, зап in enumerate(проблемы, start=5):
        for j, v in enumerate(зап, start=1):
            я = p.cell(row=i, column=j, value=v); я.font, я.border = A(), РАМКА
    for кол, ш in {"A": 32, "B": 14, "C": 8, "D": 40, "E": 44}.items():
        p.column_dimensions[кол].width = ш


# ---------------------------------------------------------------------------

def собрать(год, месяц, лог=print):
    проблемы.clear()
    лог(f"Расчётный период: {месяц:02d}.{год}")

    события, реестр, показ_имён = прочитать_журналы()
    всего_соб = sum(len(v) for v in события.values())
    лог(f"Прочитано событий: {всего_соб}, водителей в журналах: {len(события)}")

    сотрудники = прочитать_сотрудников()
    ставки, общая = прочитать_ставки()
    синонимы = прочитать_справочник(ФАЙЛ_СИНОНИМОВ, колонок=2)
    исключения = прочитать_справочник(ФАЙЛ_ИСКЛЮЧЕНИЙ, колонок=1)

    global показ
    показ = показ_имён

    данные, первое, последнее = посчитать(
        события, реестр, год, месяц, синонимы, исключения,
        сотрудники, ставки, общая)

    os.makedirs(ПАПКА_ОТЧЕТОВ, exist_ok=True)
    книга = openpyxl.Workbook(); книга.remove(книга.active)
    лист = f"{месяц:02d}.{год}"
    записать(данные, первое, последнее, книга, лист)
    путь = следующее_имя_файла(год, месяц)
    книга.save(путь)

    всего_дней = sum((b - a).days + 1
                     for з in данные.values() for a, b, _ in з["отрезки"])
    лог(f"Водителей в ведомости: {len(данные)}, человеко-дней: {всего_дней}")
    лог(f"Расхождений для ручной проверки: {len(проблемы)}")
    return путь, len(данные), list(проблемы)


def main():
    if len(sys.argv) >= 3:
        год, месяц = int(sys.argv[1]), int(sys.argv[2])
    else:
        п = dt.date.today().replace(day=1) - dt.timedelta(days=1)
        год, месяц = п.year, п.month
    путь, _, _ = собрать(год, месяц)
    print(f"\nГотово: {путь}")


if __name__ == "__main__":
    main()
