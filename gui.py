# -*- coding: utf-8 -*-
"""
Окно запуска. Всё, что должен сделать человек: выбрать месяц и нажать кнопку.
"""

import os
import sys
import subprocess
import datetime as dt
import threading
import traceback

import tkinter as tk
from tkinter import ttk, messagebox

import sborka

МЕСЯЦЫ = ["styczeń", "luty", "marzec", "kwiecień", "maj", "czerwiec", "lipiec",
          "sierpień", "wrzesień", "październik", "listopad", "grudzień"]
МЕСЯЦЫ_RU = ["январь", "февраль", "март", "апрель", "май", "июнь", "июль",
             "август", "сентябрь", "октябрь", "ноябрь", "декабрь"]

ФОН = "#F4F4F2"
ЗЕЛЁНЫЙ = "#1F7A3D"
КРАСНЫЙ = "#B03030"


def открыть_папку(путь):
    """Открыть проводник на нужной папке."""
    try:
        if sys.platform == "win32":
            os.startfile(путь)
        elif sys.platform == "darwin":
            subprocess.run(["open", путь])
        else:
            subprocess.run(["xdg-open", путь])
    except Exception:
        pass


class Окно:
    def __init__(self, корень):
        self.корень = корень
        self.результат = None
        корень.title("Расчёт выплат водителям")
        корень.configure(bg=ФОН)
        корень.resizable(False, False)

        рамка = tk.Frame(корень, bg=ФОН, padx=28, pady=22)
        рамка.pack(fill="both", expand=True)

        tk.Label(рамка, text="Расчёт выплат водителям", bg=ФОН,
                 font=("Segoe UI", 16, "bold")).pack(anchor="w")
        tk.Label(рамка, text="Выберите месяц и нажмите кнопку", bg=ФОН,
                 fg="#666666", font=("Segoe UI", 10)).pack(anchor="w", pady=(2, 16))

        # --- выбор периода ---
        строка = tk.Frame(рамка, bg=ФОН)
        строка.pack(anchor="w", pady=(0, 18))

        прошлый = dt.date.today().replace(day=1) - dt.timedelta(days=1)

        tk.Label(строка, text="Месяц:", bg=ФОН,
                 font=("Segoe UI", 12)).grid(row=0, column=0, padx=(0, 8))
        self.месяц = ttk.Combobox(строка, width=14, state="readonly",
                                  font=("Segoe UI", 12),
                                  values=[f"{i+1:02d} — {м}" for i, м in enumerate(МЕСЯЦЫ_RU)])
        self.месяц.current(прошлый.month - 1)
        self.месяц.grid(row=0, column=1, padx=(0, 18))

        tk.Label(строка, text="Год:", bg=ФОН,
                 font=("Segoe UI", 12)).grid(row=0, column=2, padx=(0, 8))
        годы = [str(г) for г in range(прошлый.year - 2, прошлый.year + 2)]
        self.год = ttk.Combobox(строка, width=7, state="readonly",
                                font=("Segoe UI", 12), values=годы)
        self.год.set(str(прошлый.year))
        self.год.grid(row=0, column=3)

        # --- кнопка ---
        self.кнопка = tk.Button(рамка, text="Сформировать отчёт",
                                font=("Segoe UI", 13, "bold"),
                                bg=ЗЕЛЁНЫЙ, fg="white", activebackground="#17612F",
                                activeforeground="white", relief="flat",
                                padx=20, pady=12, cursor="hand2",
                                command=self.запустить)
        self.кнопка.pack(fill="x")

        # --- журнал ---
        self.журнал = tk.Text(рамка, height=11, width=62, relief="flat",
                              bg="white", fg="#333333", font=("Consolas", 9),
                              padx=10, pady=8, state="disabled", wrap="word")
        self.журнал.pack(fill="both", pady=(16, 0))
        self.журнал.tag_config("ok", foreground=ЗЕЛЁНЫЙ, font=("Consolas", 9, "bold"))
        self.журнал.tag_config("err", foreground=КРАСНЫЙ, font=("Consolas", 9, "bold"))
        self.журнал.tag_config("warn", foreground="#B07000")

        # --- нижние кнопки ---
        низ = tk.Frame(рамка, bg=ФОН)
        низ.pack(fill="x", pady=(12, 0))
        self.кнопка_отчёт = tk.Button(низ, text="Открыть отчёт", state="disabled",
                                      font=("Segoe UI", 10), relief="flat",
                                      bg="#E2E2DE", padx=14, pady=7, cursor="hand2",
                                      command=self.открыть_отчёт)
        self.кнопка_отчёт.pack(side="left")
        tk.Button(низ, text="Папка «Источники»", font=("Segoe UI", 10), relief="flat",
                  bg="#E2E2DE", padx=14, pady=7, cursor="hand2",
                  command=lambda: (os.makedirs(sborka.ПАПКА_ИСТОЧНИКОВ, exist_ok=True),
                                   открыть_папку(sborka.ПАПКА_ИСТОЧНИКОВ))
                  ).pack(side="left", padx=8)
        tk.Button(низ, text="Папка «Отчеты»", font=("Segoe UI", 10), relief="flat",
                  bg="#E2E2DE", padx=14, pady=7, cursor="hand2",
                  command=lambda: (os.makedirs(sborka.ПАПКА_ОТЧЕТОВ, exist_ok=True),
                                   открыть_папку(sborka.ПАПКА_ОТЧЕТОВ))
                  ).pack(side="left")

        self.проверить_исходники()

    # ------------------------------------------------------------------
    def писать(self, текст, тег=None):
        self.журнал.configure(state="normal")
        self.журнал.insert("end", str(текст) + "\n", тег)
        self.журнал.see("end")
        self.журнал.configure(state="disabled")
        self.корень.update_idletasks()

    def очистить(self):
        self.журнал.configure(state="normal")
        self.журнал.delete("1.0", "end")
        self.журнал.configure(state="disabled")

    def проверить_исходники(self):
        папка = sborka.ПАПКА_ИСТОЧНИКОВ
        есть = os.path.isdir(папка) and [f for f in os.listdir(папка)
                                         if f.lower().endswith(".xlsx")
                                         and not f.startswith("~$")]
        if есть:
            self.писать(f"Найдено файлов в папке «Источники»: {len(есть)}")
            for f in есть:
                self.писать(f"   • {f}")
        else:
            self.писать("Папка «Источники» пуста.", "err")
            self.писать("Положите туда календарь водителей и нажмите кнопку "
                        "«Папка Источники», чтобы её открыть.", "warn")

    # ------------------------------------------------------------------
    def запустить(self):
        год = int(self.год.get())
        месяц = int(self.месяц.get().split(" ")[0])
        self.кнопка.configure(state="disabled", text="Считаю...")
        self.кнопка_отчёт.configure(state="disabled")
        self.очистить()
        threading.Thread(target=self.работа, args=(год, месяц), daemon=True).start()

    def работа(self, год, месяц):
        try:
            путь, водителей, проблемы = sborka.собрать(год, месяц, лог=self.писать)
            self.результат = путь
            self.писать("")
            self.писать("ГОТОВО", "ok")
            self.писать(f"Файл: {os.path.basename(путь)}")
            if проблемы:
                self.писать(f"Расхождений: {len(проблемы)} — "
                            f"смотрите лист «{месяц:02d}.{год} uwagi»", "warn")
            else:
                self.писать("Расхождений не найдено.")
            self.кнопка_отчёт.configure(state="normal")
        except Exception as e:
            self.писать("")
            self.писать("ОШИБКА", "err")
            self.писать(str(e))
            подробно = traceback.format_exc()
            with open(os.path.join(sborka.БАЗА, "ошибка.txt"), "w", encoding="utf-8") as f:
                f.write(подробно)
            self.писать("Подробности записаны в файл «ошибка.txt» рядом "
                        "с программой — покажите его тому, кто настраивал.", "warn")
        finally:
            self.кнопка.configure(state="normal", text="Сформировать отчёт")

    def открыть_отчёт(self):
        if self.результат and os.path.exists(self.результат):
            открыть_папку(self.результат)


if __name__ == "__main__":
    корень = tk.Tk()
    try:
        корень.call("tk", "scaling", 1.3)
    except Exception:
        pass
    Окно(корень)
    корень.mainloop()
