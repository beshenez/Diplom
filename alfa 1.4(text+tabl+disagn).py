import os
import random
import tkinter as tk
from tkinter import messagebox
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import A4

# ----------------------------------------------------------
# Регистрируем шрифты (обычный + жирный)
# ----------------------------------------------------------
pdfmetrics.registerFont(TTFont("Arial", r"C:\Windows\Fonts\arial.ttf"))
pdfmetrics.registerFont(TTFont("Arial-Bold", r"C:\Windows\Fonts\arialbd.ttf"))

PAGE_WIDTH, PAGE_HEIGHT = A4
LEFT_MARGIN = 40
RIGHT_MARGIN = 40
FONT_NAME = "Arial"
FONT_BOLD = "Arial-Bold"
FONT_SIZE = 12

# ----------------------------------------------------------
# Данные таблицы
# ----------------------------------------------------------
TABLE_ELEMENTS = [
    "Обечайка корпуса цилиндрическая", "Коническое днище №1", "Коническое днище №2",
    "Люк А", "Крышка люка А", "Люк Б", "Крышка люка Б",
    "Штуцер В1", "Штуцер В2", "Штуцер Г (отвод)", "Штуцер Д",
    "Переход штуцера Д", "Штуцер Е", "Штуцер Ж"
]

TABLE_MIN_THICKNESS = [
    8.9, 7.5, 7.7, 7.4, 24.9, 7.8, 35.5,
    9.7, 9.7, 12.2, 5.6, 8.1, 6.0, 9.1
]

# ----------------------------------------------------------
# Вспомогательные утилиты
# ----------------------------------------------------------
def fmt_num(v):
    """Форматируем число: точка -> запятая. Если не число — возвращаем как есть."""
    try:
        # если это число (float/int)
        if isinstance(v, (int, float)):
            s = f"{v:.1f}" if isinstance(v, float) else str(v)
        else:
            s = str(v)
        return s.replace('.', ',')
    except Exception:
        return str(v)

# ----------------------------------------------------------
# Функция создания PDF
# ----------------------------------------------------------
def create_protocol_pdf(data, table_data):
    save_dir = r"C:\Users\sanya\Desktop\UGNTU\DIPLOM\probniki"
    os.makedirs(save_dir, exist_ok=True)
    pdf_path = os.path.join(save_dir, "protokol_auto.pdf")

    c = canvas.Canvas(pdf_path, pagesize=A4)
    width, height = A4

    # ---------- Шрифты ----------
    pdfmetrics.registerFont(TTFont("Arial-Bold", r"C:\Windows\Fonts\arialbd.ttf"))
    c.setFont(FONT_NAME, FONT_SIZE)

    # ---------- Заголовок ----------
    c.drawRightString(width - RIGHT_MARGIN, height - 40, "Приложение 4.3")
    c.setFont("Arial-Bold", 16)
    c.drawString(width / 2 - 120, height - 90, "ПРОТОКОЛ №")

    # Прямоугольник вокруг номера
    protocol_num = data["Номер протокола"]
    num_x = width / 2 + 10
    num_y = height - 90
    num_width = c.stringWidth(protocol_num, "Arial-Bold", 16) + 10
    num_height = 20
    c.rect(num_x, num_y - 5, num_width, num_height)
    c.drawString(num_x + 5, num_y, protocol_num)

    # Подзаголовки
    c.setFont("Arial-Bold", 12)
    c.drawCentredString(width / 2, height - 115,
        "по ультразвуковому контролю (ультразвуковая толщинометрия - УЗТ)")
    c.drawCentredString(width / 2, height - 135, data["Описание емкости"])
    c.setLineWidth(2)
    c.line(LEFT_MARGIN, height - 150, width - RIGHT_MARGIN, height - 150)

    # ---------- Основной блок ----------
    c.setLineWidth(1)  # линии и подчеркивания — тонкие
    y = height - 180
    text_width_limit = width - LEFT_MARGIN - RIGHT_MARGIN - 2

    def wrap_text(text, max_width):
        """Перенос длинного текста по словам, с умным запасом"""
        words = text.split()
        lines = []
        current_line = ""

        # небольшой динамический запас (в пикселях)
        safe_margin = max_width * 0.985  # 98,5% от реальной ширины

        for word in words:
            test_line = (current_line + " " + word).strip()
            if c.stringWidth(test_line, FONT_NAME, FONT_SIZE) < safe_margin:
                current_line = test_line
            else:
                lines.append(current_line)
                current_line = word
        if current_line:
            lines.append(current_line)
        return lines




    # ---------------- Новый add_line ----------------
    def add_line(label, value, skip_parts=None):
        """Слова из skip_parts остаются обычным шрифтом и без подчеркивания."""
        nonlocal y
        if skip_parts is None:
            skip_parts = []

        c.setFont(FONT_NAME, FONT_SIZE)
        prefix = f"{label} — "
        prefix_width = c.stringWidth(prefix, FONT_NAME, FONT_SIZE)
        c.drawString(LEFT_MARGIN, y, prefix)

        # Основной текст с переносом
        lines = []
        current_line = ""
        safe_margin = text_width_limit - prefix_width

        words = value.split()
        for word in words:
            test_line = (current_line + " " + word).strip()
            if c.stringWidth(test_line, FONT_BOLD, FONT_SIZE) <= safe_margin:
                current_line = test_line
            else:
                lines.append(current_line)
                current_line = word
                # для следующих строк убираем префикс
                safe_margin = text_width_limit
        if current_line:
            lines.append(current_line)

        # Рисуем строки с жирностью и подчеркиванием
        for line_idx, line_text in enumerate(lines):
            x = LEFT_MARGIN + (prefix_width if line_idx == 0 else 0)
            y_offset = y - (line_idx * 18)
            underline_start = None

            for part in line_text.split():
                part_width = c.stringWidth(part + " ", FONT_BOLD, FONT_SIZE)
                skip = any(part.startswith(skip_word) for skip_word in skip_parts)

                # Выбираем шрифт
                if skip:
                    c.setFont(FONT_NAME, FONT_SIZE)  # обычный
                else:
                    c.setFont(FONT_BOLD, FONT_SIZE)  # жирный

                c.drawString(x, y_offset, part)

                # Подчеркивание
                if not skip:
                    if underline_start is None:
                        underline_start = x
                else:
                    if underline_start is not None:
                        delta = 2.83
                        c.line(underline_start, y_offset - 1.3, x - delta, y_offset - 1.3)
                        underline_start = None

                x += part_width

            # Завершаем подчеркивание, если линия не закрыта
            if underline_start is not None:
                c.line(underline_start, y_offset - 1.3, x, y_offset - 1.3)

        y -= (len(lines) * 18) + 6




    # Строки
    add_line("Место эксплуатации и проведения НК", data["Место эксплуатации"])
    add_line("Дата проведения измерений", data["Дата измерений"])
    add_line("Толщинометрия проводилась в соответствии с требованиями", data["Требования толщинометрии"])
    add_line("Оценка качества проводилась в соответствии с требованиями", data["Требования качества"])
    add_line("Прибор", f"{data['Прибор']} зав. № {data['Зав. № прибора']} дата очередной проверки {data['Дата проверки']}",
             skip_parts=["зав.", "дата", "очередной", "№", "проверки"],)
    add_line("Преобразователь", f"{data['Преобразователь']} зав. № {data['Зав. № преобразователя']}",
             skip_parts=["зав.", "№"])
    add_line("Поверхность ввода ультразвуковых колебаний", data["Поверхность ввода"])
    add_line("Материальное исполнение сосуда", data["Материал сосуда"])
    add_line("Обечайка корпуса цилиндрическая из стали марки",
             f"{data['Обечайка']} ГОСТ {data['ГОСТ обечайки']}", skip_parts=["ГОСТ"])
    add_line("Днища корпуса эллиптические из стали марки",
             f"{data['Днища']} ГОСТ {data['ГОСТ днищ']}", skip_parts=["ГОСТ"])


    y -= 10
    c.drawString(LEFT_MARGIN, y, "Результаты ультразвукового контроля приведены ниже:")
    y -= 20

    # ---------- Таблица ----------
    c.setFont(FONT_NAME, FONT_SIZE)
    table_x = LEFT_MARGIN
    col_widths = [230, 120, 150]  # увеличил первую колонку
    first_row_height = 40 
    row_height = 18

    # Многострочные заголовки
    header_lines = [
        ["Элемент"],
        ["Мин. толщина", "стенки 2025г, мм"],
        ["Отбраковка по РУА-93,", "мм"]
    ]

    # Верхняя граница таблицы
    table_top_y = y
    c.line(table_x, y + 5, table_x + sum(col_widths), y + 5)

    # Заголовок
    header_y = y - 15
    for col_idx, lines in enumerate(header_lines):
        col_left = table_x + sum(col_widths[:col_idx])
        col_center = col_left + col_widths[col_idx]/2
        for line_idx, line_text in enumerate(lines):
            line_y = header_y - line_idx*12
            if col_idx == 0:  # первый столбец
                c.setFont(FONT_BOLD, FONT_SIZE)
                c.drawCentredString(col_center, line_y, line_text)
            else:  # второй и третий столбцы
                c.setFont(FONT_BOLD, FONT_SIZE)
                c.drawCentredString(col_center, line_y, line_text)



    # Вертикальные линии
    table_top_y = y + 5  # верх таблицы
    table_bottom_y = y - (first_row_height + row_height * len(TABLE_ELEMENTS)) + 3  # низ таблицы

    x_pos = table_x
    for w in col_widths:
        c.line(x_pos, table_top_y, x_pos, table_bottom_y)
        x_pos += w
    c.line(x_pos, table_top_y, x_pos, table_bottom_y)


    # Горизонтальная линия под заголовком
    c.line(table_x, y - first_row_height + 3, table_x + sum(col_widths), y - first_row_height + 3)

    # Заполнение таблицы
    y -= first_row_height
    for i, element in enumerate(TABLE_ELEMENTS):
        # первая колонка - название элемента обычным шрифтом, центрируем
        c.setFont(FONT_NAME, FONT_SIZE)
        c.drawString(table_x + 2, y - 11, element)

        # вторая колонка - мин. толщина жирным, центрируем
        c.setFont(FONT_BOLD, FONT_SIZE)
        col_x = table_x + col_widths[0]
        col_width = col_widths[1]
        text = str(TABLE_MIN_THICKNESS[i]).replace('.', ',')
        text_width = c.stringWidth(text, FONT_BOLD, FONT_SIZE)
        x_centered = col_x + (col_width - text_width) / 2
        c.drawString(x_centered, y - 11, text)

        # третья колонка - значение жирным, центрируем
        value = table_data.get(element, "")
        text = str(value).replace('.', ',')
        text_width = c.stringWidth(text, FONT_BOLD, FONT_SIZE)
        col_x = table_x + col_widths[0] + col_widths[1]
        col_width = col_widths[2]
        x_centered = col_x + (col_width - text_width) / 2
        c.drawString(x_centered, y - 11, text)

        # Горизонтальная линия
        y -= row_height
        c.line(table_x, y + 3, table_x + sum(col_widths), y + 3)


    # Нижняя граница таблицы
    c.line(table_x, y + 3, table_x + sum(col_widths), y + 3)

    c.save()
    return pdf_path


# ----------------------------------------------------------
# Имитация данных (заполняет левую форму и правую колонку таблицы)
# ----------------------------------------------------------
def auto_fill():
    # Левый столбец
    fields = {
        "Номер протокола": "1017/УНХ/УЗТ/25",
        "Описание емкости": "Емкость подземная поз. Е-1101/2, зав. №34120, рег. №108",
        "Место эксплуатации": "На открытой площадке переработки попутного нефтяного газа Южно-Балыкского ГПЗ АО «СибурТюменьГаз», УПГ",
        "Дата измерений": "14 июня 2025 года",
        "Требования толщинометрии": "ГОСТ Р ИСО 16809-2015 п. 3.1 РУА-93",
        "Требования качества": "ГОСТ 34233.1÷12-2017",
        "Прибор": "А1270",
        "Зав. № прибора": "420695",
        "Дата проверки": "25.05.2026",
        "Преобразователь": "S3950 (5 MHz)",
        "Зав. № преобразователя": "1230076",
        "Поверхность ввода": "наружная / внутренняя",
        "Материал сосуда": "сталь низколегированная",
        "Обечайка": "09Г2Д",
        "ГОСТ обечайки": "5520",
        "Днища": "09Г2Д",
        "ГОСТ днищ": "5520"
    }
    for key, entry in entries.items():
        entry.delete(0, tk.END)
        entry.insert(0, fields[key])

    # Правый столбец - случайные значения таблицы (с одной десятичной)
    for element, entry in table_entries.items():
        val = round(random.uniform(5.0, 20.0), 1)
        # записываем с точкой в поле интерфейса (в PDF преобразим в запятую)
        entry.delete(0, tk.END)
        entry.insert(0, str(val))

# ----------------------------------------------------------
# Создание PDF и открытие
# ----------------------------------------------------------
def submit(open_after=False):
    data = {key: entry.get() for key, entry in entries.items()}
    table_data = {key: entry.get() for key, entry in table_entries.items()}
    pdf_path = create_protocol_pdf(data, table_data)
    messagebox.showinfo("Успех", f"PDF создан:\n{pdf_path}")
    if open_after:
        try:
            os.startfile(pdf_path)
        except Exception:
            pass

# ----------------------------------------------------------
# Интерфейс
# ----------------------------------------------------------
root = tk.Tk()
root.title("Создание протокола УЗТ")

labels = [
    "Номер протокола",
    "Описание емкости",
    "Место эксплуатации",
    "Дата измерений",
    "Требования толщинометрии",
    "Требования качества",
    "Прибор",
    "Зав. № прибора",
    "Дата проверки",
    "Преобразователь",
    "Зав. № преобразователя",
    "Поверхность ввода",
    "Материал сосуда",
    "Обечайка",
    "ГОСТ обечайки",
    "Днища",
    "ГОСТ днищ"
]

entries = {}
left_frame = tk.Frame(root)
left_frame.pack(side="left", padx=10, pady=10)
for label_text in labels:
    row = tk.Frame(left_frame)
    row.pack(fill="x", pady=2)
    lbl = tk.Label(row, text=label_text, width=35, anchor="w")
    lbl.pack(side="left")
    entry = tk.Entry(row, width=60)
    entry.pack(side="left", padx=5)
    entries[label_text] = entry

# Правая колонка - элементы таблицы
right_frame = tk.Frame(root)
right_frame.pack(side="left", padx=30, pady=10)
table_entries = {}
for element in TABLE_ELEMENTS:
    row = tk.Frame(right_frame)
    row.pack(fill="x", pady=1)
    lbl = tk.Label(row, text=element, width=30, anchor="w")
    lbl.pack(side="left")
    entry = tk.Entry(row, width=10)
    entry.pack(side="left", padx=2)
    table_entries[element] = entry

# ----------------------------------------------------------
# Кнопки
# ----------------------------------------------------------
button_frame = tk.Frame(root)
button_frame.pack(pady=10)
simulate_button = tk.Button(button_frame, text="Имитация", command=auto_fill, bg="#2196F3", fg="white", width=20)
simulate_button.pack(side="left", padx=5)
save_open_button = tk.Button(button_frame, text="Создать и открыть PDF", command=lambda: submit(open_after=True), bg="#4CAF50", fg="white", width=25)
save_open_button.pack(side="left", padx=5)

root.mainloop()
