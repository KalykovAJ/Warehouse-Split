import os
from tkinter import Tk, filedialog
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break
from collections import defaultdict


# =====================================================
# ОЧИСТКА ТЕКСТА
# =====================================================

def clean_text(text):
    if pd.isna(text):
        return ""
    text = str(text)
    text = text.replace("\n", " ")
    text = text.replace("\r", " ")
    text = text.replace(chr(160), " ")
    return " ".join(text.lower().split())


# =====================================================
# ЧТЕНИЕ CONFIG (ОПТИМИЗИРОВАННОЕ И БЕЗОПАСНОЕ)
# =====================================================

def read_config():
    config_path = os.path.join(os.path.dirname(__file__), "config.xlsx")
    
    if not os.path.exists(config_path):
        print(f"Ошибка: Файл конфигурации не найден по пути: {config_path}")
        exit()
        
    wb = load_workbook(config_path, data_only=True)
    
    # 1. Читаем первый (активный) лист с основными путями/настройками
    ws_main = wb.active
    config = {}
    for row in ws_main.iter_rows(min_row=1, values_only=True):
        if row and len(row) > 1:
            if row[0] and row[1]:
                config[str(row[0]).strip()] = str(row[1]).strip()

    # 2. Читаем второй лист со списком сетей АЗС
    try:
        ws_networks = wb.worksheets[1] 
    except IndexError:
        print("Ошибка: В файле config.xlsx не найден второй лист со списком сетей АЗС!")
        exit()

    networks_list = []
    for row in ws_networks.iter_rows(min_row=1, values_only=True):
        if row and row[0]:
            networks_list.append(str(row[0]).strip())
            
    config["__networks_list__"] = networks_list
    return config


# Загрузка конфигурации
config = read_config()

if "Справочник товаров по сетям" not in config:
    print("Ошибка: В config.xlsx на первом листе не найден параметр 'Справочник товаров по сетям'!")
    exit()

config_file = config["Справочник товаров по сетям"]
allowed_networks = config["__networks_list__"]

if not allowed_networks:
    print("Предупреждение: Список сетей на втором листе config.xlsx пуст!")


# =====================================================
# ВЫБОР ФАЙЛОВ ОПЕРАТОРОМ
# =====================================================

root = Tk()
root.withdraw()

main_file = filedialog.askopenfilename(
    title="Выберите основной файл для обработки",
    filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")]
)

if not main_file:
    print("Действие отменено пользователем.")
    exit()


# =====================================================
# ЗАГРУЗКА И СЛИЯНИЕ КНИГ EXCEL
# =====================================================

wb = load_workbook(main_file, keep_vba=True)

if not os.path.exists(config_file):
    print(f"Ошибка: Указанный в конфиге файл-справочник не найден:\n{config_file}")
    exit()

wb_config = load_workbook(config_file, data_only=True)
not_found = []


# =====================================================
# СТИЛИ ОФОРМЛЕНИЯ ТАБЛИЦ
# =====================================================

title_font = Font(name="Calibri", size=16, bold=True)
warehouse_font = Font(name="Calibri", size=15, bold=True)
header_font = Font(name="Calibri", size=14, bold=True)
text_font = Font(name="Calibri", size=13)

center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left = Alignment(horizontal="left", vertical="center", wrap_text=True)

gray_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
thin_side = Side(style="thin", color="000000")
border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)


# =====================================================
# ОСНОВНОЙ ЦИКЛ ОБРАБОТКИ ЛИСТОВ
# =====================================================

for ws in wb.worksheets:
    sheet_title_lower = ws.title.strip().lower()
    network_name = None
    
    for net in allowed_networks:
        if net.lower() in sheet_title_lower:
            network_name = net
            break
            
    if network_name is None:
        continue

    azs_name = ws.title.upper()

    config_sheet = None
    for sh in wb_config.worksheets:
        if sh.title.strip().lower() == network_name.strip().lower():
            config_sheet = sh
            break

    if config_sheet is None:
        print(f"В справочнике товаров отсутствует лист для сети: {network_name}")
        continue

    product_dict = {}
    for row in config_sheet.iter_rows(min_row=2, values_only=True):
        if row and row[0]:
            product = clean_text(row[0])
            warehouse = str(row[1]).strip() if len(row) > 1 and row[1] else ""
            if product:
                product_dict[product] = warehouse

    data = []
    for row in range(2, ws.max_row + 1):
        num = ws.cell(row=row, column=1).value
        product_name = ws.cell(row=row, column=2).value
        qty = ws.cell(row=row, column=3).value

        clean_product = clean_text(product_name)

        if clean_product == "" or clean_product in ["товар", "наименование", "продукт"]:
            continue

        warehouse = product_dict.get(clean_product, "НЕ НАЙДЕНО")

        if warehouse == "НЕ НАЙДЕНО":
            not_found.append((ws.title, product_name))

        data.append({
            "num": num,
            "product": product_name,
            "qty": qty,
            "warehouse": warehouse
        })

    warehouses = []
    for item in data:
        if item["warehouse"] not in warehouses:
            warehouses.append(item["warehouse"])

    ws.delete_rows(1, ws.max_row)
    ws.merged_cells.ranges = []
    ws.row_breaks = type(ws.row_breaks)()

    current_row = 1

    for wh in warehouses:
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=3)
        cell = ws.cell(current_row, 1, value=azs_name.upper())
        cell.font = title_font
        cell.alignment = Alignment(horizontal="left", vertical="center")
        for col in range(1, 4):
            ws.cell(current_row, col).border = border
        ws.row_dimensions[current_row].height = 30
        current_row += 1

        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=3)
        cell = ws.cell(current_row, 1, value=str(wh).upper())
        cell.font = warehouse_font
        cell.alignment = Alignment(horizontal="left", vertical="center")
        cell.fill = gray_fill
        for col in range(1, 4):
            ws.cell(current_row, col).border = border
        ws.row_dimensions[current_row].height = 25
        current_row += 1

        headers = ["№", "ТОВАР", "КОЛ-ВО"]
        for col, header in enumerate(headers, start=1):
            cell = ws.cell(current_row, col, value=header)
            cell.font = header_font
            cell.alignment = center
            cell.fill = gray_fill
            cell.border = border
        ws.row_dimensions[current_row].height = 22
        current_row += 1

        for item in data:
            if item["warehouse"] == wh:
                ws.cell(current_row, 1, value=item["num"])
                ws.cell(current_row, 2, value=item["product"])
                ws.cell(current_row, 3, value=item["qty"])

                for col in range(1, 4):
                    cell = ws.cell(current_row, col)
                    cell.font = text_font
                    cell.alignment = left
                    cell.border = border
                current_row += 1

        ws.row_breaks.append(Break(id=current_row))
        current_row += 2

    widths = {1: 10, 2: 70, 3: 15}
    for col_num, width in widths.items():
        ws.column_dimensions[get_column_letter(col_num)].width = width

    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = False
    ws.sheet_view.showGridLines = False


# =====================================================
# СОХРАНЕНИЕ РЕЗУЛЬТАТА С ЗАЩИТОЙ ОТ ПЕРЕЗАПИСИ
# =====================================================

base_name = os.path.splitext(main_file)[0]
output_file = f"{base_name}_отсортированный.xlsm"

# Если старый файл существует, удаляем его перед сохранением
if os.path.exists(output_file):
    os.remove(output_file)

wb.save(output_file)


# Логирование результатов в консоль (Группировка по листам с нумерацией)
if not_found:
    print("\n[ВНИМАНИЕ] Не найдены следующие позиции в справочнике:\n")
    
    # 1. Группируем товары по листам: { 'Имя_листа': {'Товар1', 'Товар2'} }
    grouped_errors = defaultdict(set)
    for sheet, item in not_found:
        grouped_errors[sheet].add(item)
        
    # 2. Выводим листы с нумерацией, а товары — красивым списком
    # Сортируем листы по алфавиту для порядка
    for i, sheet_name in enumerate(sorted(grouped_errors.keys()), start=1):
        print(f"{i}. Лист: \"{sheet_name}\"")
        
        # Сортируем товары внутри этого листа и выводим с маркером
        for item in sorted(grouped_errors[sheet_name]):
            print(f"   ▪ {item}")  # Вместо кубика можно использовать любой маркер: •, -, ▫
        print() # Пустая строка для визуального разделения листов
else:
    print("\n[ОК] Все товары успешно сопоставлены со справочником.")

print("Обработка завершена!")
print(f"Результат сохранен в файл:\n{output_file}")