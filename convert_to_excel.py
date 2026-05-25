import pandas as pd

# Чтение CSV файла
df = pd.read_csv('Teen_Mental_Health_Dataset.csv')

# Создание Excel файла с отформатированными данными
output_file = 'Teen_Mental_Health_Dataset.xlsx'

# Запись в Excel с форматированием
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name='Teen Mental Health Data')
    
    # Получаем доступ к рабочему листу для форматирования
    worksheet = writer.sheets['Teen Mental Health Data']
    
    # Авто-ширина колонок
    for column in worksheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2) * 1.2
        worksheet.column_dimensions[column_letter].width = min(adjusted_width, 25)
    
    # Форматирование заголовков (жирный шрифт, фон)
    from openpyxl.styles import Font, PatternFill, Alignment
    
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=12)
    header_alignment = Alignment(horizontal="center", vertical="center")
    
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
    
    # Форматирование всех ячеек
    data_alignment = Alignment(horizontal="left", vertical="center")
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row):
        for cell in row:
            cell.alignment = data_alignment
    
    # Заморозка первой строки (заголовков)
    worksheet.freeze_panes = 'A2'

print(f"Файл успешно создан: {output_file}")
print(f"Всего строк данных: {len(df)}")
print(f"Колонки: {', '.join(df.columns)}")
