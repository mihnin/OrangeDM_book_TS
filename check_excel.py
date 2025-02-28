import pandas as pd
import os

def check_excel_columns():
    """Проверяет наличие столбцов 'ЧОК' и 'Запасы' в Excel-файле"""
    file_path = os.path.join(r"C:\OrangeDM_book_TS", "financial_dataset.xlsx")
    
    if not os.path.exists(file_path):
        print(f"Файл {file_path} не найден!")
        return
    
    print(f"Чтение файла {file_path}...")
    df = pd.read_excel(file_path)
    
    print(f"\nСписок всех столбцов в Excel-файле:")
    for i, col in enumerate(df.columns, 1):
        print(f"{i}. {col}")
    
    print("\nПроверка наличия нужных столбцов:")
    if 'Запасы' in df.columns:
        print(f"Столбец 'Запасы' найден. Первые 5 значений: {df['Запасы'].head().tolist()}")
    else:
        print("Столбец 'Запасы' НЕ найден!")
    
    if 'ЧОК' in df.columns:
        print(f"Столбец 'ЧОК' найден. Первые 5 значений: {df['ЧОК'].head().tolist()}")
    else:
        print("Столбец 'ЧОК' НЕ найден!")

if __name__ == "__main__":
    check_excel_columns()
