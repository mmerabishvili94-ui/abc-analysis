import pandas as pd
import random

products = [
    "Подшипник 6204-2RS", "Вал приводной ГП-450", "Сальник 25x52x7", 
    "Масло редукторное ISO 220", "Ремень клиновой SPA 1250",
    "Прокладка ГБЦ металлическая", "Фильтр масляный основной",
    "Цепь ГРМ усиленная", "Колодка тормозная передняя",
    "Диск тормозной вентилируемый", "Амортизатор газовый передний",
    "Пружина подвески задняя", "Свеча зажигания иридиевая",
    "Катушка зажигания тип B", "Датчик давления масла",
    "Термостат 87C", "Помпа водяная GMB", "Радиатор охлаждения",
    "Вентилятор принудительный", "Реле стартера 40А",
    "Предохранитель 15А", "Лампа H7 +50%", "Щетка стеклоочистителя 600мм",
    "Антифриз G12+ (1л)", "Тормозная жидкость DOT4 (0.5л)"
]

warehouses = ["Центральный склад", "Склад 'Север'", "Склад 'Юг'", "Минский филиал"]
units = ["шт", "комплект", "л", "кг"]

data = []
# Generate 50 rows of data to have some overlap
for _ in range(50):
    prod = random.choice(products)
    wh = random.choice(warehouses)
    unit = random.choice(units)
    # Most items have small consumption, some have large (for ABC)
    if random.random() < 0.1:
        qty = -random.randint(50, 200)
    else:
        qty = -random.randint(1, 15)
    
    # Store in a format compatible with Column C, D, E, F
    # We'll create a dataframe and then shift it to specific columns
    data.append({
        "A": "", # Placeholder col A
        "B": "", # Placeholder col B
        "Product": prod,  # Col C
        "Warehouse": wh, # Col D
        "Unit": unit,    # Col E
        "Qty": qty       # Col F
    })

df = pd.DataFrame(data)

# Rename to header names expected by the app search logic
df.columns = ["", "", "Товар", "Склад", "Ед. изм.", "Кол-во"]

# Create Excel writer
with pd.ExcelWriter('test_abc_analysis.xlsx', engine='xlsxwriter') as writer:
    df.to_excel(writer, index=False, sheet_name='Списания')
    
    # Adjust column widths for readability
    workbook = writer.book
    worksheet = writer.sheets['Списания']
    worksheet.set_column('C:C', 30)
    worksheet.set_column('D:D', 20)
    worksheet.set_column('E:E', 12)
    worksheet.set_column('F:F', 12)

print("Test file 'test_abc_analysis.xlsx' created.")
