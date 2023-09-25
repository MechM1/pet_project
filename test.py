import openpyxl
import pandas as pd

Sa = float(input("Введите ставку аренды: "))
APrPl = float(input("Введите арендо-пригодную площадь: "))
Poteru = float(input("Введите потери: "))
Dop_dox = float(input("Введите дополнительный доход: "))
operating_doxod = float(input("Введите операционный доход: "))
step = int(input("Введите шаг расчета или период: "))
zatratu = float(input("Введите затраты за период: "))
Kt = float(input("Введите капитальные вложения: "))
T = int(input("Введите горизонт расчета: "))

df = pd.read_excel('/content/data2.xlsx')
discount_rate = df.iloc[0,0]

PVD = Sa * APrPl
DVD = PVD - Poteru + Dop_dox
CHOD = DVD - operating_doxod

cash_flows = [-Kt]
discounted_cash_flows = [cash_flows[0] / (1 + discount_rate)]
for i in range(1, T+1, step):
    Rt = CHOD
    zt = zatratu + Kt
    cash_flow = Rt - zt
    cash_flows.append(cash_flow)
    discounted_cash_flow = cash_flow / ((1 + discount_rate) ** i)
    discounted_cash_flows.append(discounted_cash_flow + discounted_cash_flows[-1])

factors = ['Общие экономические тенденции', 'Внешняя экономическая деятельность', 'Инфляция', 'Инвестиция', 'Доходы и сбережения населения', 'Система налогооблажения', 'Угроза передела собственности', 'Внутреполитическая стабильность', 'Внешняя политическая деятельность', 'Угроза террористических актов', 'Социальная стабильность в стране', 'Социальная обеспечённость граждан', 'Тенденция развития экономики в регионе', 'Социальная стабильность в регионе', 'Ликвидность', 'Уровень конкуренции в отрасли', 'Инвестиционная привлекательность проекта', 'Тенденция развития отрасли', 'Природные чрезвычайные ситуации', 'Антропогенные чрезвычайные ситуации', 'Ускоренный износ объекта', 'Экологические факторы', 'Криминогенные факторы', 'Финансовые проверки', 'Фактор изменения окружающей застройки', 'Увеличение числа конкурирующих объектов', 'Изменение общей экономической ситуации']
ratings = {}
for factor in factors:
    while True: 
        rating = input(f'Введите оценку для {factor} от 0 до 10: ')
        if rating and rating.isdigit() and int(rating) >= 0 and int(rating) <= 10:
            ratings[factor] = int(rating)
            break
        elif not rating:
            break

# записываем оценки в файл Excel
workbook = openpyxl.load_workbook('/content/data2.xlsx')
worksheet = workbook.active
for factor, rating in ratings.items():
    column = factors.index(factor) + 1
    row = rating + 1
    worksheet.cell(row=1, column=1, value=1)
workbook.save('/content/data2.xlsx')

workbook.save('/content/data2.xlsx')

# Рассчитывае1м В1Н

CHDD1 = sum([(Rt - 3*t) / ((1 + discount_rate)**t) for t, Rt in enumerate(cash_flows)])
CHDD2 = sum([(Rt - 3*t) / ((1 + 100)**t) for t, Rt in enumerate(cash_flows)])
E1 = discount_rate
E2 = 100
VND = E1 + (CHDD1 / (CHDD1 - CHDD2)) * (E2 - E1)

print("Дисконтированные потоки денежных средств: ", discounted_cash_flows)
print("Чистая приведенная стоимость: ", round(VND, 2))