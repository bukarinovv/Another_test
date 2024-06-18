import pandas as pd

may = pd.read_excel('data.xlsx', usecols="B:E, G, H", header=0, skiprows=[1, 2], skipfooter=731-259)

bonus_list = {}

m_new = may[(may['document'] == 'оригинал') & (may['status'] != 'ПРОСРОЧЕНО')]
for deal, summa, name, date, status in zip(m_new['new/current'], m_new['sum'], m_new['sale'], m_new['receiving_date'], m_new['status']):
    if date.month in (5, 6):
        if deal == 'новая' and status == 'ОПЛАЧЕНО':
            bonus = 0.07
        elif deal == 'текущая':
            bonus = 0.05 if summa > 10000 else 0.03
        bonus_list[name] = bonus_list.get(name, 0) + summa * bonus


for name, bonus in bonus_list.items():
    print(f'{name}: {round(bonus, 2)} руб.')