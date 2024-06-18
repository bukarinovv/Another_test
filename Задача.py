import pandas as pd

may = pd.read_excel('data.xlsx', usecols="B:E, G, H", header=0, skiprows=[1, 2], skipfooter=731-259)

may_bonus = {}

m_new = may[(may['new/current'] == 'новая') & (may['document'] == 'оригинал') & (may['status'] == 'ОПЛАЧЕНО')]
for summa, name, date in zip(m_new['sum'], m_new['sale'], m_new['receiving_date']):
    if date.month in (5, 6):
        may_bonus[name] = may_bonus.get(name, 0) + summa * 0.07

m_cur = may[(may['new/current'] == 'текущая') & (may['document'] == 'оригинал') & (may['status'] != 'ПРОСРОЧЕНО')]
for summa, name, date in zip(m_cur['sum'], m_cur['sale'], m_cur['receiving_date']):
    if date.month in (5, 6):
        if summa > 10000:
            may_bonus[name] = may_bonus.get(name, 0) + summa * 0.05
        else:
            may_bonus[name] = may_bonus.get(name, 0) + summa * 0.03

for name, bonus in may_bonus.items():
    print(f'{name}: {round(bonus, 2)} руб.')