import pandas as pd
import matplotlib.pyplot as plt
import calendar, locale

# 1 Вопрос
df = pd.read_excel('data.xlsx', usecols="B, C", header=0, skiprows=list(range(1, 260)), skipfooter=731-370)
filtered_df = df[df['status'] != 'ПРОСРОЧЕНО']
print(f'Первый вопрос: {round(filtered_df['sum'].sum(), 2)} руб.')


# 2 Вопрос
df = pd.read_excel('data.xlsx', usecols="B, H", header=0, skiprows=[1, 2])
data = {}
for mon, summa in zip(df['receiving_date'], df['sum']):
    data[mon.month] = data.get(mon.month, 0) + summa

data = dict(sorted([elem for elem in data.items() if isinstance(elem[0], int)]))

locale.setlocale(locale.LC_ALL, 'ru_RU.UTF-8')

months = list(map(lambda x: calendar.month_name[x], data.keys()))

plt.plot(months, data.values(), c = 'r')
plt.grid(True)
plt.title('Выручка по месяцам, руб.')
plt.show()



# 3 Вопрос
df = pd.read_excel('data.xlsx', usecols="B, D", header=0, skiprows=list(range(1, 486)), skipfooter=731-595)
group = df.groupby('sale').sum()
sorted_group = group.sort_values(by='sum', ascending=False)
print(f'Третий вопрос: {sorted_group.iloc[0].name}')

