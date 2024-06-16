import pandas as pd

# Чтение файла Excel

# 1 Вопрос
df = pd.read_excel('data.xlsx', usecols="A:E,G,H", header=0, skiprows=list(range(1, 260)), skipfooter=731-370)
filtered_df = df[df['status'] != 'ПРОСРОЧЕНО']
print(f'{round(filtered_df['sum'].sum(), 2)} руб.')

# 2 Вопрос
