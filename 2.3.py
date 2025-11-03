import pandas as pd
import numpy as np
import matplotlib.pyplot as plt 
from scipy.stats import spearmanr
from sklearn.linear_model import LinearRegression
import seaborn as sns

# Загрузка данных
dan = pd.read_excel(r"C:\Users\sanya\Desktop\UGNTU\Pytnoh\danie.xlsx", header=None)

# Формируем заголовки из первых 4 строк
dan.columns = (
    dan.iloc[0:4]
    .fillna('')
    .astype(str)
    .agg('_'.join, axis=0)
    .str.replace(r'\s+', ' ', regex=True)
    .str.replace('\n', ' ') 
    .str.strip('_')
)

# Обрабатываем дубликаты в названиях столбцов
seen = {}
new_columns = []
for col in dan.columns:
    if col in seen:
        seen[col] += 1
        new_columns.append(f"{col}_{seen[col]}")
    else:
        seen[col] = 0
        new_columns.append(col)

dan.columns = new_columns

dan = dan.iloc[5:].reset_index(drop=True)


# Преобразуем к числовому типу (некорректные заменяются на NaN)
dan = dan.apply(pd.to_numeric, errors='coerce')

# 1. Фильтрация столбцов: удаляем, если больше половины значений - не числа
filtered_columns = []
for col in dan.columns:
    kol_znac = dan[col].count()  # Количество числовых значений
    if kol_znac > (len(dan) / 2):  # Если числовых значений больше половины
        filtered_columns.append(col)

dan = dan[filtered_columns]

# 2. Сортировка по "Текущая наработка"
dan = dan.sort_values(by="Текущая наработка").reset_index(drop=True)

# 3. Поиск выбросов методом 3 сигм и удаление строк с выбросами
for col in dan.columns:
    if col == "Текущая наработка":
        continue
    
    mean = dan[col].mean()
    std = dan[col].std()
    
    # Границы нормальных значений
    lower_bound = mean - 3 * std
    upper_bound = mean + 3 * std

    # Фильтруем строки, где текущее значение в пределах 3 сигм
    dan = dan[(dan[col] >= lower_bound) & (dan[col] <= upper_bound)]

#  4. Сбрасываем индексы после удаления выбросов
dan = dan.reset_index(drop=True)

# 5. Замена некорректных значений случайными числами между соседними
for col in dan.columns:
    if col == "Текущая наработка":
        continue

    for i in range(len(dan)):
        if pd.isna(dan.at[i, col]):  # Если значение некорректное
            prev_value = dan[col].iloc[i - 1] if i > 0 else dan[col].dropna().iloc[0]
            next_value = dan[col].iloc[i + 1] if i < len(dan) - 1 else dan[col].dropna().iloc[-1]

            # Генерируем случайное число между соседними значениями
            dan.at[i, col] = np.random.uniform(prev_value, next_value)

# 5. Расчет корреляции Спирмена
narabotka = dan["Текущая наработка"]
correl = {}

for factor in dan.columns:
    if factor == "Текущая наработка":
        continue

    x = dan[factor]
    y = narabotka

    if x.count() < 3:  # Пропускаем столбцы с недостаточным числом данных
        continue

    correlll, _ = spearmanr(x, y, nan_policy='omit')

    if pd.notna(correlll) and abs(correlll) >= 0:  
        correl[factor] = correlll

# Сортируем факторы по абсолютному значению корреляции
factorcorrel = sorted(correl.items(), key=lambda x: abs(x[1]), reverse=True)

# Оставляем 35 случайных факторов (не обязательно самых значимых)
random_factors = np.random.choice(list(correl.keys()), size=35, replace=False)

# Создаем DataFrame для случайных факторов
correlation_df_random = pd.DataFrame({factor: correl.get(factor, np.nan) for factor in random_factors}, index=['Корреляция'])

# Отображаем тепловую карту для выбранных 35 случайных факторов
plt.figure(figsize=(14, 10))  # Уменьшаем высоту, увеличиваем ширину
sns.heatmap(
    correlation_df_random.T, annot=True, cmap='coolwarm', fmt=".2f", 
    annot_kws={"size": 10}, linewidths=1, cbar_kws={"shrink": 0.6}, square=False
)
plt.title("Корреляция факторов с наработкой", fontsize=14)
plt.savefig(r"C:\Users\sanya\Desktop\UGNTU\Pytnoh\teplovoi grafik.png", dpi=300, bbox_inches='tight')  # Сохранение тепловой карты
plt.show()

# 6. Построение модели линейной регрессии для выбранных факторов
vazn_fact = random_factors  # Используем случайные 35 факторов
X = dan[vazn_fact].dropna()
y = narabotka.loc[X.index]

if not X.empty:
    model = LinearRegression()
    model.fit(X, y)
    
    print("\nУравнение регрессии:")
    fac_cof = [f"({coef:.3f} * {name})" for coef, name in zip(model.coef_, vazn_fact)]
    urav = "Наработка = \n" + " +\n".join(fac_cof) + f" +\n({model.intercept_:.3f})"
    #print(urav)
else:
    print("\nНедостаточно данных для построения модели.")

# Определяем нужные факторы
selected_factors = ["Обводненность", "Сод-е мехпримесей (КВЧ)", "В-ть жидкости", 
                    "ГФ", "Р заб (расчетное)", "P затр"]

# Графики зависимости факторов
plt.figure(figsize=(15, 10))

for i, factor in enumerate(selected_factors):
    if factor in dan.columns:
        plt.subplot(2, 3, i + 1)
        
        x = dan[factor]
        y = narabotka

        plt.scatter(x, y, label="Данные")  # Основные точки

        # Добавляем линию тренда (линейную регрессию)
        coef = np.polyfit(x, y, 1)  # Аппроксимация 1-й степени (прямая линия)
        trend = np.poly1d(coef)
        plt.plot(x, trend(x), color="red", linestyle="--", label="Линия регресии")

        plt.xlabel(factor)
        plt.ylabel("Наработка на отказ")
        plt.title(f"Зависимость наработки от {factor}")
        plt.legend()

plt.tight_layout()
plt.savefig(r"C:\Users\sanya\Desktop\UGNTU\Pytnoh\gfrafik.png", dpi=300, bbox_inches='tight') # Сохранение графиков зависимости
plt.show()


'''
# Определяем факторы для построения графиков
selected_factors = ["Обводненность", "Сод-е мехпримесей (КВЧ)", "В-ть жидкости", 
                    "ГФ", "Р заб (расчетное)", "P затр"]

plt.figure(figsize=(15, 10))  # Размер всего изображения

for i, factor in enumerate(selected_factors):
    if factor in dan.columns:
        plt.subplot(2, 3, i + 1)  # 2 строки, 3 столбца
        x = dan[factor]
        y = dan["Текущая наработка"]

        plt.scatter(x, y, label="Данные", alpha=0.7)  # Точки данных

        # Линия регрессии
        regression_line = model.coef_[selected_factors.index(factor)] * x + model.intercept_
        plt.plot(x, regression_line, color="red", linestyle="--", label="Линия регрессии")

        plt.xlabel(factor)
        plt.ylabel("Текущая наработка")
        plt.title(f"Регрессия: {factor}")
        plt.legend()
        plt.grid()

plt.tight_layout()

# Сохраняем картинку
plt.savefig("regression_factors.png", dpi=300, bbox_inches="tight")  # Сохранение файла
plt.show()
'''
