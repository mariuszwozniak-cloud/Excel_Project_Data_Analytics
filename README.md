# Excel_Project_Data_Analytics
 My project demonstrating my Excel skills 


# 🎯 Cel analizy

Celem niniejszego raportu jest identyfikacja i ocena **czynników wpływających na występowanie cukrzycy** typu 2 na podstawie dostarczonego zbioru danych. Analiza pozwala określić, które cechy zdrowotne, demograficzne i stylu życia mają największe znaczenie w przewidywaniu diagnozy cukrzycy (zmienna `Outcome`).


# 🔍 Eksploracyjna analiza danych (EDA)

Zbiór zawiera **9539 wierszy** (rekordów) i **18 kolumn** (cech).

Dane pochodzą z arkusza `diabetes_dataset (4)` i zawierają informacje o:

- parametrach medycznych (np. poziom glukozy, HbA1c, BMI, ciśnienie krwi),
- historii rodziny (`FamilyHistory`),
- cechach demograficznych (wiek, liczba ciąż),
- stylu życia (typ diety, stosowanie leków, obecność nadciśnienia).

Zmienna docelowa `Outcome` informuje, czy dana osoba została zdiagnozowana z cukrzycą (1) czy nie (0).

**Wstępne przetwarzanie i przegląd danych**

- Wczytano dane z pliku `dietetycy vol1.xlsx`.
- Przegląd podstawowych statystyk opisowych (mean, median, std, min, max) dla zmiennych liczbowych.
    
    ---
    

## **Analiza korelacji - Heatmapa** 🌡️📈🧠

- Stworzono mapę korelacji zmiennych, która pozwoliła zidentyfikować zależności między cechami
- Zidentyfikowano wysoką korelację pomiędzy BMI, poziomem glukozy i ciśnieniem tętniczym, a także bardzo wysoką korelację (0,91) pomiędzy historią rodzinną choroby a jej występowaniem

![Wykres 1. Heatmap - Mapa korelacji](images/Figure_1_heatmap.png)

- **Kod użyty do wygenerowania heatmapy:**

```python
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

file_path = "S:/Project Excel 1/Excel_Project_Data_Analytics/dietetycy vol1.xlsx"
df_main = pd.read_excel(file_path, sheet_name="diabetes_dataset (4)")

# Korelacje
corr_matrix = df_main.corr(numeric_only=True)

# Wizualizacja
plt.figure(figsize=(12, 10))
sns.heatmap(corr_matrix, annot=True, fmt=".2f", cmap="coolwarm", square=True)
plt.title("🔥 Mapa korelacji zmiennych 🔥")
plt.tight_layout()
plt.show()
```

## **Wizualizacje danych** 🎨📊🔍

- **Boxploty** (wykresów pudełkowych) dla zmiennych: BMI, BloodPressure, LDL, HDL pokazujące wartości odstające (outliery) wg reguły 1.5x IQR. (Stworzone w MS Excel)
    
    ![image.png](images/image.png)
    
    Wartości odstające częściej występują po prawej stronie – rozkłady są **prawoskośne** (więcej osób z niskimi wartościami, kilku z bardzo wysokimi). 
    
    Mediana glukozy, ciśnienia, poziomu LDL i HDL – można łatwo odczytać i porównać między cechami.
    
    ---
    
- **Histogramy** zmiennych liczbowych dla wstępnego przeglądu rozkładu.
    
    ![Wykres 3. Histogramy](images/Figure_2_Histogramy.png)
    

---

- Historia rodzinna a cukrzyca:
    
    ![image.png](images/image.png)
    
    - Wszystkie osoby z `FamilyHistory = 1` miały cukrzycę (`Outcome = 1`)
    - Korelacja z `Outcome` wynosi **0.91**, co świadczy o **bardzo silnym związku**
    - To najważniejszy predyktor w zbiorze — historia rodzinna niemal zawsze wiąże się z chorobą

---

- Korelacje cech bez historii chorobowej
    
    ![figure_3 heatmap bez familyhistory.png](images/figure_3_heatmap_bez_familyhistory.png)
    
    ✅ **Wniosek:** Gdy brak obciążeń genetycznych, cukrzyca jest silniej związana z:
    
    - **parametrami metabolicznymi** (glukoza ~0.50, HbA1c ~0.46)
    - **otyłością** (BMI~0,29, talia 0,22)
    - **ciśnieniem krwi (0,27)**

---

- Dashboard w MS Excel pokazujący zależność BMI od kluczowych parametrów metabolicznych wpływających na rozwój cukrzycy.
    
    ![image.png](images/image.png)
    
    BMI **nie powoduje bezpośrednio cukrzycy**, ale:
    
    - Podnosi poziom **glukozy**, **insulinooporność** i **HbA1c**
    - Te parametry **są bezpośrednio powiązane z cukrzycą**

---

# 🏁Podsumowanie najważniejszych czynników ryzyka

**Historia cukrzycy w rodzinie to najsilniejszy czynnik ryzyka i powinna być brana pod uwagę w pierwszej kolejności przy ocenie zagrożenia cukrzycą.**

### Pozostałe istotne wskaźniki

### 1. Glukoza we krwi 📈

- Najsilniejszy metaboliczny wskaźnik cukrzycy w tej grupie.
- Korelacja z `Outcome`: **0.50**
- Wysokie wartości glukozy znacząco zwiększają ryzyko.

### 2. Hemoglobina glikowana (HbA1c) 🩸

- Korelacja: **0.46** — silna, bardzo zbliżona do glukozy.
- Odzwierciedla średni poziom glukozy z dłuższego okresu.

### 3. Wskaźnik masy ciała (BMI) ⚖️

- Korelacja: **0.29**
- Otyłość zwiększa ryzyko zachorowania, choć słabiej niż parametry glukozy.
- Można mieć wysokie BMI przy dobrych parametrach metabolicznych (tzw. „metabolicznie zdrowa otyłość")
- Część osób z cukrzycą ma **prawidłowe BMI**, lecz inne czynniki ryzyka, jak predyspozycje genetyczne lub niewłaściwa dieta - To osłabia bezpośrednią korelację BMI z cukrzycą
- BMI może mieć **nieliniowy wpływ** — ryzyko wzrasta znacząco dopiero powyżej wartości 30
- Korelacja Pearsona **nie wykrywa** takich nieliniowych zależności stąd niska korelacja  - 0,29

### 4. Ciśnienie tętnicze 💉

- Korelacja: **0.27**
- Umiarkowana zależność — może współdziałać z innymi czynnikami.

---

# **Zakończenie** ✅

Analiza danych wykazała, że **występowanie cukrzycy (Outcome = 1)** jest wynikiem złożonej interakcji czynników genetycznych i metabolicznych.

 Najsilniejszym pojedynczym predyktorem była **historia rodzinna cukrzycy**, która miała bezpośredni i bardzo wyraźny wpływ – wszystkie osoby z `FamilyHistory = 1` miały cukrzycę. To podkreśla **dominującą rolę czynników genetycznych**.

Wśród osób **bez historii rodzinnej**, kluczowe znaczenie mają parametry metaboliczne:

- **Glukoza** i **HbA1c** – najbardziej powiązane z cukrzycą
- **BMI** i **ciśnienie krwi** – umiarkowane, ale istotne predyktory

Warto zauważyć, że **BMI wpływa pośrednio**, poprzez pogarszanie parametrów takich jak poziom glukoza czy insulinooporność, ale samo w sobie **nie jest silnie skorelowane z cukrzycą**.