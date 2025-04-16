import pandas as pd
import numpy as np

# Wczytanie danych z drugiego arkusza (indeks 1)
df_wms = pd.read_excel(r"C:\Users\wiktor.daszynski\Downloads\Document1.xlsx", sheet_name=0, header=1)

# Usuwanie kolumn, których nagłówek kończy się na 'week' i nie zaczyna się od 'Brak'
columns_to_drop = [col for col in df_wms.columns if col.endswith('week') and not col.startswith('Brak')]
df_wms = df_wms.drop(columns=columns_to_drop)

# Usuwanie wierszy gdzie 'Ilość zamówiona' > 'Brak 8 week'
if 'Ilość zamówiona' in df_wms.columns and 'Brak 8 week' in df_wms.columns:
    df_wms = df_wms[df_wms['Ilość zamówiona'] <= df_wms['Brak 8 week']]

# Usuwanie wierszy gdzie 'Brak 13 week' < 1
if 'Brak 13 week' in df_wms.columns:
    df_wms = df_wms[df_wms['Brak 13 week'] >= 1]

# Sortowanie po kolumnie 'Brak 13 week' malejąco
if 'Brak 13 week' in df_wms.columns:
    df_wms = df_wms.sort_values(by='Brak 13 week', ascending=False)

# Nazwa kolumny z kodem towaru
column_name = 'Kod towaru'  # Zmień na dokładną nazwę kolumny

if column_name in df_wms.columns:
    # Grupowanie po 'Kod towaru' i sumowanie wartości liczbowych
    df_summed = df_wms.groupby(column_name).sum(numeric_only=True).reset_index()

    # Sortowanie po kolumnie 'Brak 13 week' malejąco po sumowaniu
    if 'Brak 13 week' in df_summed.columns:
        df_summed = df_summed.sort_values(by='Brak 13 week', ascending=False)

    print("Zsumowane wartości dla nieunikatowych rekordów w 'Kod towaru':")
    print(df_summed)

    # Zapisz wynik do pliku Excel
    df_summed.to_excel("wynik_zsumowany.xlsx", index=False)

    # Sprawdzenie unikatowości wartości w kolumnie 'Kod towaru' po sumowaniu
    num_unique = df_summed[column_name].nunique()
    total = len(df_summed[column_name])
    if num_unique == total:
        print("Po sumowaniu wszystkie wartości w 'Kod towaru' są unikatowe.")
    else:
        print("Po sumowaniu nadal występują duplikaty w 'Kod towaru'.")
else:
    print(f"Kolumna '{column_name}' nie istnieje w DataFrame.")