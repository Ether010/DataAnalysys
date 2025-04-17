import pandas as pd
import numpy as np

def prepare_tbo_dataframe(input_path, output_path):
    df_wms = pd.read_excel(input_path, sheet_name=0, header=1)
    columns_to_drop = [col for col in df_wms.columns if col.endswith('week') and not col.startswith('Brak')]
    df_wms = df_wms.drop(columns=columns_to_drop)
    if 'Ilość zamówiona' in df_wms.columns and 'Brak 8 week' in df_wms.columns:
        df_wms = df_wms[df_wms['Ilość zamówiona'] <= df_wms['Brak 8 week']]
    if 'Brak 13 week' in df_wms.columns:
        df_wms = df_wms[df_wms['Brak 13 week'] >= 1]
        df_wms = df_wms.sort_values(by='Brak 13 week', ascending=False)
    column_name = 'Kod towaru'
    if column_name in df_wms.columns:
        df_summed = df_wms.groupby(column_name).sum(numeric_only=True).reset_index()
        if 'Brak 13 week' in df_summed.columns:
            df_summed = df_summed.sort_values(by='Brak 13 week', ascending=False)
        df_summed.to_excel(output_path, index=False)
        return df_summed
    else:
        raise ValueError(f"Kolumna '{column_name}' nie istnieje w DataFrame.")

def calculate_tbo_statistics(df):
    # Przykład: oblicz średnią i nachylenie wzrostu dla 'Brak 13 week'
    if 'Brak 13 week' in df.columns:
        mean_value = df['Brak 13 week'].mean()
        x = np.arange(len(df))
        y = df['Brak 13 week'].values
        slope, _ = np.polyfit(x, y, 1)
        print(f"Średnia: {mean_value}, Nachylenie wzrostu: {slope}")
    else:
        print("Brak kolumny 'Brak 13 week' w przekazanym DataFrame.")

if __name__ == "__main__":
    # Tworzenie i obrabianie pliku
    df_tbo = prepare_tbo_dataframe(
        r"C:\Users\wiktor.daszynski\Downloads\Document1.xlsx",
        "wynik_zsumowany.xlsx"
    )
    # Statystyki
    calculate_tbo_statistics(df_tbo)