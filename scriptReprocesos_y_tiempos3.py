import pandas as pd
import numpy as np

import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')


df = pd.read_excel(r"C:\Users\jcarchila\Downloads\tiempos_incidentes.xlsx", parse_dates=['Start time', 'End Time'])

# Asignar orden por fecha dentro de cada incidente
def asignar_orden_por_incidente(df):
    # Ordenar los datos por referencia e inicio de tiempo
    df = df.sort_values(by=['Referencia', 'Start time']).reset_index(drop=True)
    # Asignar orden de manera secuencial dentro de cada incidente
    df['Orden'] = df.groupby('Referencia').cumcount() + 1
    return df

# Calcular reprocesos por incidente
def calcular_reprocesos(df):
    df = df.sort_values(by=['Referencia', 'Orden']).reset_index(drop=True)
    df['Reproceso'] = 0
    df['Rechazo'] = 0
    df['Agendado'] = 0

    for i in range(1, len(df)):
        if df.iloc[i]['Referencia'] == df.iloc[i - 1]['Referencia']:            #Si el CM es el mismo que el anterior
            if df.iloc[i]['Etapas CM'] == 'Originación' and df.iloc[i - 1]['Etapas CM'] != 'Originación':       #Si la etapa es Originacion y viene de una distinta a Originacion
                df.at[i, 'Reproceso'] += 1      #Le agrega 1 reproceso
                df.at[i - 1, 'Rechazo'] += 1      #Le agrega 1 reproceso
            elif df.iloc[i]['Etapas CM'] == '2. Soporte Crediticio' and df.iloc[i - 1]['Etapas CM'] != 'Originación':       #Si es soporte Crediti
                df.at[i, 'Reproceso'] += 1
            elif df.iloc[i]['Etapas CM'] in ('Aprobación','4.1 Aprobación por Facultamiento') and df.iloc[i - 1]['Etapas CM'] in ('3.1 Empresarial Menor','3.2 Empresarial Mayor'):
                df.at[i-1, 'Agendado'] += 1


    return df

# Parte adicional nueva 
# Calcular duración laboral considerando días y horas hábiles


# Lista de días festivos (días no laborables)
dias_festivos = [
    pd.Timestamp('2025-05-01').date(), 
    pd.Timestamp('2025-04-16').date(), 
    pd.Timestamp('2025-05-17').date(), 
    pd.Timestamp('2025-05-18').date()
    ]

def calcular_duracion_laboral(df, dias_festivos):
    try:
        df['Start time'] = pd.to_datetime(df['Start time'], errors='coerce')
        df['End Time'] = pd.to_datetime(df['End Time'], errors='coerce')

        df['Duración (horas laborales)'] = 0.0
        df['Duración (días laborales)'] = 0.0
        df['Duración (horas sobrantes)'] = 0.0
        df['Duración Formateada'] = "0 días, 0 horas"

        def calcular_horas_laborales(start, end):
            if pd.isna(start) or pd.isna(end) or start > end:
                return 0.0

            rango_dias = pd.date_range(start=start.date(), end=end.date(), freq='D')
            total_horas = 0.0

            for dia in rango_dias:
                if dia.weekday() >= 5 or dia.date() in dias_festivos:
                    continue

                inicio_dia = max(start, pd.Timestamp(dia.replace(hour=0, minute=0, second=0)))
                fin_dia = min(end, pd.Timestamp(dia.replace(hour=23, minute=59, second=0)))

                if inicio_dia < fin_dia:
                    horas = (fin_dia - inicio_dia).total_seconds() / 3600
                    total_horas += horas

            return total_horas

        df['Duración (horas laborales)'] = df.apply(
            lambda row: calcular_horas_laborales(row['Start time'], row['End Time']), axis=1
        )

        df['Duración (días laborales)'] = (df['Duración (horas laborales)'] / 24).round(2)
        df['Duración (horas sobrantes)'] = (df['Duración (horas laborales)'] % 9).astype(int)
        df['Duración Formateada'] = (
            df['Duración (días laborales)'].astype(str) + ' días, ' +
            df['Duración (horas sobrantes)'].astype(str) + ' horas'
        )

        return df

    except Exception as e:
        print(f"Error al calcular la duración laboral: {e}")
        raise



# Ejecutar la función
df = asignar_orden_por_incidente(df)

df1 = calcular_reprocesos(df)
# df1.to_excel("reprocesos3.xlsx", index=False)

df_resultante = calcular_duracion_laboral(df1, dias_festivos)

df_resultante.to_excel(r"C:\Users\jcarchila\Downloads\Reprocesos_y_Tiempos_TEST2.xlsx", index=False)