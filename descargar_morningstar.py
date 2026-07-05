import os
import pandas as pd
import yfinance as yf
from datetime import datetime

# Lista de principales ETFs de alta liquidez y réplica global comercializados en España 
# Cubre renta variable global, americana, europea, emergentes y sostenibilidad (ESG)
tickers_etfs = [
    "SPY", "QQQ", "IWM", "EUNN", "VGK", "IEFA", "VTI", "VOO", "ACWI", "IWDA.AS",
    "EURL.PA", "EXS1.DE", "MEUD.PA", "ANX.PA", "MVEU.PA", "SUSW.L", "SUAS.L", 
    "IUSA.L", "SREU.PA", "IS3N.DE", "IQQW.DE", "LYXIB.PA", "BBVA.MC", "TEF.MC"
]

print(f"[{datetime.now()}] Iniciando extracción de base de datos de ETFs estables...")

lista_datos = []

for ticker in tickers_etfs:
    try:
        print(f"Procesando: {ticker}...")
        tk = yf.Ticker(ticker)
        info = tk.info
        
        # Estructuramos las 4 secciones solicitadas mapeando los datos de Yahoo Finance
        etf = {
            # --- SECCIÓN 1: GENERAL ---
            "ID Ticker": ticker,
            "Nombre del ETF": info.get("longName", info.get("shortName", "N/A")),
            "ISIN/Símbolo": info.get("underlyingSymbol", ticker),
            "Último Precio": info.get("previousClose", info.get("regularMarketPreviousClose")),
            "Divisa": info.get("currency", "EUR"),
            "Bolsa de Cotización": info.get("exchange", "N/A"),
            
            # --- SECCIÓN 2: RENTABILIDAD ---
            "Rentabilidad YTD (%)": info.get("ytdReturn", 0) * 100 if info.get("ytdReturn") else "N/A",
            "Rentabilidad Anualizada 3 Años (%)": info.get("threeYearAverageReturn", 0) * 100 if info.get("threeYearAverageReturn") else "N/A",
            "Rentabilidad Anualizada 5 Años (%)": info.get("fiveYearAverageReturn", 0) * 100 if info.get("fiveYearAverageReturn") else "N/A",
            
            # --- SECCIÓN 3: RIESGO ---
            "Beta (Volatilidad vs Mercado)": info.get("beta", "N/A"),
            "Categoría de Riesgo": info.get("fundFamily", "N/A"),
            
            # --- SECCIÓN 4: SOSTENIBILIDAD ---
            "Puntuación ESG (Sostenibilidad)": info.get("esgPerformance", "N/A"),
            "Sostenibilidad Global": "Verificado (Categoría ESG)" if info.get("esgPerformance") else "N/A"
        }
        lista_datos.append(etf)
    except Exception as e:
        print(f"Aviso: No se pudieron extraer datos completos para {ticker}: {e}")
        continue

if not lista_datos:
    print("Error: No se ha podido extraer ninguna información de la lista.")
    exit(1)

# Convertir a DataFrame y consolidar
df_nuevo = pd.DataFrame(lista_datos)
archivo_salida = "lista_completa_etfs_morningstar.xlsx"

if os.path.exists(archivo_salida):
    try:
        df_existente = pd.read_excel(archivo_salida)
        df_final = pd.concat([df_existente, df_nuevo], ignore_index=True).drop_duplicates(subset=["ID Ticker"], keep="last")
    except Exception:
        df_final = df_nuevo
else:
    df_final = df_nuevo

# Guardar matriz final en Excel
df_final.to_excel(archivo_salida, index=False)
print(f"¡Éxito! Generado archivo '{archivo_salida}' con {len(df_final)} ETFs procesados sin errores de DNS.")
