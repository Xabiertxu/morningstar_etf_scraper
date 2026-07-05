import os
import requests
import pandas as pd
from datetime import datetime

# URL regionalizada de la API de Morningstar (Evita fallos de resolución de DNS)
API_URL = "https://ie2api.morningstar.com/api/v2/search/v1/solr"

# Parámetros necesarios para extraer la lista completa con los campos de las 4 secciones
params = {
    "q": "*:*",
    "fq": "AssetClass:\"ETF\" AND FundShareClassLocaleId:\"OSESP$$$$$\"", 
    "wt": "json",
    "start": 0,
    "rows": 5000, 
    "fl": "Id,Name,LegalName,Ticker,ISIN,SecId,Universe,ClosePrice,Currency,ExchangeName,StandardDeviationThreeYear,SharpeRatioThreeYear,ReturnM1,ReturnM3,ReturnM12,Return3YrAvg,SustainabilityRating,CarbonScore"
}

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "application/json"
}

print(f"[{datetime.now()}] Conectando con la API Regional de Morningstar...")

try:
    response = requests.get(API_URL, params=params, headers=headers)
    response.raise_for_status()
    data = response.json()
    
    docs = data.get("response", {}).get("docs", [])
    total_found = data.get("response", {}).get("numFound", 0)
    print(f"Se han encontrado un total de {total_found} ETFs en la base de datos.")

    if not docs:
        print("No se pudieron extraer registros. Comprueba los filtros.")
        exit(1)

    # Procesar y mapear las 4 secciones solicitadas
    lista_etfs = []
    for doc in docs:
        etf = {
            # --- SECCIÓN 1: GENERAL ---
            "ID Morningstar": doc.get("Id"),
            "Nombre del ETF": doc.get("Name"),
            "Ticker": doc.get("Ticker"),
            "ISIN": doc.get("ISIN"),
            "Último Precio": doc.get("ClosePrice"),
            "Divisa": doc.get("Currency"),
            "Bolsa de Cotización": doc.get("ExchangeName"),
            
            # --- SECCIÓN 2: RENTABILIDAD ---
            "Rentabilidad 1 Mes (%)": doc.get("ReturnM1"),
            "Rentabilidad 3 Meses (%)": doc.get("ReturnM3"),
            "Rentabilidad 12 Meses (%)": doc.get("ReturnM12"),
            "Rentabilidad Anualizada 3 Años (%)": doc.get("Return3YrAvg"),
            
            # --- SECCIÓN 3: RIESGO ---
            "Desviación Estándar (3y)": doc.get("StandardDeviationThreeYear"),
            "Ratio de Sharpe (3y)": doc.get("SharpeRatioThreeYear"),
            
            # --- SECCIÓN 4: SOSTENIBILIDAD ---
            "Globos de Sostenibilidad (1-5)": doc.get("SustainabilityRating"),
            "Puntuación de Carbono": doc.get("CarbonScore")
        }
        lista_etfs.append(etf)

    # Convertir a DataFrame de Pandas
    df_nuevo = pd.DataFrame(lista_etfs)
    archivo_salida = "lista_completa_etfs_morningstar.xlsx"
    
    # Consolidar con datos anteriores si existen, evitando duplicados
    if os.path.exists(archivo_salida):
        try:
            df_existente = pd.read_excel(archivo_salida)
            df_final = pd.concat([df_existente, df_nuevo], ignore_index=True)
            if "ID Morningstar" in df_final.columns:
                df_final = df_final.drop_duplicates(subset=["ID Morningstar"], keep="last")
        except Exception:
            df_final = df_nuevo
    else:
        df_final = df_nuevo

    # Guardar en Excel con formato limpio
    df_final.to_excel(archivo_salida, index=False)
    print(f"¡Éxito! Archivo '{archivo_salida}' generado con {len(df_final)} ETFs totales procesados.")

except Exception as e:
    print(f"Error durante la ejecución del scraping: {e}")
    exit(1)
