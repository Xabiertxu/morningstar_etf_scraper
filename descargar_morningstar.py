import os
import requests
import pandas as pd
from datetime import datetime

# URL de la API oculta del buscador de Morningstar España
API_URL = "https://data.morningstar.com/api/v2/search/v1/solr"

# Parámetros necesarios para extraer la lista completa con los campos de las 4 secciones
# (General, Rentabilidad, Riesgo, Sostenibilidad)
params = {
    "q": "*:*",
    "fq": "AssetClass:\"ETF\" AND FundShareClassLocaleId:\"OSESP$$$$$\"", # Filtra por ETFs disponibles en España
    "wt": "json",
    "start": 0,
    "rows": 5000, # Descarga hasta 5000 ETFs de golpe para asegurar la lista completa
    # Campos seleccionados que cubren las 4 pestañas solicitadas:
    "fl": "Id,Name,LegalName,Ticker,ISIN,SecId,Universe,ClosePrice,Currency,ExchangeName,StandardDeviationThreeYear,SharpeRatioThreeYear,ReturnM1,ReturnM3,ReturnM12,Return3YrAvg,SustainabilityRating,CarbonScore"
}

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
}

print(f"[{datetime.now()}] Conectando con la API de Morningstar...")

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
            "Rentabilidad Anualizada 3 Añós (%)": doc.get("Return3YrAvg"),
            
            # --- SECCIÓN 3: RIESGO ---
            "Desviación Estándar (3y)": doc.get("StandardDeviationThreeYear"),
            "Ratio de Sharpe (3y)": doc.get("SharpeRatioThreeYear"),
            
            # --- SECCIÓN 4: SOSTENIBILIDAD ---
            "Globos de Sostenibilidad (1-5)": doc.get("SustainabilityRating"),
            "Puntuación de Carbono": doc.get("CarbonScore")
        }
        lista_etfs.append(etf)

    # Convertir a DataFrame de Pandas
    df = pd.DataFrame(lista_etfs)
    
    # Nombre del archivo unificado
    archivo_salida = "lista_completa_etfs_morningstar.xlsx"
    
    # Guardar en Excel con formato limpio
    df.to_excel(archivo_salida, index=False)
    print(f"¡Éxito! Archivo '{archivo_salida}' generado con {len(df)} ETFs procesados.")

except Exception as e:
    print(f"Error durante la ejecución del scraping: {e}")
    exit(1)
