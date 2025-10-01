import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# --- Configurazione della Pagina ---
st.set_page_config(
    page_title="Dashboard Consumi Energetici",
    page_icon="⚡",
    layout="wide"
)

# --- Caricamento e Pulizia Dati ---
@st.cache_data
def load_and_clean_data():
    # Carica il foglio "Consolidato" del file Excel
    df = pd.read_excel("Dati consumi e costi energetici.xlsx", sheet_name="Consolidato")

    # Rimuove le righe e colonne completamente vuote
    df = df.dropna(how='all').reset_index(drop=True)
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

    # Pulisce e converte le colonne numeriche
    numeric_cols = [
        'anno', 'mese', 'ore_produzione', 'pezzi_prodotti', 'consumo_kwh', 
        'lettura', 'costo_energia_per_kwh', 'costo_macchina', 'consumo_bolletta_kwh', 'totale_bolletta'
    ]
    for col in numeric_cols:
        if df[col].dtype == 'object':
            df[col] = df[col].astype(str).str.replace(' €', '', regex=False).str.replace(',', '.', regex=False)
        df[col] = pd.to_numeric(df[col], errors='coerce')

    # --- CALCOLO FORZATO PER RISOLVERE PROBLEMI DI VISUALIZZAZIONE ---
    # Questa operazione sovrascrive la colonna 'costo_macchina' letta dal file,
    # garantendo che sia sempre un valore numerico calcolato.
    df['costo_macchina'] = df['consumo_kwh'].fillna(0) * df['costo_energia_per_kwh'].fillna(0)

    # Crea una colonna data per facilitare i filtri
    df['data'] = pd.to_datetime(df['anno'].astype(str) + '-' + df['mese'].astype(str) + '-01', errors='coerce')

    # Ricalcola le metriche dipendenti con il nuovo costo_macchina
    df['consumo_per_pezzo'] = df['consumo_kwh'] / df['pezzi_prodotti'].replace(0, pd.NA)
    df['consumo_per_ora'] = df['consumo_kwh'] / df['ore_produzione'].replace(0, pd.NA)
    df['costo_per_pezzo'] = df['costo_macchina'] / df['pezzi_prodotti'].replace(0, pd.NA)

    return df

df = load_and_clean_data()

# --- Titolo e Intestazione ---
st.title("⚡ Dashboard Consumi Energetici - Vetronaviglio")
st.markdown("Un'applicazione interattiva per analizzare i consumi energetici delle macchine nello stabilimento.")

# --- Sidebar per i Filtri ---
st.sidebar.header("Filtri")

# Filtro per Macchina
macchine_univoche = df['macchina'].dropna().unique()
macchine_selezionate = st.sidebar.multiselect(
    "Seleziona Macchina/Impianto:",
    options=macchine_univoche,
    default=[m for m in macchine_univoche[:3] if m in macchine_univoche]  # Pre-seleziona le prime 3 valide
)

# Filtro per Anno
anni_univoci = sorted(df['anno'].dropna().unique())
anno_selezionato = st.sidebar.selectbox(
    "Seleziona Anno:",
    options=["Tutti"] + list(anni_univoci),
    index=0
)

# Filtro per Mese
mesi_univoci = sorted(df['mese'].dropna().unique())
mese_selezionato = st.sidebar.selectbox(
    "Seleziona Mese:",
    options=["Tutti"] + list(mesi_univoci),
    index=0
)

# Applica i filtri
df_filtrato = df.copy()

if macchine_selezionate:
    df_filtrato = df_filtrato[df_filtrato['macchina'].isin(macchine_selezionate)]

if anno_selezionato != "Tutti":
    df_filtrato = df_filtrato[df_filtrato['anno'] == anno_selezionato]

if mese_selezionato != "Tutti":
    df_filtrato = df_filtrato[df_filtrato['mese'] == mese_selezionato]

# --- Visualizzazione dei Dati Filtrati ---
st.subheader("📊 Dati Filtrati")

# Rimuovi le colonne 'lettura' e 'data' solo per la visualizzazione della tabella
df_display = df_filtrato.drop(columns=['lettura', 'data'], errors='ignore')

st.dataframe(df_display.style.format(
        formatter={
            "costo_macchina": lambda x: f'{x:,.2f} €' if pd.notna(x) else '-',
            "costo_energia_per_kwh": "{:,.4f} €",
            "totale_bolletta": "{:,.2f} €",
            "consumo_kwh": "{:,.2f}",
            "ore_produzione": "{:,.2f}",
            "pezzi_prodotti": "{:,.0f}"
        }
    ), use_container_width=True)

# --- Grafici Interattivi ---
st.subheader("📈 Analisi dei Consumi")

tab1, tab2, tab3, tab4 = st.tabs(["Consumo kWh", "Costo Macchina", "Consumo per Pezzo", "Consumo per Ora"])

with tab1:
    if not df_filtrato.empty:
        fig1 = px.line(
            df_filtrato,
            x='data',
            y='consumo_kwh',
            color='macchina',
            markers=True,
            title="Consumo Energetico (kWh) nel Tempo"
        )
        fig1.update_layout(xaxis_title="Data", yaxis_title="Consumo (kWh)")
        st.plotly_chart(fig1, use_container_width=True)
    else:
        st.info("Nessun dato disponibile con i filtri selezionati.")

with tab2:
    if not df_filtrato.empty:
        fig2 = px.bar(
            df_filtrato,
            x='macchina',
            y='costo_macchina',
            color='data',
            title="Costo Energetico per Macchina",
            barmode='group'
        )
        fig2.update_layout(xaxis_title="Macchina", yaxis_title="Costo (€)")
        st.plotly_chart(fig2, use_container_width=True)
    else:
        st.info("Nessun dato disponibile con i filtri selezionati.")

with tab3:
    if not df_filtrato.empty:
        fig3 = px.scatter(
            df_filtrato.dropna(subset=['consumo_per_pezzo']),
            x='pezzi_prodotti',
            y='consumo_per_pezzo',
            color='macchina',
            size='consumo_kwh',
            hover_data=['data', 'anno', 'mese'],
            title="Efficienza: Consumo per Pezzo Prodotto"
        )
        fig3.update_layout(xaxis_title="Pezzi Prodotti", yaxis_title="Consumo (kWh) per Pezzo")
        st.plotly_chart(fig3, use_container_width=True)
    else:
        st.info("Nessun dato disponibile con i filtri selezionati.")

with tab4:
    if not df_filtrato.empty:
        fig4 = px.box(
            df_filtrato.dropna(subset=['consumo_per_ora']),
            x='macchina',
            y='consumo_per_ora',
            title="Distribuzione del Consumo per Ora di Lavoro"
        )
        fig4.update_layout(xaxis_title="Macchina", yaxis_title="Consumo (kWh) per Ora")
        st.plotly_chart(fig4, use_container_width=True)
    else:
        st.info("Nessun dato disponibile con i filtri selezionati.")

# --- Tabella Riepilogativa ---
st.subheader("📋 Riepilogo per Macchina")

if not df_filtrato.empty:
    riepilogo = df_filtrato.groupby('macchina').agg({
        'consumo_kwh': 'sum',
        'costo_macchina': 'sum',
        'pezzi_prodotti': 'sum',
        'ore_produzione': 'sum'
    }).round(2)

    riepilogo['consumo_per_pezzo'] = (riepilogo['consumo_kwh'] / riepilogo['pezzi_prodotti']).round(4)
    riepilogo['costo_per_pezzo'] = (riepilogo['costo_macchina'] / riepilogo['pezzi_prodotti']).round(4)

    st.dataframe(riepilogo, use_container_width=True)
else:
    st.info("Nessun dato disponibile con i filtri selezionati.")

# --- Sezione "Informazioni" ---
st.sidebar.markdown("---")
st.sidebar.info(
    """
    **Note sull'Applicazione:**
    - I dati sono tratti dal foglio 'Consolidato' del file Excel.
    - I valori mancanti o non numerici (es. "- 0 €") sono stati convertiti in `NaN`.
    - Puoi filtrare per macchina, anno e mese in modo indipendente.
    - I grafici sono interattivi: puoi fare zoom, pan e hover sui dati.
    """
)