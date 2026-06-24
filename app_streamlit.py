import streamlit as st
import pandas as pd
import io
import datetime

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Processador de Obras",
    page_icon="🏗️",
    layout="centered"
)

# --- FUNÇÕES DE PROCESSAMENTO ---
def extract_empreendimento(val):
    if pd.isna(val): 
        return None
    txt = str(val).strip()
    if txt.startswith("Listagem de Execução Orçamentária"):
        partes = [p.strip() for p in txt.split(" - ")]
        if len(partes) >= 3:
            return partes[1]
    return None

def extract_hierarquia(val):
    if pd.isna(val): 
        return None
    txt = str(val).strip()
    if "/" in txt:
        primeiro = txt[0] if len(txt) > 0 else ""
        if not primeiro.isdigit():
            return txt
    return None

def is_date(val):
    if pd.isna(val):
        return False
    if isinstance(val, (datetime.datetime, datetime.date)):
        return True
    try:
        val_str = str(val).strip().lower()
        if val_str == "data" or val_str == "":
            return False
        pd.to_datetime(val_str)
        return True
    except:
        return False

def processar_dataframe(df_raw):
    if len(df_raw.columns) < 4:
        raise ValueError("O arquivo não tem pelo menos 4 colunas. Verifique o formato do relatório.")
        
    df = df_raw.iloc[:, :4].copy()
    df.columns = ["Column1", "Column2", "Column3", "Column4"]

    df["Empreendimento"] = df["Column1"].apply(extract_empreendimento)
    df["Empreendimento"] = df["Empreendimento"].ffill()

    df["Hierarquia"] = df["Column1"].apply(extract_hierarquia)
    df["Hierarquia"] = df["Hierarquia"].ffill()

    df_valid = df.dropna(subset=["Column1", "Column2", "Column3", "Column4"]).copy()
    mask_is_date = df_valid["Column1"].apply(is_date)
    df_valid = df_valid[mask_is_date].copy()

    df_valid.rename(columns={
        "Column1": "Data",
        "Column2": "Origem",
        "Column3": "Descrição",
        "Column4": "Valor"
    }, inplace=True)

    df_valid["Data"] = pd.to_datetime(df_valid["Data"], errors='coerce').dt.date
    df_valid["Origem"] = df_valid["Origem"].astype(str).str.strip()
    df_valid["Descrição"] = df_valid["Descrição"].astype(str).str.strip()
    
    if df_valid["Valor"].dtype == 'O': 
        df_valid["Valor"] = df_valid["Valor"].astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    df_valid["Valor"] = pd.to_numeric(df_valid["Valor"], errors='coerce')
    
    df_valid["Hierarquia"] = df_valid["Hierarquia"].astype(str)
    df_valid["Empreendimento"] = df_valid["Empreendimento"].astype(str)

    split_h = df_valid["Hierarquia"].str.split("/", expand=True)
    for col_idx in range(3):
        if col_idx not in split_h.columns:
            split_h[col_idx] = None
            
    df_valid["Etapa de Obra"] = split_h[0].astype(str).str.strip().replace("None", None)
    df_valid["Categoria"] = split_h[1].astype(str).str.strip().replace("None", None)
    df_valid["Insumo"] = split_h[2].astype(str).str.strip().replace("None", None)

    meses_pt = {
        1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
        5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
        9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
    }
    
    data_dt = pd.to_datetime(df_valid["Data"])
    df_valid["Mês"] = data_dt.dt.month.map(meses_pt)
    df_valid["Ano"] = data_dt.dt.year

    colunas_finais = [
        "Empreendimento",
        "Etapa de Obra",
        "Categoria",
        "Insumo",
        "Data",
        "Mês",
        "Ano",
        "Origem",
        "Descrição",
        "Valor"
    ]
    
    df_final = df_valid[[c for c in colunas_finais if c in df_valid.columns]]
    return df_final

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Lançamentos Tratados')
    processed_data = output.getvalue()
    return processed_data

# --- UI DA APLICAÇÃO ---
st.title("🏗️ Processador de Obras")
st.markdown("Faça o upload do seu relatório de execução orçamentária desestruturado para tratá-lo automaticamente.")

uploaded_file = st.file_uploader("Arraste e solte o arquivo Excel", type=["xlsx", "xls"])

if uploaded_file is not None:
    with st.spinner("Lendo e processando o arquivo..."):
        try:
            # Lendo o arquivo submetido
            df_raw = pd.read_excel(uploaded_file, header=None)
            
            # Processamento
            df_tratado = processar_dataframe(df_raw)
            
            # Sucesso
            st.success("✅ Processamento concluído com sucesso!")
            
            # Mostrar métricas
            col1, col2 = st.columns(2)
            col1.metric("Lançamentos Tratados", len(df_tratado))
            col2.metric("Empreendimentos Únicos", df_tratado["Empreendimento"].nunique())
            
            # Preview dos dados
            with st.expander("Pré-visualização dos dados tratados"):
                st.dataframe(df_tratado.head(15))
            
            # Botão de download
            excel_data = to_excel(df_tratado)
            st.download_button(
                label="📥 Baixar Planilha Tratada (.xlsx)",
                data=excel_data,
                file_name="relatorio_obras_tratado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )

        except Exception as e:
            st.error(f"❌ Ocorreu um erro ao processar o arquivo. Detalhes técnicos: {str(e)}")
