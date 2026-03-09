import streamlit as st
import pandas as pd
import gdown
import os
from dotenv import load_dotenv

# =========================
# CONFIG
# =========================

load_dotenv()

FILE_ID = os.getenv("DRIVE_FILE_ID")
FILE_PATH = "dados.xlsx"

# =========================
# LOGIN
# =========================

USUARIOS = {
    "admin": "123",
    "usuario1": "123",
    "usuario2": "123",
    "usuario3": "123",
    "usuario4": "123"
}

def login():
    st.title("Sistema de Busca de Cadastros")

    user = st.text_input("Usuário")
    senha = st.text_input("Senha", type="password")

    if st.button("Entrar"):
        if user in USUARIOS and senha == USUARIOS[user]:
            st.session_state["logado"] = True
        else:
            st.error("Usuário ou senha inválidos")

# =========================
# DOWNLOAD DO DRIVE
# =========================

def baixar_planilha():
    if not os.path.exists(FILE_PATH):
        url = f"https://drive.google.com/uc?id={FILE_ID}"
        gdown.download(url, FILE_PATH, quiet=False)

# =========================
# CARREGAR ABA
# =========================

@st.cache_data
def carregar_aba(nome_aba):

    df = pd.read_excel(
        FILE_PATH,
        sheet_name=nome_aba,
        header=2
    )

    df.columns = df.columns.str.strip()

    return df

# =========================
# INTERFACE
# =========================

def app():

    st.title("🔎 Busca de Cadastro de Eletricistas")

    abas = [
        "PARÁ",
        "MARANHÃO",
        "PIAUÍ",
        "ALAGOAS",
        "AMAPÁ",
        "RIO GRANDE DO SUL",
        "GOIÁS",
        "TODOS"
    ]

    aba = st.selectbox("Selecione a Região", abas)

    if aba != "TODOS":
        df = carregar_aba(aba)

    else:
        planilha = pd.ExcelFile(FILE_PATH)
        todos = []

        for a in planilha.sheet_names:
            d = pd.read_excel(FILE_PATH, sheet_name=a, header=2)
            todos.append(d)

        df = pd.concat(todos)

    col1, col2 = st.columns(2)

    with col1:
        busca_nota = st.text_input("Número da Nota")

    with col2:
        busca_nome = st.text_input("Nome do Eletricista")

    busca_data = st.date_input("Data da Nota")

    st.divider()

    st.subheader("Busca em Massa")

    lista_notas = st.text_area(
        "Cole vários números de nota (1 por linha)"
    )

    resultado = df.copy()

    # =========================
    # FILTROS
    # =========================

    if busca_nota:
        resultado = resultado[
            resultado["NOTA"].astype(str).str.contains(busca_nota)
        ]

    if busca_nome:
        resultado = resultado[
            resultado["ELETRICISTA"].str.contains(busca_nome, case=False)
        ]

    if busca_data:
        resultado = resultado[
            pd.to_datetime(resultado["DATA"]).dt.date == busca_data
        ]

    if lista_notas:

        notas = lista_notas.splitlines()

        resultado = df[
            df["NOTA"].astype(str).isin(notas)
        ]

    st.divider()

    st.write("Resultados encontrados:", len(resultado))

    st.dataframe(resultado, use_container_width=True)

    # DOWNLOAD

    if not resultado.empty:

        st.download_button(
            "Baixar resultado em Excel",
            resultado.to_excel(index=False),
            "resultado.xlsx"
        )

# =========================
# MAIN
# =========================

if "logado" not in st.session_state:
    st.session_state["logado"] = False

if not st.session_state["logado"]:
    login()
else:
    baixar_planilha()
    app()
