import re
from io import BytesIO
from datetime import date

import pandas as pd
import requests
import streamlit as st
import streamlit.components.v1 as components

# ======================================================
# CONFIG
# ======================================================
st.set_page_config(
    page_title="Buscar Cadastro",
    layout="wide"
)

# ======================================================
# CSS
# ======================================================
st.markdown("""
<style>
.stApp { background: #6fa6d6; }
.block-container{ padding-top: 0.6rem; max-width: 1500px; }

/* LOGIN MODE */
.stApp.login-mode .block-container{
  max-width: 980px !important;
  padding-top: 20px !important;
}
.stApp.login-mode header,
.stApp.login-mode footer,
.stApp.login-mode [data-testid="stSidebar"],
.stApp.login-mode [data-testid="stToolbar"]{
  display:none !important;
}

.login-shell{
  display:flex;
  justify-content:center;
  align-items:flex-start;
  padding: 12px 0 0 0;
}
.login-card{
  width: min(980px, 94vw);
  border-radius: 26px;
  overflow: hidden;
  box-shadow: 0 16px 28px rgba(0,0,0,0.28);
  border: 2px solid rgba(10,40,70,0.20);
}
.login-head{
  padding: 22px 26px 18px 26px;
  background: #2f6f8c;
}
.login-title{
  display:flex; align-items:center; gap:12px;
  font-weight: 950;
  color:#ffffff;
  font-size: 34px;
  line-height: 1;
}
.login-sub{
  margin-top: 8px;
  color: rgba(255,255,255,0.92);
  font-weight: 800;
  font-size: 16px;
}
.login-band{
  height: 46px;
  background: rgba(255,255,255,0.35);
}
.login-body{
  padding: 26px 22px 10px 22px;
  background: transparent;
}
.stApp.login-mode div[data-testid="stTextInput"]{
  margin-bottom: 10px;
}
.stApp.login-mode div[data-testid="stTextInput"] input{
  height: 44px !important;
  border-radius: 10px !important;
  background: #15181c !important;
  border: 1px solid rgba(255,255,255,0.18) !important;
  color: #ffffff !important;
  font-weight: 900 !important;
  font-size: 16px !important;
}
.stApp.login-mode div[data-testid="stTextInput"] input::placeholder{
  color: rgba(255,255,255,0.45) !important;
  font-weight: 800 !important;
}
.stApp.login-mode div[data-testid="stTextInput"] svg{
  color: rgba(255,255,255,0.75) !important;
}
.stApp.login-mode div.stButton > button{
  border-radius: 10px !important;
  font-weight: 900 !important;
  border: 2px solid rgba(10,40,70,0.22) !important;
  background: rgba(255,255,255,0.55) !important;
  color:#0b2b45 !important;
  padding: .30rem .90rem !important;
}
.stApp.login-mode div.stButton > button:hover{
  background: rgba(255,255,255,0.75) !important;
  border-color: rgba(10,40,70,0.35) !important;
}
.login-foot{
  text-align:center;
  padding: 10px 0 16px 0;
  font-weight: 900;
  color: rgba(11,43,69,0.92);
  font-size: 12px;
}
.login-foot .ok{
  display:inline-flex; align-items:center; gap:8px;
}
.login-foot .dot{
  width: 12px; height: 12px;
  border-radius: 3px;
  background: #2e7d32;
  box-shadow: 0 0 0 2px rgba(255,255,255,0.35) inset;
}

/* DASHBOARD */
.topbar{
  background: rgba(255,255,255,0.35);
  border: 2px solid rgba(10,40,70,0.22);
  border-radius: 18px;
  padding: 10px 14px;
  display:flex;
  justify-content:space-between;
  align-items:center;
  margin-bottom: 10px;
}
.brand{
  display:flex; align-items:center; gap:12px;
}
.brand-badge{
  width:46px; height:46px; border-radius: 14px;
  background: rgba(255,255,255,0.55);
  border: 2px solid rgba(10,40,70,0.22);
  display:flex; align-items:center; justify-content:center;
  font-weight: 950; color:#0b2b45;
}
.brand-text .t1{ font-weight:950; color:#0b2b45; line-height:1.1; }
.brand-text .t2{ font-weight:800; color:#0b2b45; opacity:.85; font-size:12px; }

.right-note{
  text-align:right; font-weight:950; color:#0b2b45;
}
.right-note small{ font-weight:800; opacity:.9; font-size:12px; }

.card{
  background: #b9d3ee;
  border: 2px solid rgba(10,40,70,0.30);
  border-radius: 18px;
  padding: 14px 16px;
  box-shadow: 0 10px 18px rgba(0,0,0,0.18);
  margin-bottom: 14px;
}
.card-title{
  font-weight: 950;
  color:#0b2b45;
  font-size: 13px;
  text-transform: uppercase;
  margin-bottom: 10px;
  letter-spacing: .3px;
}
.info-chip{
  display:inline-block;
  background: rgba(255,255,255,0.55);
  border: 1px solid rgba(10,40,70,0.20);
  color:#0b2b45;
  border-radius: 999px;
  padding: 6px 12px;
  margin-right: 8px;
  margin-bottom: 8px;
  font-weight: 800;
  font-size: 12px;
}
div.stButton > button{
  border-radius: 10px;
  font-weight: 900;
  border: 2px solid rgba(10,40,70,0.22);
  background: rgba(255,255,255,0.45);
  color:#0b2b45;
  padding: .25rem .6rem;
}
div.stButton > button:hover{
  background: rgba(255,255,255,0.65);
  border-color: rgba(10,40,70,0.35);
}
div[data-baseweb="segmented-control"]{
  background: rgba(255,255,255,0.35);
  border: 2px solid rgba(10,40,70,0.22);
  border-radius: 14px;
  padding: 6px;
}
div[data-baseweb="segmented-control"] span{
  font-weight: 900 !important;
  color: #0b2b45 !important;
}
div[data-baseweb="segmented-control"] div[aria-checked="true"]{
  background: #0b2b45 !important;
  border-radius: 10px !important;
}
div[data-baseweb="segmented-control"] div[aria-checked="true"] span{
  color: #ffffff !important;
}
</style>
""", unsafe_allow_html=True)

# ======================================================
# LOGIN HELPERS
# ======================================================
def _set_login_mode(on: bool):
    if on:
        components.html("""
        <script>
          const app = window.parent.document.querySelector('.stApp');
          if (app) app.classList.add('login-mode');
        </script>
        """, height=0)
    else:
        components.html("""
        <script>
          const app = window.parent.document.querySelector('.stApp');
          if (app) app.classList.remove('login-mode');
        </script>
        """, height=0)

def carregar_usuarios():
    """
    Espera em st.secrets:
    [auth]
    usuarios = [
      {usuario="admin", senha="123", perfil="admin"},
      {usuario="usuario1", senha="123", perfil="consulta"},
      ...
    ]
    """
    auth = st.secrets.get("auth", {})
    usuarios = auth.get("usuarios", [])

    mapa = {}
    for item in usuarios:
        user = str(item.get("usuario", "")).strip()
        senha = str(item.get("senha", "")).strip()
        perfil = str(item.get("perfil", "consulta")).strip()
        if user:
            mapa[user] = {"senha": senha, "perfil": perfil}
    return mapa

def autenticar(usuario: str, senha: str):
    usuarios = carregar_usuarios()
    if usuario in usuarios and senha == usuarios[usuario]["senha"]:
        return True, usuarios[usuario]["perfil"]
    return False, None

def logout():
    st.session_state["logado"] = False
    st.session_state["usuario_logado"] = None
    st.session_state["perfil_logado"] = None
    st.rerun()

def tela_login():
    _set_login_mode(True)

    st.markdown("""
    <div class="login-shell">
      <div class="login-card">
        <div class="login-head">
          <div class="login-title">🔐&nbsp;Acesso Restrito</div>
          <div class="login-sub">Buscar Cadastro • Notas / Eletricista</div>
        </div>
        <div class="login-band"></div>
        <div class="login-body">
    """, unsafe_allow_html=True)

    usuario = st.text_input(
        label="Usuário",
        placeholder="Digite seu usuário",
        key="login_user",
        label_visibility="collapsed",
    )
    senha = st.text_input(
        label="Senha",
        placeholder="Digite sua senha",
        type="password",
        key="login_pass",
        label_visibility="collapsed",
    )

    c1, c2, c3 = st.columns([1, 1, 6])
    with c1:
        entrar = st.button("Entrar", key="btn_entrar")
    with c2:
        limpar = st.button("Limpar", key="btn_limpar")

    if limpar:
        st.session_state["login_user"] = ""
        st.session_state["login_pass"] = ""
        st.rerun()

    if entrar:
        ok, perfil = autenticar(usuario, senha)
        if ok:
            st.session_state["logado"] = True
            st.session_state["usuario_logado"] = usuario
            st.session_state["perfil_logado"] = perfil
            st.rerun()
        else:
            st.error("Usuário ou senha inválidos.")

    st.markdown("""
        </div>
        <div class="login-foot">
          <span class="ok"><span class="dot"></span>Segurança via st.secrets</span>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

# ======================================================
# INICIALIZAÇÃO LOGIN
# ======================================================
if "logado" not in st.session_state:
    st.session_state["logado"] = False
if "usuario_logado" not in st.session_state:
    st.session_state["usuario_logado"] = None
if "perfil_logado" not in st.session_state:
    st.session_state["perfil_logado"] = None

if not st.session_state["logado"]:
    tela_login()
    st.stop()
else:
    _set_login_mode(False)

# ======================================================
# TOPO
# ======================================================
st.markdown(f"""
<div class="topbar">
  <div class="brand">
    <div class="brand-badge">BC</div>
    <div class="brand-text">
      <div class="t1">BUSCAR CADASTRO</div>
      <div class="t2">Consulta por nota, data e eletricista</div>
    </div>
  </div>
  <div class="right-note">
    Usuário: {st.session_state.get("usuario_logado", "-")}<br>
    <small>Perfil: {st.session_state.get("perfil_logado", "-")}</small>
  </div>
</div>
""", unsafe_allow_html=True)

c_top1, c_top2 = st.columns([1, 7])
with c_top1:
    if st.button("🚪 Sair"):
        logout()
with c_top2:
    st.caption("Sistema de consulta em XLSX/CSV hospedado no Google Drive.")

# ======================================================
# HELPERS
# ======================================================
def normalizar_colunas(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = (
        df.columns.astype(str)
        .str.upper()
        .str.strip()
        .str.replace("\n", " ", regex=False)
    )
    return df

def achar_coluna(df: pd.DataFrame, palavras):
    for col in df.columns:
        col_up = str(col).upper().strip()
        for p in palavras:
            if str(p).upper().strip() in col_up:
                return col
    return None

def extrair_drive_id(url: str):
    if not url:
        return None
    m = re.search(r"[?&]id=([a-zA-Z0-9-_]+)", url)
    if m:
        return m.group(1)
    m = re.search(r"/file/d/([a-zA-Z0-9-_]+)", url)
    if m:
        return m.group(1)
    return None

def drive_direct_download(url: str) -> str:
    did = extrair_drive_id(url)
    if did:
        return f"https://drive.google.com/uc?id={did}&export=download"
    return url

def bytes_is_html(raw: bytes) -> bool:
    head = raw[:1000].lstrip().lower()
    return head.startswith(b"<!doctype html") or b"<html" in head

def bytes_is_xlsx(raw: bytes) -> bool:
    return raw[:2] == b"PK"

@st.cache_data(ttl=600, show_spinner="🔄 Carregando base...")
def carregar_base(url_original: str, sheet_name=0) -> pd.DataFrame:
    url = drive_direct_download(url_original)
    r = requests.get(url, timeout=60)
    r.raise_for_status()
    raw = r.content

    if bytes_is_html(raw):
        raise RuntimeError(
            "O link retornou HTML em vez do arquivo. "
            "No Google Drive, deixe como 'Qualquer pessoa com o link' e Visualizador."
        )

    if bytes_is_xlsx(raw):
        df = pd.read_excel(BytesIO(raw), sheet_name=sheet_name, engine="openpyxl")
        return normalizar_colunas(df)

    for enc in ["utf-8-sig", "utf-8", "cp1252", "latin1"]:
        try:
            df = pd.read_csv(BytesIO(raw), sep=None, engine="python", encoding=enc)
            return normalizar_colunas(df)
        except Exception:
            pass

    raise RuntimeError("Não foi possível ler o arquivo como XLSX ou CSV.")

def tentar_converter_data(serie: pd.Series) -> pd.Series:
    return pd.to_datetime(serie, errors="coerce", dayfirst=True)

def limpar_texto(s: pd.Series) -> pd.Series:
    return s.fillna("").astype(str).str.upper().str.strip()

def somente_digitos(s: pd.Series) -> pd.Series:
    return s.fillna("").astype(str).str.replace(r"\D", "", regex=True)

def download_excel(df: pd.DataFrame, nome_arquivo: str):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="RESULTADO")
    output.seek(0)
    st.download_button(
        label="⬇️ Baixar resultado em Excel",
        data=output,
        file_name=nome_arquivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

def exibir_resumo(df_base: pd.DataFrame, col_nota, col_data, col_eletricista):
    st.markdown('<div class="card"><div class="card-title">Resumo da Base</div>', unsafe_allow_html=True)

    total = len(df_base)
    datas_validas = 0
    if col_data:
        datas_validas = df_base[col_data].notna().sum()

    st.markdown(
        f"""
        <span class="info-chip">Linhas: {total:,}</span>
        <span class="info-chip">Coluna Nota: {col_nota or "não encontrada"}</span>
        <span class="info-chip">Coluna Data: {col_data or "não encontrada"}</span>
        <span class="info-chip">Coluna Eletricista: {col_eletricista or "não encontrada"}</span>
        <span class="info-chip">Datas válidas: {datas_validas:,}</span>
        """.replace(",", "."),
        unsafe_allow_html=True
    )
    st.markdown("</div>", unsafe_allow_html=True)

# ======================================================
# CONFIG DA BASE
# ======================================================
st.markdown('<div class="card"><div class="card-title">Configuração da Base</div>', unsafe_allow_html=True)

url_padrao = st.secrets.get("drive", {}).get("url_base", "")
sheet_padrao = st.secrets.get("drive", {}).get("sheet_name", 0)

c_cfg1, c_cfg2, c_cfg3 = st.columns([4, 1, 1])

with c_cfg1:
    url_base = st.text_input("Link do arquivo no Google Drive", value=url_padrao)

with c_cfg2:
    sheet_name = st.text_input("Aba", value=str(sheet_padrao))

with c_cfg3:
    st.write("")
    st.write("")
    if st.button("🔄 Atualizar base"):
        st.cache_data.clear()
        st.rerun()

st.markdown("</div>", unsafe_allow_html=True)

# ======================================================
# CARREGAR BASE
# ======================================================
try:
    sheet_name_final = 0 if str(sheet_name).strip() == "" else sheet_name
    df = carregar_base(url_base, sheet_name=sheet_name_final)
except Exception as e:
    st.error(f"Erro ao carregar a base: {e}")
    st.stop()

# ======================================================
# IDENTIFICAÇÃO DE COLUNAS
# ======================================================
COL_NOTA = achar_coluna(df, [
    "NUMERO DA NOTA", "NÚMERO DA NOTA", "NOTA", "NUM NOTA", "Nº NOTA", "NF", "NUMERO NOTA"
])

COL_DATA = achar_coluna(df, [
    "DATA", "DATA DA NOTA", "DT NOTA", "EMISSAO", "EMISSÃO"
])

COL_ELETRICISTA = achar_coluna(df, [
    "ELETRICISTA", "NOME DO ELETRICISTA", "NOME ELETRICISTA", "RESPONSAVEL", "RESPONSÁVEL", "NOME"
])

if COL_DATA:
    df[COL_DATA] = tentar_converter_data(df[COL_DATA])

if COL_NOTA:
    df["_NOTA_TXT_"] = df[COL_NOTA].fillna("").astype(str).str.strip()
    df["_NOTA_NUM_"] = somente_digitos(df[COL_NOTA])
else:
    df["_NOTA_TXT_"] = ""
    df["_NOTA_NUM_"] = ""

if COL_ELETRICISTA:
    df["_ELETRICISTA_"] = limpar_texto(df[COL_ELETRICISTA])
else:
    df["_ELETRICISTA_"] = ""

exibir_resumo(df, COL_NOTA, COL_DATA, COL_ELETRICISTA)

# ======================================================
# BUSCAS
# ======================================================
modo = st.segmented_control(
    "Modo de consulta",
    options=["Consulta individual", "Consulta em massa"],
    default="Consulta individual",
    key="modo_consulta"
)

# ======================================================
# CONSULTA INDIVIDUAL
# ======================================================
if modo == "Consulta individual":
    st.markdown('<div class="card"><div class="card-title">Consulta Individual</div>', unsafe_allow_html=True)

    tipo_busca = st.segmented_control(
        "Buscar por",
        options=["Número da nota", "Data", "Eletricista"],
        default="Número da nota",
        key="tipo_busca_individual"
    )

    resultado = pd.DataFrame()

    if tipo_busca == "Número da nota":
        valor_nota = st.text_input("Digite o número da nota")
        busca_exata = st.checkbox("Busca exata", value=False)

        if st.button("🔎 Pesquisar nota"):
            if not COL_NOTA:
                st.warning("A coluna de nota não foi encontrada.")
            else:
                termo = re.sub(r"\D", "", str(valor_nota))
                if not termo:
                    st.warning("Digite um número de nota válido.")
                else:
                    if busca_exata:
                        resultado = df[df["_NOTA_NUM_"] == termo].copy()
                    else:
                        resultado = df[df["_NOTA_NUM_"].str.contains(termo, na=False)].copy()

    elif tipo_busca == "Data":
        if not COL_DATA:
            st.warning("A coluna de data não foi encontrada.")
        else:
            data_busca = st.date_input("Selecione a data", value=date.today())
            if st.button("🔎 Pesquisar data"):
                data_busca_ts = pd.to_datetime(data_busca)
                resultado = df[df[COL_DATA].dt.date == data_busca_ts.date()].copy()

    elif tipo_busca == "Eletricista":
        nome_eletricista = st.text_input("Digite o nome do eletricista")
        busca_exata_nome = st.checkbox("Busca exata do nome", value=False)

        if st.button("🔎 Pesquisar eletricista"):
            if not COL_ELETRICISTA:
                st.warning("A coluna de eletricista não foi encontrada.")
            else:
                termo = str(nome_eletricista).strip().upper()
                if not termo:
                    st.warning("Digite um nome para pesquisar.")
                else:
                    if busca_exata_nome:
                        resultado = df[df["_ELETRICISTA_"] == termo].copy()
                    else:
                        resultado = df[df["_ELETRICISTA_"].str.contains(termo, na=False)].copy()

    if not resultado.empty:
        st.success(f"Encontrados {len(resultado):,} registros.".replace(",", "."))
        st.dataframe(resultado, use_container_width=True, hide_index=True)
        download_excel(resultado, "resultado_consulta_individual.xlsx")
    elif "resultado" in locals():
        if isinstance(resultado, pd.DataFrame) and resultado.empty:
            st.info("Nenhum registro encontrado.")

    st.markdown("</div>", unsafe_allow_html=True)

# ======================================================
# CONSULTA EM MASSA
# ======================================================
else:
    st.markdown('<div class="card"><div class="card-title">Consulta em Massa</div>', unsafe_allow_html=True)

    tipo_massa = st.segmented_control(
        "Consulta em massa por",
        options=["Número da nota", "Eletricista"],
        default="Número da nota",
        key="tipo_busca_massa"
    )

    if tipo_massa == "Número da nota":
        texto_notas = st.text_area(
            "Cole os números das notas (um por linha, vírgula ou espaço)",
            height=180,
            placeholder="Exemplo:\n12345\n12346\n12347"
        )

        busca_exata_massa = st.checkbox("Busca exata das notas", value=True)

        if st.button("🔎 Pesquisar em massa"):
            if not COL_NOTA:
                st.warning("A coluna de nota não foi encontrada.")
            else:
                itens = re.split(r"[\n,;\t ]+", str(texto_notas).strip())
                itens = [re.sub(r"\D", "", x) for x in itens if str(x).strip()]
                itens = [x for x in itens if x]

                if not itens:
                    st.warning("Informe pelo menos uma nota.")
                else:
                    itens_unicos = list(dict.fromkeys(itens))

                    if busca_exata_massa:
                        resultado = df[df["_NOTA_NUM_"].isin(itens_unicos)].copy()
                    else:
                        padrao = "|".join(re.escape(x) for x in itens_unicos)
                        resultado = df[df["_NOTA_NUM_"].str.contains(padrao, na=False, regex=True)].copy()

                    encontrados = set(resultado["_NOTA_NUM_"].dropna().astype(str).tolist())
                    nao_encontrados = [x for x in itens_unicos if x not in encontrados]

                    c_res1, c_res2 = st.columns([3, 2])

                    with c_res1:
                        st.success(
                            f"Notas consultadas: {len(itens_unicos):,} | "
                            f"Encontradas: {len(encontrados):,} | "
                            f"Não encontradas: {len(nao_encontrados):,}".replace(",", ".")
                        )
                        st.dataframe(resultado, use_container_width=True, hide_index=True)

                    with c_res2:
                        st.markdown("**Não encontradas**")
                        if nao_encontrados:
                            st.dataframe(pd.DataFrame({"NOTA_NAO_ENCONTRADA": nao_encontrados}), use_container_width=True, hide_index=True)
                        else:
                            st.info("Todas foram encontradas.")

                    if not resultado.empty:
                        download_excel(resultado, "resultado_consulta_massa_notas.xlsx")

    else:
        texto_nomes = st.text_area(
            "Cole os nomes dos eletricistas (um por linha)",
            height=180,
            placeholder="Exemplo:\nJOAO\nMARCOS\nCARLOS"
        )

        busca_exata_nome_massa = st.checkbox("Busca exata dos nomes", value=False)

        if st.button("🔎 Pesquisar nomes em massa"):
            if not COL_ELETRICISTA:
                st.warning("A coluna de eletricista não foi encontrada.")
            else:
                itens = re.split(r"[\n]+", str(texto_nomes).strip())
                itens = [str(x).strip().upper() for x in itens if str(x).strip()]
                itens_unicos = list(dict.fromkeys(itens))

                if not itens_unicos:
                    st.warning("Informe pelo menos um nome.")
                else:
                    if busca_exata_nome_massa:
                        resultado = df[df["_ELETRICISTA_"].isin(itens_unicos)].copy()
                    else:
                        padrao = "|".join(re.escape(x) for x in itens_unicos)
                        resultado = df[df["_ELETRICISTA_"].str.contains(padrao, na=False, regex=True)].copy()

                    encontrados = set(resultado["_ELETRICISTA_"].dropna().astype(str).tolist())
                    nao_encontrados = [x for x in itens_unicos if x not in encontrados]

                    c_res1, c_res2 = st.columns([3, 2])

                    with c_res1:
                        st.success(
                            f"Nomes consultados: {len(itens_unicos):,} | "
                            f"Encontrados: {len(encontrados):,} | "
                            f"Não encontrados: {len(nao_encontrados):,}".replace(",", ".")
                        )
                        st.dataframe(resultado, use_container_width=True, hide_index=True)

                    with c_res2:
                        st.markdown("**Não encontrados**")
                        if nao_encontrados:
                            st.dataframe(pd.DataFrame({"ELETRICISTA_NAO_ENCONTRADO": nao_encontrados}), use_container_width=True, hide_index=True)
                        else:
                            st.info("Todos foram encontrados.")

                    if not resultado.empty:
                        download_excel(resultado, "resultado_consulta_massa_eletricistas.xlsx")

    st.markdown("</div>", unsafe_allow_html=True)

# ======================================================
# VISUALIZAÇÃO DA BASE
# ======================================================
with st.expander("📄 Visualizar base completa"):
    st.dataframe(df, use_container_width=True, hide_index=True)
