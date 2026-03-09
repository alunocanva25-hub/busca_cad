import os
import io
import re
import unicodedata
from pathlib import Path

import pandas as pd
import streamlit as st

# =========================================================
# CONFIG INICIAL
# =========================================================

st.set_page_config(
    page_title="Buscar Cadastro",
    page_icon="🔎",
    layout="wide"
)

try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass

try:
    import gdown
except Exception:
    gdown = None


def get_secret_or_env(name, default=""):
    try:
        if name in st.secrets:
            return st.secrets[name]
    except Exception:
        pass
    return os.getenv(name, default)


DEFAULT_XLSX_PATH = get_secret_or_env("XLSX_PATH", "dados.xlsx")
DEFAULT_DRIVE_FILE_ID = get_secret_or_env("DRIVE_FILE_ID", "")

USUARIOS = {
    "admin": "123",
    "usuario1": "123",
    "usuario2": "123",
    "usuario3": "123",
    "usuario4": "123",
}

# =========================================================
# CSS
# =========================================================

st.markdown("""
<style>
.block-container {
    padding-top: 1.2rem;
    padding-bottom: 1rem;
}

.login-wrap {
    max-width: 460px;
    margin: 5vh auto 0 auto;
}

.login-box {
    background: linear-gradient(180deg, #0f172a 0%, #111827 100%);
    border: 1px solid rgba(255,255,255,0.08);
    border-radius: 18px;
    padding: 28px;
    box-shadow: 0 10px 30px rgba(0,0,0,0.25);
}

.login-title {
    font-size: 28px;
    font-weight: 700;
    margin-bottom: 6px;
    text-align: center;
}

.login-subtitle {
    color: #cbd5e1;
    margin-bottom: 20px;
    text-align: center;
}

.small-note {
    color: #94a3b8;
    font-size: 0.9rem;
    text-align: center;
    margin-top: 12px;
}

div[data-testid="stTextInput"] input {
    border-radius: 10px;
}

div[data-testid="stButton"] button {
    border-radius: 10px;
    font-weight: 600;
}

.result-box {
    padding: 12px;
    border-radius: 12px;
    background: rgba(59,130,246,0.08);
    border: 1px solid rgba(59,130,246,0.15);
}
</style>
""", unsafe_allow_html=True)

# =========================================================
# UTILITÁRIOS
# =========================================================

def normalizar_texto(valor):
    if pd.isna(valor):
        return ""
    valor = str(valor).strip()
    valor = unicodedata.normalize("NFKD", valor)
    valor = "".join(c for c in valor if not unicodedata.combining(c))
    valor = re.sub(r"\s+", " ", valor)
    return valor.upper().strip()

def slug_coluna(valor):
    txt = normalizar_texto(valor)
    txt = txt.replace("Ç", "C")
    txt = re.sub(r"[^A-Z0-9]+", "_", txt)
    txt = re.sub(r"_+", "_", txt).strip("_")
    return txt

def safe_str_contains(series, termo):
    termo_norm = normalizar_texto(termo)
    serie_norm = series.fillna("").astype(str).map(normalizar_texto)
    return serie_norm.str.contains(re.escape(termo_norm), na=False)

def detectar_cabecalho(path_xlsx, sheet_name, max_linhas=12):
    bruto = pd.read_excel(path_xlsx, sheet_name=sheet_name, header=None, nrows=max_linhas)
    melhor_idx = 2
    melhor_score = -1

    chaves_nota = ["NOTA", "NF", "NUMERO_NOTA", "N_NOTA"]
    chaves_data = ["DATA", "DT", "DATA_NOTA"]
    chaves_nome = ["ELETRICISTA", "NOME", "NOME_ELETRICISTA", "PRESTADOR", "COLABORADOR"]

    for idx in range(len(bruto)):
        linha = [slug_coluna(v) for v in bruto.iloc[idx].tolist()]
        score = 0

        for c in linha:
            if any(k in c for k in chaves_nota):
                score += 3
            if any(k in c for k in chaves_data):
                score += 3
            if any(k in c for k in chaves_nome):
                score += 4

        if score > melhor_score:
            melhor_score = score
            melhor_idx = idx

    return melhor_idx

def localizar_coluna(df, candidatos):
    mapa = {slug_coluna(c): c for c in df.columns}

    for cand in candidatos:
        cand_slug = slug_coluna(cand)
        if cand_slug in mapa:
            return mapa[cand_slug]

    for slug, original in mapa.items():
        for cand in candidatos:
            cand_slug = slug_coluna(cand)
            if cand_slug in slug:
                return original

    return None

def limpar_dataframe(df):
    df = df.dropna(axis=1, how="all").copy()

    novas_cols = []
    for i, col in enumerate(df.columns):
        col_str = str(col).strip()
        if not col_str or col_str.upper().startswith("UNNAMED"):
            col_str = f"COLUNA_{i+1}"
        novas_cols.append(col_str)
    df.columns = novas_cols

    df = df.dropna(axis=0, how="all").copy()
    return df

def preparar_dataframe(path_xlsx, sheet_name):
    header_idx = detectar_cabecalho(path_xlsx, sheet_name)
    df = pd.read_excel(path_xlsx, sheet_name=sheet_name, header=header_idx)
    df = limpar_dataframe(df)

    col_nota = localizar_coluna(df, [
        "NOTA", "NF", "NUMERO DA NOTA", "NUMERO NOTA", "N DA NOTA", "N_NOTA"
    ])
    col_data = localizar_coluna(df, [
        "DATA", "DATA DA NOTA", "DT", "EMISSAO", "DATA EMISSAO"
    ])
    col_nome = localizar_coluna(df, [
        "ELETRICISTA", "NOME", "NOME DO ELETRICISTA", "PRESTADOR", "COLABORADOR", "NOME COMPLETO"
    ])

    if len(df) > 0:
        primeira = [normalizar_texto(x) for x in df.iloc[0].tolist()]
        cab = [normalizar_texto(x) for x in df.columns.tolist()]
        if primeira == cab:
            df = df.iloc[1:].copy()

    if col_nota:
        df[col_nota] = df[col_nota].astype(str).str.strip()

    if col_nome:
        df[col_nome] = df[col_nome].astype(str).str.strip()

    if col_data:
        df[col_data] = pd.to_datetime(df[col_data], errors="coerce")

    return df, col_nota, col_data, col_nome, header_idx

def listar_abas(path_xlsx):
    xls = pd.ExcelFile(path_xlsx)
    return xls.sheet_names

def baixar_do_drive(file_id, destino):
    if not file_id:
        return False, "DRIVE_FILE_ID não informado."
    if gdown is None:
        return False, "Biblioteca gdown não instalada. Adicione gdown no requirements.txt."

    url = f"https://drive.google.com/uc?id={file_id}"
    try:
        gdown.download(url, destino, quiet=False)
        return True, f"Arquivo baixado com sucesso para: {destino}"
    except Exception as e:
        return False, f"Erro ao baixar do Drive: {e}"

def validar_xlsx(path_xlsx):
    p = Path(path_xlsx)
    if not p.exists():
        return False, "Arquivo não encontrado."
    if p.suffix.lower() not in [".xlsx", ".xlsm", ".xls"]:
        return False, "Arquivo inválido. Informe um Excel."
    return True, "OK"

def df_para_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="resultado")
    output.seek(0)
    return output.getvalue()

# =========================================================
# ESTADO
# =========================================================

if "logado" not in st.session_state:
    st.session_state.logado = False

if "usuario_logado" not in st.session_state:
    st.session_state.usuario_logado = ""

if "xlsx_path" not in st.session_state:
    st.session_state.xlsx_path = DEFAULT_XLSX_PATH

if "drive_file_id" not in st.session_state:
    st.session_state.drive_file_id = DEFAULT_DRIVE_FILE_ID

if "usar_drive" not in st.session_state:
    st.session_state.usar_drive = False

# =========================================================
# CACHE
# =========================================================

@st.cache_data(show_spinner=False)
def carregar_abas_cache(path_xlsx):
    return listar_abas(path_xlsx)

@st.cache_data(show_spinner=False)
def carregar_df_cache(path_xlsx, sheet_name):
    return preparar_dataframe(path_xlsx, sheet_name)

# =========================================================
# LOGIN
# =========================================================

def tela_login():
    st.markdown('<div class="login-wrap">', unsafe_allow_html=True)
    st.markdown('<div class="login-box">', unsafe_allow_html=True)
    st.markdown('<div class="login-title">🔐 Buscar Cadastro</div>', unsafe_allow_html=True)
    st.markdown('<div class="login-subtitle">Entre com seu usuário e senha para acessar o sistema.</div>', unsafe_allow_html=True)

    usuario = st.text_input("Usuário", key="login_usuario")
    senha = st.text_input("Senha", type="password", key="login_senha")

    if st.button("Entrar", use_container_width=True):
        if usuario in USUARIOS and USUARIOS[usuario] == senha:
            st.session_state.logado = True
            st.session_state.usuario_logado = usuario
            st.rerun()
        else:
            st.error("Usuário ou senha inválidos.")

    st.markdown('<div class="small-note">Acesso inicial do sistema</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

# =========================================================
# SIDEBAR / CONFIGURAÇÕES
# =========================================================

def painel_configuracoes():
    with st.sidebar:
        st.markdown("## ⚙️ Configurações")
        st.write(f"Usuário logado: **{st.session_state.usuario_logado}**")

        with st.expander("Arquivo XLSX", expanded=False):
            usar_drive = st.checkbox(
                "Baixar do Google Drive",
                value=st.session_state.usar_drive
            )

            xlsx_path = st.text_input(
                "Caminho do XLSX",
                value=st.session_state.xlsx_path,
                help="Deixe padrão se quiser usar o mesmo arquivo local do app."
            )

            drive_file_id = st.text_input(
                "Drive File ID",
                value=st.session_state.drive_file_id,
                help="Use se quiser baixar a planilha direto do Google Drive."
            )

            col_a, col_b = st.columns(2)

            with col_a:
                if st.button("Salvar config", use_container_width=True):
                    st.session_state.usar_drive = usar_drive
                    st.session_state.xlsx_path = xlsx_path.strip() or DEFAULT_XLSX_PATH
                    st.session_state.drive_file_id = drive_file_id.strip()
                    st.success("Configurações salvas.")
                    st.rerun()

            with col_b:
                if st.button("Restaurar padrão", use_container_width=True):
                    st.session_state.usar_drive = False
                    st.session_state.xlsx_path = DEFAULT_XLSX_PATH
                    st.session_state.drive_file_id = DEFAULT_DRIVE_FILE_ID
                    st.success("Padrão restaurado.")
                    st.rerun()

            if st.button("Baixar/atualizar do Drive", use_container_width=True):
                destino = st.session_state.xlsx_path
                ok, msg = baixar_do_drive(st.session_state.drive_file_id, destino)
                if ok:
                    carregar_abas_cache.clear()
                    carregar_df_cache.clear()
                    st.success(msg)
                else:
                    st.error(msg)

        ok, msg = validar_xlsx(st.session_state.xlsx_path)
        if ok:
            st.success(f"Arquivo ativo: {st.session_state.xlsx_path}")
        else:
            st.warning(f"Arquivo ativo: {st.session_state.xlsx_path}")
            st.caption(msg)

        if st.button("Sair", use_container_width=True):
            st.session_state.logado = False
            st.session_state.usuario_logado = ""
            st.rerun()

# =========================================================
# APP PRINCIPAL
# =========================================================

def app():
    painel_configuracoes()

    st.title("🔎 Buscar Cadastro de Eletricista")

    caminho = st.session_state.xlsx_path
    ok, msg = validar_xlsx(caminho)
    if not ok:
        st.error(f"Não foi possível abrir o XLSX. Motivo: {msg}")
        st.info("No Streamlit Cloud, use o painel lateral para informar o caminho padrão ou baixar do Google Drive.")
        st.stop()

    try:
        abas = carregar_abas_cache(caminho)
    except Exception as e:
        st.error(f"Erro ao listar abas da planilha: {e}")
        st.stop()

    if not abas:
        st.error("Nenhuma aba encontrada no arquivo.")
        st.stop()

    abas_opcoes = abas + ["TODOS"]

    top1, top2 = st.columns([1.1, 2.2])

    with top1:
        aba_escolhida = st.selectbox("Selecione a aba", abas_opcoes)

    with top2:
        st.info("O seletor mostra automaticamente as abas existentes na planilha.")

    if aba_escolhida == "TODOS":
        dfs = []
        for aba in abas:
            try:
                df_tmp, _, _, _, _ = carregar_df_cache(caminho, aba)
                if len(df_tmp) > 0:
                    df_tmp = df_tmp.copy()
                    df_tmp["ABA_ORIGEM"] = aba
                    dfs.append(df_tmp)
            except Exception:
                continue

        if not dfs:
            st.error("Não foi possível carregar nenhuma aba válida.")
            st.stop()

        df = pd.concat(dfs, ignore_index=True)

        col_nota = localizar_coluna(df, ["NOTA", "NF", "NUMERO DA NOTA", "NUMERO NOTA", "N DA NOTA", "N_NOTA"])
        col_data = localizar_coluna(df, ["DATA", "DATA DA NOTA", "DT", "EMISSAO", "DATA EMISSAO"])
        col_nome = localizar_coluna(df, ["ELETRICISTA", "NOME", "NOME DO ELETRICISTA", "PRESTADOR", "COLABORADOR", "NOME COMPLETO"])
        header_idx = "múltiplas abas"
    else:
        try:
            df, col_nota, col_data, col_nome, header_idx = carregar_df_cache(caminho, aba_escolhida)
        except Exception as e:
            st.error(f"Erro ao carregar a aba '{aba_escolhida}': {e}")
            st.stop()

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Linhas carregadas", len(df))
    m2.metric("Cabeçalho detectado", str(header_idx + 1) if isinstance(header_idx, int) else str(header_idx))
    m3.metric("Coluna Nota", col_nota if col_nota else "não identificada")
    m4.metric("Coluna Nome", col_nome if col_nome else "não identificada")

    if not col_nome:
        st.warning("Não consegui identificar automaticamente a coluna do nome/eletricista.")
    if not col_nota:
        st.warning("Não consegui identificar automaticamente a coluna da nota.")

    st.divider()

    c1, c2 = st.columns(2)

    with c1:
        busca_nome = st.text_input("Digite o nome do eletricista")
        busca_exata = st.checkbox("Busca exata do nome", value=False)

    with c2:
        busca_nota = st.text_input("Digite o número da nota")

    usar_data = st.checkbox("Filtrar por data", value=False)
    busca_data = None
    if usar_data:
        busca_data = st.date_input("Selecione a data")

    st.divider()

    st.subheader("Busca em massa por nomes")
    lista_nomes = st.text_area(
        "Cole vários nomes de eletricistas (1 por linha)",
        height=160,
        help="Exemplo: um nome por linha. O sistema vai buscar todos de uma vez."
    )

    b1, b2 = st.columns(2)
    pesquisar = b1.button("🔎 Pesquisar", use_container_width=True)
    limpar = b2.button("Limpar filtros", use_container_width=True)

    if limpar:
        st.rerun()

    resultado = df.copy()

    if pesquisar:
        if busca_nome and col_nome:
            serie_nome = resultado[col_nome].fillna("").astype(str)

            if busca_exata:
                nome_ref = normalizar_texto(busca_nome)
                resultado = resultado[serie_nome.map(normalizar_texto) == nome_ref]
            else:
                resultado = resultado[safe_str_contains(serie_nome, busca_nome)]

        if busca_nota and col_nota:
            serie_nota = resultado[col_nota].fillna("").astype(str).str.strip()
            nota_ref = str(busca_nota).strip()
            resultado = resultado[serie_nota.str.contains(re.escape(nota_ref), na=False)]

        if usar_data and busca_data and col_data:
            serie_data = pd.to_datetime(resultado[col_data], errors="coerce")
            resultado = resultado[serie_data.dt.date == busca_data]

        # BUSCA EM MASSA POR NOMES
        if lista_nomes.strip() and col_nome:
            nomes = [linha.strip() for linha in lista_nomes.splitlines() if linha.strip()]
            nomes_normalizados = {normalizar_texto(nome) for nome in nomes}

            serie_nome_norm = resultado[col_nome].fillna("").astype(str).map(normalizar_texto)
            resultado = resultado[serie_nome_norm.isin(nomes_normalizados)]

    st.divider()

    if pesquisar:
        if resultado.empty:
            st.warning("Nenhum registro encontrado.")
        else:
            st.success(f"{len(resultado)} registro(s) encontrado(s).")

            with st.container():
                st.markdown('<div class="result-box">Resultado da pesquisa carregado com sucesso.</div>', unsafe_allow_html=True)

            st.dataframe(resultado, use_container_width=True, height=420)

            excel_bytes = df_para_excel_bytes(resultado)
            st.download_button(
                "⬇️ Baixar resultado em Excel",
                data=excel_bytes,
                file_name="resultado_busca_cadastro.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    with st.expander("📄 Visualizar base completa", expanded=False):
        st.dataframe(df, use_container_width=True, height=350)
        st.caption("Colunas encontradas na base:")
        st.write(list(df.columns))

# =========================================================
# MAIN
# =========================================================

def main():
    if not st.session_state.logado:
        tela_login()
    else:
        app()

if __name__ == "__main__":
    main()
