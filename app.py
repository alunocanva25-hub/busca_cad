import os
import io
import re
import unicodedata
from pathlib import Path

import pandas as pd
import streamlit as st
from dotenv import load_dotenv

try:
    import gdown
except ImportError:
    gdown = None

# =========================================================
# CONFIG INICIAL
# =========================================================

load_dotenv()

st.set_page_config(
    page_title="Buscar Cadastro",
    page_icon="🔎",
    layout="wide"
)

DEFAULT_XLSX_PATH = os.getenv("XLSX_PATH", "dados.xlsx")
DEFAULT_DRIVE_FILE_ID = os.getenv("DRIVE_FILE_ID", "")

USUARIOS = {
    "admin": "123",
    "usuario1": "123",
    "usuario2": "123",
    "usuario3": "123",
    "usuario4": "123",
}

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

def detectar_cabecalho(path_xlsx, sheet_name, max_linhas=10):
    """
    Tenta achar automaticamente a linha do cabeçalho.
    Dá prioridade para linhas que contenham algo parecido com:
    NOTA / DATA / ELETRICISTA / NOME
    """
    bruto = pd.read_excel(path_xlsx, sheet_name=sheet_name, header=None, nrows=max_linhas)
    melhor_idx = 2  # fallback: linha 3 da planilha => índice 2
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

    # match exato
    for cand in candidatos:
        cand_slug = slug_coluna(cand)
        if cand_slug in mapa:
            return mapa[cand_slug]

    # match parcial
    for slug, original in mapa.items():
        for cand in candidatos:
            cand_slug = slug_coluna(cand)
            if cand_slug in slug:
                return original

    return None

def limpar_dataframe(df):
    # remove colunas totalmente vazias
    df = df.dropna(axis=1, how="all").copy()

    # renomear colunas vazias/unnamed se necessário
    novas_cols = []
    for i, col in enumerate(df.columns):
        col_str = str(col).strip()
        if not col_str or col_str.upper().startswith("UNNAMED"):
            col_str = f"COLUNA_{i+1}"
        novas_cols.append(col_str)
    df.columns = novas_cols

    # remove linhas totalmente vazias
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

    # Se vier uma primeira linha repetindo o cabeçalho, remove
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
        try:
            df[col_data] = pd.to_datetime(df[col_data], errors="coerce")
        except Exception:
            pass

    return df, col_nota, col_data, col_nome, header_idx

def listar_abas(path_xlsx):
    xls = pd.ExcelFile(path_xlsx)
    return xls.sheet_names

def baixar_do_drive(file_id, destino):
    if not file_id:
        return False, "DRIVE_FILE_ID não informado."
    if gdown is None:
        return False, "Biblioteca gdown não instalada. Instale com: pip install gdown"

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
        return False, "O arquivo informado não parece ser um Excel válido."
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
# LOGIN
# =========================================================

def tela_login():
    st.title("🔐 Buscar Cadastro")
    st.write("Faça login para acessar o sistema.")

    c1, c2, c3 = st.columns([1, 1.2, 1])

    with c2:
        usuario = st.text_input("Usuário")
        senha = st.text_input("Senha", type="password")

        if st.button("Entrar", use_container_width=True):
            if usuario in USUARIOS and USUARIOS[usuario] == senha:
                st.session_state.logado = True
                st.session_state.usuario_logado = usuario
                st.success("Login realizado com sucesso.")
                st.rerun()
            else:
                st.error("Usuário ou senha inválidos.")

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
                help="Pode deixar o padrão, por exemplo: dados.xlsx"
            )

            drive_file_id = st.text_input(
                "Drive File ID",
                value=st.session_state.drive_file_id,
                help="Preencha somente se quiser baixar o arquivo do Google Drive"
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
                    st.success("Configuração padrão restaurada.")
                    st.rerun()

            if st.button("Baixar/agualizar do Drive", use_container_width=True):
                destino = st.session_state.xlsx_path
                ok, msg = baixar_do_drive(
                    st.session_state.drive_file_id,
                    destino
                )
                if ok:
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

@st.cache_data(show_spinner=False)
def carregar_abas_cache(path_xlsx):
    return listar_abas(path_xlsx)

@st.cache_data(show_spinner=False)
def carregar_df_cache(path_xlsx, sheet_name):
    return preparar_dataframe(path_xlsx, sheet_name)

def app():
    painel_configuracoes()

    st.title("🔎 Buscar Cadastro de Eletricista")

    caminho = st.session_state.xlsx_path
    ok, msg = validar_xlsx(caminho)
    if not ok:
        st.error(f"Não foi possível abrir o XLSX. Motivo: {msg}")
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

    col_top1, col_top2 = st.columns([1.2, 2.2])

    with col_top1:
        aba_escolhida = st.selectbox(
            "Selecione a aba",
            abas_opcoes,
            index=0
        )

    with col_top2:
        st.info("O seletor identifica automaticamente as abas existentes na planilha.")

    if aba_escolhida == "TODOS":
        dfs = []
        mapa_info = []

        for aba in abas:
            try:
                df_tmp, col_nota_tmp, col_data_tmp, col_nome_tmp, header_idx_tmp = carregar_df_cache(caminho, aba)
                if len(df_tmp) > 0:
                    df_tmp = df_tmp.copy()
                    df_tmp["ABA_ORIGEM"] = aba
                    dfs.append(df_tmp)
                    mapa_info.append((aba, col_nota_tmp, col_data_tmp, col_nome_tmp, header_idx_tmp))
            except Exception:
                continue

        if not dfs:
            st.error("Não foi possível carregar nenhuma aba válida.")
            st.stop()

        df = pd.concat(dfs, ignore_index=True)

        # detecta colunas principais no consolidado
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

    info_cols = st.columns(4)
    info_cols[0].metric("Linhas carregadas", len(df))
    info_cols[1].metric("Cabeçalho detectado", str(header_idx + 1) if isinstance(header_idx, int) else str(header_idx))
    info_cols[2].metric("Coluna Nota", col_nota if col_nota else "não identificada")
    info_cols[3].metric("Coluna Nome", col_nome if col_nome else "não identificada")

    if not col_nome:
        st.warning(
            "Não consegui identificar automaticamente a coluna do eletricista/nome. "
            "Confira os cabeçalhos da planilha."
        )

    if not col_nota:
        st.warning(
            "Não consegui identificar automaticamente a coluna da nota. "
            "Confira os cabeçalhos da planilha."
        )

    st.divider()

    col1, col2 = st.columns(2)

    with col1:
        busca_nome = st.text_input("Digite o nome do eletricista")
        busca_exata = st.checkbox("Busca exata do nome", value=False)

    with col2:
        busca_nota = st.text_input("Digite o número da nota")

    usar_data = st.checkbox("Filtrar por data", value=False)
    busca_data = None
    if usar_data:
        busca_data = st.date_input("Selecione a data")

    st.divider()

    st.subheader("Busca em massa")
    lista_notas = st.text_area(
        "Cole vários números de nota (1 por linha)",
        height=140
    )

    col_btn1, col_btn2 = st.columns([1, 1])
    pesquisar = col_btn1.button("🔎 Pesquisar", use_container_width=True)
    limpar = col_btn2.button("Limpar filtros", use_container_width=True)

    if limpar:
        st.rerun()

    resultado = df.copy()

    if pesquisar:
        # filtro por nome
        if busca_nome and col_nome:
            serie_nome = resultado[col_nome].fillna("").astype(str)

            if busca_exata:
                nome_ref = normalizar_texto(busca_nome)
                resultado = resultado[
                    serie_nome.map(normalizar_texto) == nome_ref
                ]
            else:
                resultado = resultado[
                    safe_str_contains(serie_nome, busca_nome)
                ]

        # filtro por nota
        if busca_nota and col_nota:
            serie_nota = resultado[col_nota].fillna("").astype(str).str.strip()
            nota_ref = str(busca_nota).strip()
            resultado = resultado[
                serie_nota.str.contains(re.escape(nota_ref), na=False)
            ]

        # filtro por data
        if usar_data and busca_data and col_data:
            serie_data = pd.to_datetime(resultado[col_data], errors="coerce")
            resultado = resultado[serie_data.dt.date == busca_data]

        # busca em massa
        if lista_notas.strip() and col_nota:
            notas = [
                linha.strip()
                for linha in lista_notas.splitlines()
                if linha.strip()
            ]
            notas_set = set(notas)
            serie_nota = resultado[col_nota].fillna("").astype(str).str.strip()
            resultado = resultado[serie_nota.isin(notas_set)]

    st.divider()

    if pesquisar:
        if resultado.empty:
            st.warning("Nenhum registro encontrado.")
        else:
            st.success(f"{len(resultado)} registro(s) encontrado(s).")

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
