import streamlit as st
import pandas as pd
from datetime import datetime
import io
import base64
import os

st.set_page_config(page_title="Inventário de Equipamentos", layout="wide")

st.markdown("""
<style>
    .stApp { background-color: #f5f5f5; }
    .titulo-inventario { font-size: 22px; font-weight: bold; color: #1a1a1a; margin-bottom: 10px; }
    .badge-data {
        background-color: #c8f0c8; border: 1px solid #5cb85c;
        border-radius: 6px; padding: 6px 14px; font-size: 13px;
        color: #2d6a2d; font-style: italic; text-align: center;
    }
    .secao-header {
        background-color: #5cb85c; color: white; font-weight: bold;
        padding: 8px 14px; border-radius: 6px; font-size: 15px; margin-bottom: 4px;
    }
    .label-nome {
        background-color: #fde8d8; border: 1.5px solid #e0a080;
        border-radius: 6px; padding: 8px 14px; font-weight: bold;
        font-size: 15px; color: #333; margin-bottom: 4px;
    }
    .label-uo {
        background-color: #d4edda; border: 1.5px solid #5cb85c;
        border-radius: 6px; padding: 8px 14px; font-weight: bold;
        font-size: 15px; color: #333; margin-bottom: 4px;
    }
    .num-item { font-weight: bold; font-size: 16px; color: #333; padding-top: 8px; }
    .stButton > button {
        background-color: #1565C0; color: white; border-radius: 8px;
        padding: 10px 28px; font-size: 15px; font-weight: bold; border: none;
    }
    .stButton > button:hover { background-color: #0d47a1; color: white; }
    div[data-testid="stSelectbox"] label,
    div[data-testid="stTextInput"] label { display: none; }
    hr { margin: 8px 0; }
</style>
""", unsafe_allow_html=True)

TIPO_CELULAR  = ["Celular obra 360°", "Celular padrão"]
ARQUIVO_EXCEL = "inventario_equipamentos.xlsx"
PLANILHA_UO   = "bases/Planilha para atualização - Notebooks e Desktops.xlsx"
SENHA_CORRETA = "Sucesso_obra_PEP"

try:
    GITHUB_TOKEN  = st.secrets["GITHUB_TOKEN"]
    GITHUB_REPO   = st.secrets["GITHUB_REPO"]
    ARQUIVO_EXCEL = st.secrets.get("ARQUIVO_EXCEL", ARQUIVO_EXCEL)
    BRANCH        = st.secrets.get("BRANCH", "main")
    MODO_GITHUB   = True
except Exception:
    MODO_GITHUB   = False

@st.cache_data(show_spinner=False)
def carregar_lista_uo() -> list:
    try:
        if MODO_GITHUB:
            from github import Github
            repo     = Github(GITHUB_TOKEN).get_repo(GITHUB_REPO)
            conteudo = repo.get_contents(PLANILHA_UO, ref=BRANCH)
            buffer   = io.BytesIO(base64.b64decode(conteudo.content))
            df_uo    = pd.read_excel(buffer, sheet_name="UO")
        else:
            df_uo = pd.read_excel(PLANILHA_UO, sheet_name="UO")
        lista = df_uo.iloc[:, 0].dropna().drop_duplicates().astype(str).tolist()
        lista = [v.strip() for v in lista if v.strip()]
        lista.sort()
        return lista
    except Exception as e:
        st.warning(f"⚠️ Não foi possível carregar a lista de UOs: {e}")
        return []

LISTA_UO = carregar_lista_uo()

def ler_excel_local() -> pd.DataFrame:
    if os.path.exists(ARQUIVO_EXCEL):
        return pd.read_excel(ARQUIVO_EXCEL)
    return pd.DataFrame()

def salvar_excel_local(dados: dict):
    df_existente = ler_excel_local()
    df_final     = pd.concat([df_existente, pd.DataFrame([dados])], ignore_index=True)
    df_final.to_excel(ARQUIVO_EXCEL, index=False)

def get_repo():
    from github import Github
    return Github(GITHUB_TOKEN).get_repo(GITHUB_REPO)

def ler_excel_github() -> pd.DataFrame:
    try:
        repo     = get_repo()
        conteudo = repo.get_contents(ARQUIVO_EXCEL, ref=BRANCH)
        buffer   = io.BytesIO(base64.b64decode(conteudo.content))
        return pd.read_excel(buffer)
    except Exception:
        return pd.DataFrame()

def salvar_excel_github(dados: dict):
    repo         = get_repo()
    df_existente = ler_excel_github()
    df_final     = pd.concat([df_existente, pd.DataFrame([dados])], ignore_index=True)
    buffer = io.BytesIO()
    df_final.to_excel(buffer, index=False)
    buffer.seek(0)
    conteudo_bytes = buffer.read()
    msg = f"Novo registro: {dados.get('Nome','')} | {dados.get('UO','')} | {dados.get('Data/Hora','')}"
    try:
        arq = repo.get_contents(ARQUIVO_EXCEL, ref=BRANCH)
        repo.update_file(ARQUIVO_EXCEL, msg, conteudo_bytes, arq.sha, branch=BRANCH)
    except Exception:
        repo.create_file(ARQUIVO_EXCEL, "Criação da planilha de inventário", conteudo_bytes, branch=BRANCH)

def ler_excel() -> pd.DataFrame:
    return ler_excel_github() if MODO_GITHUB else ler_excel_local()

def salvar_excel(dados: dict):
    if MODO_GITHUB:
        salvar_excel_github(dados)
    else:
        salvar_excel_local(dados)

def init_state():
    defaults = {
        "qt_desktop"      : 1,
        "qt_notebook"     : 1,
        "qt_celular"      : 1,
        "enviado"         : False,
        "senha_ok"        : False,
        "senha_digitada"  : "",
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val

init_state()

def resetar_formulario():
    keys_para_limpar = [k for k in list(st.session_state.keys()) if k not in ("enviado", "senha_ok", "senha_digitada")]
    for k in keys_para_limpar:
        del st.session_state[k]
    st.session_state.enviado        = False
    st.session_state.qt_desktop     = 1
    st.session_state.qt_notebook    = 1
    st.session_state.qt_celular     = 1

col_titulo, _, col_data = st.columns([4, 3, 2])
with col_titulo:
    st.markdown('<div class="titulo-inventario">📋 INVENTÁRIO DE EQUIPAMENTOS</div>', unsafe_allow_html=True)
    if MODO_GITHUB:
        st.caption("☁️ Modo cloud — salvando no GitHub")
    else:
        st.caption("💻 Modo local — salvando em arquivo .xlsx na pasta do projeto")

with col_data:
    data_hoje = datetime.now().strftime("%d/%m/%Y %H:%M")
    st.markdown(f'<div class="badge-data">📅 Data automática: <b>{data_hoje}</b></div>', unsafe_allow_html=True)

st.markdown("---")

col_nome, _, col_uo, col_obs = st.columns([3, 0.5, 3, 3])

with col_nome:
    st.markdown('<div class="label-nome">👤 Nome</div>', unsafe_allow_html=True)
    nome = st.text_input("nome_input", placeholder="Digite seu nome completo", key="nome", label_visibility="collapsed")

with col_uo:
    st.markdown('<div class="label-uo">🏢 UO</div>', unsafe_allow_html=True)
    uo = st.selectbox("uo_select", options=["Selecione..."] + LISTA_UO, key="uo", label_visibility="collapsed")

with col_obs:
    st.markdown("""
    <div style='font-size:12px; color:#555; padding-top:36px;'>
    ℹ️ Digite para pesquisar sua UO na lista suspensa ao lado.
    </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

col_d, col_n, col_c = st.columns(3)

with col_d:
    st.markdown('<div class="secao-header">🖥️ Quantitativo desktop</div>', unsafe_allow_html=True)
    qt_desktop = st.selectbox("qt_desktop_sel", list(range(1, 21)), key="qt_desktop", label_visibility="collapsed")

with col_n:
    st.markdown('<div class="secao-header">💻 Quantitativo notebook</div>', unsafe_allow_html=True)
    qt_notebook = st.selectbox("qt_notebook_sel", list(range(1, 21)), key="qt_notebook", label_visibility="collapsed")

with col_c:
    st.markdown('<div class="secao-header">📱 Quantitativo celulares</div>', unsafe_allow_html=True)
    qt_celular = st.selectbox("qt_celular_sel", list(range(1, 21)), key="qt_celular", label_visibility="collapsed")

st.markdown("<br>", unsafe_allow_html=True)

col_d2, col_n2, col_c2 = st.columns(3)

patrimonios_desktop  = []
patrimonios_notebook = []
patrimonios_cel_tipo = []
patrimonios_cel_num  = []

with col_d2:
    st.markdown("**Patrimônios Desktop:**")
    for i in range(1, st.session_state.qt_desktop + 1):
        c1, c2 = st.columns([0.3, 2.7])
        with c1:
            st.markdown(f'<div class="num-item">{i} -</div>', unsafe_allow_html=True)
        with c2:
            val = st.text_input(f"desk_{i}", placeholder="Inserir patrimônio", key=f"desktop_{i}", label_visibility="collapsed")
            patrimonios_desktop.append(val)

with col_n2:
    st.markdown("**Patrimônios Notebook:**")
    for i in range(1, st.session_state.qt_notebook + 1):
        c1, c2 = st.columns([0.3, 2.7])
        with c1:
            st.markdown(f'<div class="num-item">{i} -</div>', unsafe_allow_html=True)
        with c2:
            val = st.text_input(f"note_{i}", placeholder="Inserir patrimônio", key=f"notebook_{i}", label_visibility="collapsed")
            patrimonios_notebook.append(val)

with col_c2:
    st.markdown("**Patrimônios Celular:**")
    for i in range(1, st.session_state.qt_celular + 1):
        c1, c2 = st.columns([0.3, 2.7])
        with c1:
            st.markdown(f'<div class="num-item">{i} -</div>', unsafe_allow_html=True)
        with c2:
            tipo = st.selectbox(f"cel_tipo_{i}", TIPO_CELULAR, key=f"celular_tipo_{i}", label_visibility="collapsed")
            patrimonios_cel_tipo.append(tipo)
        _, c3 = st.columns([0.3, 2.7])
        with c3:
            num = st.text_input(f"cel_num_{i}", placeholder="Inserir patrimônio", key=f"celular_num_{i}", label_visibility="collapsed")
            patrimonios_cel_num.append(num)

st.markdown("<br>", unsafe_allow_html=True)

_, col_btn1, col_btn2 = st.columns([6, 1, 1.5])

with col_btn1:
    if st.button("🔄 Resetar"):
        resetar_formulario()
        st.rerun()

with col_btn2:
    enviar = st.button("📤 Enviar")

if enviar:
    erros = []
    if not nome.strip():
        erros.append("⚠️ O campo **Nome** é obrigatório.")
    if uo == "Selecione...":
        erros.append("⚠️ Selecione uma **UO**.")

    if erros:
        for e in erros:
            st.error(e)
    else:
        dados = {
            "Data/Hora"  : datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "Nome"       : nome.strip(),
            "UO"         : uo,
            "Qt_Desktop" : st.session_state.qt_desktop,
            "Qt_Notebook": st.session_state.qt_notebook,
            "Qt_Celular" : st.session_state.qt_celular,
        }
        for i, p in enumerate(patrimonios_desktop, 1):
            dados[f"Desktop_{i}_Patrimônio"] = p.strip() if p else ""
        for i, p in enumerate(patrimonios_notebook, 1):
            dados[f"Notebook_{i}_Patrimônio"] = p.strip() if p else ""
        for i, (tipo, num) in enumerate(zip(patrimonios_cel_tipo, patrimonios_cel_num), 1):
            dados[f"Celular_{i}_Tipo"]       = tipo
            dados[f"Celular_{i}_Patrimônio"] = num.strip() if num else ""

        modo_label = "GitHub ☁️" if MODO_GITHUB else "arquivo local 💻"
        with st.spinner(f"💾 Salvando em {modo_label}..."):
            try:
                salvar_excel(dados)
                st.success(f"✅ Inventário de **{nome}** ({uo}) salvo com sucesso!")
                st.balloons()
                st.session_state.enviado = True
            except Exception as ex:
                st.error(f"❌ Erro ao salvar: {ex}")

if st.session_state.enviado:
    with st.expander("📊 Ver todos os registros salvos"):
        st.markdown("🔒 **Acesso restrito** — insira a senha para visualizar e baixar a planilha.")

        senha_input = st.text_input(
            "Senha de acesso",
            type="password",
            placeholder="Digite a senha...",
            key="campo_senha"
        )

        if st.button("🔓 Acessar"):
            if senha_input == SENHA_CORRETA:
                st.session_state.senha_ok = True
            else:
                st.session_state.senha_ok = False
                st.error("❌ Senha incorreta. Tente novamente.")

        if st.session_state.senha_ok:
            df_view = ler_excel()
            if not df_view.empty:
                st.dataframe(df_view, use_container_width=True)
                buffer_dl = io.BytesIO()
                df_view.to_excel(buffer_dl, index=False)
                buffer_dl.seek(0)
                st.download_button(
                    label="⬇️ Baixar planilha Excel",
                    data=buffer_dl,
                    file_name=ARQUIVO_EXCEL,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("Nenhum registro encontrado ainda.")