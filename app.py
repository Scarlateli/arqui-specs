"""
app.py — Arqui Specs · Assistente de Memorial Descritivo
Estudio GL · Gabriela Lendecker · 2026

Estrutura:
  - Auth: senha simples via st.secrets (SCRUM-58)
  - Sidebar: projeto, upload/download, info do projeto
  - Layout 2 colunas: chat (esq) + preview planilha (dir) (SCRUM-54)
  - Streaming + tool use (SCRUM-49, 51, 60)
  - Auto-save draft (SCRUM-62)
  - Tratamento de erros amigável (SCRUM-55)
"""

import streamlit as st
import openpyxl
import datetime
from io import BytesIO

from core.spreadsheet import (
    create_template,
    load_workbook_from_bytes,
    save_workbook_to_bytes,
    get_preview_dataframe,
    style_preview_dataframe,
    get_project_info,
    update_project_info,
    get_sheet_data,
)
from core.claude_client import (
    stream_response,
    build_user_message,
    MODEL,
)

# ─── Configuração da página ──────────────────────────────────────────────────

st.set_page_config(
    page_title="Arqui Specs · Estudio GL",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CSS customizado ─────────────────────────────────────────────────────────

st.markdown("""
<style>
/* Sidebar mais estreita e elegante */
[data-testid="stSidebar"] { min-width: 280px; max-width: 300px; }

/* Título da marca */
.brand-title {
    font-size: 22px;
    font-weight: 700;
    color: #1C1C1C;
    letter-spacing: 0.05em;
    margin-bottom: 2px;
}
.brand-sub {
    font-size: 11px;
    color: #8B6F47;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    margin-bottom: 16px;
}

/* Balão de chat do assistente com borda esquerda bronze */
[data-testid="stChatMessage"]:has([data-testid="stChatMessageContent"]) {
    border-left: 3px solid #8B6F47;
    padding-left: 8px;
}

/* Badge de status na preview */
.status-aprovado { background:#4A7A5A; color:#fff; padding:2px 8px; border-radius:4px; font-size:11px; font-weight:600; }
.status-decidir  { background:#C4904A; color:#fff; padding:2px 8px; border-radius:4px; font-size:11px; font-weight:600; }
.status-verificar{ background:#7A6B4A; color:#fff; padding:2px 8px; border-radius:4px; font-size:11px; font-weight:600; }

/* Preview: container com scroll */
.preview-container { max-height: 75vh; overflow-y: auto; }

/* Rodapé */
footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)


# ─── Auth (SCRUM-58) ─────────────────────────────────────────────────────────

def check_auth() -> bool:
    """
    Autenticação por senha simples.
    Senha definida em .streamlit/secrets.toml como APP_PASSWORD.
    Retorna True se autenticado.
    """
    # Se não há APP_PASSWORD configurada, permite acesso (dev local)
    try:
        expected = st.secrets["APP_PASSWORD"]
    except (KeyError, FileNotFoundError):
        return True  # Sem senha configurada = acesso liberado (modo dev)

    if st.session_state.get("authenticated"):
        return True

    with st.container():
        st.markdown("## 🔐 Arqui Specs")
        st.markdown("Ferramenta interna · Estudio GL")
        password = st.text_input("Senha de acesso", type="password", key="pw_input")
        if st.button("Entrar", type="primary"):
            if password == expected:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Senha incorreta.")
    return False


def get_api_key() -> str | None:
    """Retorna a API key da Anthropic dos secrets, ou None."""
    try:
        return st.secrets["ANTHROPIC_API_KEY"]
    except (KeyError, FileNotFoundError):
        return None


# ─── Estado da sessão ────────────────────────────────────────────────────────

def init_session():
    """Inicializa o estado da sessão na primeira execução."""
    if "workbook_bytes" not in st.session_state:
        # Criar planilha template nova
        st.session_state.workbook_bytes = create_template().read()

    if "display_messages" not in st.session_state:
        st.session_state.display_messages = []  # [{role, content}]

    if "api_messages" not in st.session_state:
        st.session_state.api_messages = []  # Histórico completo para a API

    if "project_name" not in st.session_state:
        st.session_state.project_name = "Memorial Descritivo"

    if "last_saved" not in st.session_state:
        st.session_state.last_saved = None

    if "draft_count" not in st.session_state:
        st.session_state.draft_count = 0


def get_workbook() -> openpyxl.Workbook:
    """Carrega o workbook atual do session_state."""
    return load_workbook_from_bytes(st.session_state.workbook_bytes)


def save_workbook(wb: openpyxl.Workbook):
    """Salva o workbook atualizado no session_state."""
    st.session_state.workbook_bytes = save_workbook_to_bytes(wb)
    st.session_state.last_saved = datetime.datetime.now()


# ─── Sidebar ─────────────────────────────────────────────────────────────────

def render_sidebar():
    with st.sidebar:
        # Logo / marca
        st.markdown('<div class="brand-title">Arqui Specs</div>', unsafe_allow_html=True)
        st.markdown('<div class="brand-sub">Estudio GL · Ferramenta Interna</div>', unsafe_allow_html=True)
        st.divider()

        # Nome do projeto
        project_name = st.text_input(
            "Nome do projeto",
            value=st.session_state.project_name,
            placeholder="Ex: Apto Alencar — Higienópolis",
            key="project_name_input",
        )
        if project_name != st.session_state.project_name:
            st.session_state.project_name = project_name
            wb = get_workbook()
            update_project_info(wb, nome=project_name)
            save_workbook(wb)

        st.divider()

        # ── Upload de planilha existente ──────────────────────────────────
        st.markdown("**Planilha**")
        uploaded = st.file_uploader(
            "Carregar .xlsx existente",
            type=["xlsx"],
            label_visibility="collapsed",
            key="uploader",
        )
        if uploaded:
            # Só processa se for um arquivo NOVO (file_id diferente do último).
            # Sem isso, cada st.rerun() re-lê o uploader e sobrescreve
            # as modificações feitas pelo chat.
            if uploaded.file_id != st.session_state.get("last_uploaded_id"):
                try:
                    wb_test = load_workbook_from_bytes(uploaded.getvalue())
                    st.session_state.workbook_bytes = uploaded.getvalue()
                    st.session_state.last_uploaded_id = uploaded.file_id
                    info = get_project_info(wb_test)
                    st.session_state.project_name = info["nome"]
                    st.session_state.display_messages = []
                    st.session_state.api_messages = []
                    st.success("Planilha carregada!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Arquivo inválido: {e}")

        # ── Download ──────────────────────────────────────────────────────
        wb = get_workbook()
        info = get_project_info(wb)
        filename = f"Memorial_{info['nome'].replace(' ', '_').replace('/', '-')}.xlsx"

        st.download_button(
            label="⬇️ Baixar .xlsx",
            data=st.session_state.workbook_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        # Auto-save info (SCRUM-62)
        if st.session_state.last_saved:
            ts = st.session_state.last_saved.strftime("%H:%M:%S")
            st.caption(f"Último save: {ts}")

        st.divider()

        # ── Nova planilha / limpar conversa ───────────────────────────────
        col_a, col_b = st.columns(2)
        with col_a:
            if st.button("🗒️ Nova planilha", use_container_width=True, help="Começa do zero"):
                st.session_state.workbook_bytes = create_template(
                    project_name=st.session_state.project_name
                ).read()
                st.session_state.display_messages = []
                st.session_state.api_messages = []
                st.session_state.last_saved = None
                st.rerun()
        with col_b:
            if st.button("🧹 Limpar chat", use_container_width=True, help="Mantém planilha"):
                st.session_state.display_messages = []
                st.session_state.api_messages = []
                st.rerun()

        st.divider()

        # ── Info da planilha atual ─────────────────────────────────────────
        wb = get_workbook()
        data = get_sheet_data(wb)
        n_items = len(data)
        n_ambientes = len({r["ambiente"] for r in data if r["ambiente"]})
        n_decidir = sum(1 for r in data if r["status"] == "DECIDIR")

        st.markdown(f"**{n_items} itens · {n_ambientes} ambientes**")
        if n_decidir > 0:
            st.warning(f"⚠️ {n_decidir} item(s) para decidir")

        # API key status
        api_key = get_api_key()
        if not api_key:
            st.error("API key não configurada.\nCrie `.streamlit/secrets.toml`.")

        st.divider()
        st.caption(f"Modelo: `{MODEL}`")
        st.caption("v1.0 · Arqui Specs")


# ─── Processamento de mensagem ───────────────────────────────────────────────

def process_message(user_input: str, api_key: str, preview_placeholder):
    """
    Processa uma mensagem do usuário:
    1. Mostra mensagem do usuário
    2. Chama Claude com streaming + tool use
    3. Se houve tool use, salva planilha atualizada
    4. Atualiza preview ao vivo
    5. Salva mensagem no histórico
    """
    wb = get_workbook()

    # Adicionar mensagem ao histórico de display
    st.session_state.display_messages.append({"role": "user", "content": user_input})

    # Construir mensagem para a API (com contexto da planilha)
    api_user_msg = build_user_message(user_input, wb)
    st.session_state.api_messages.append(api_user_msg)

    # Streaming da resposta
    with st.chat_message("assistant", avatar="📋"):
        placeholder = st.empty()
        full_response = ""
        workbook_updated = False
        error_msg = None

        for chunk, updated_wb, err in stream_response(
            api_key=api_key,
            api_messages=st.session_state.api_messages,
            workbook=wb,
        ):
            if err:
                error_msg = err
                break
            if updated_wb:
                # Workbook foi modificado pelo tool use
                save_workbook(updated_wb)
                workbook_updated = True
            if chunk:
                # Limpa o marcador "*Atualizando planilha...*" da exibição final
                full_response += chunk
                # Remover o marcador de loading do texto final exibido
                display_text = full_response.replace("*Atualizando planilha...*\n\n", "")
                placeholder.markdown(display_text + "▌")

        if error_msg:
            placeholder.error(error_msg)
            full_response = error_msg
        else:
            display_text = full_response.replace("*Atualizando planilha...*\n\n", "")
            placeholder.markdown(display_text)
            full_response = display_text

    # Salvar resposta no histórico
    st.session_state.display_messages.append({"role": "assistant", "content": full_response})
    st.session_state.api_messages.append({"role": "assistant", "content": full_response})

    # Atualizar preview se planilha mudou (SCRUM-54)
    if workbook_updated:
        render_preview(preview_placeholder, get_workbook())
        # Auto-increment draft counter (SCRUM-62)
        st.session_state.draft_count += 1


# ─── Preview da planilha (SCRUM-54) ──────────────────────────────────────────

def render_preview(container, wb: openpyxl.Workbook):
    """Renderiza o preview da planilha no container fornecido."""
    with container:
        df = get_preview_dataframe(wb)
        info = get_project_info(wb)

        if df.empty:
            st.info("Comece descrevendo os itens do projeto no chat ao lado.")
            st.markdown("**Exemplo:** *cozinha: cuba Tramontina inox 50x40, torneira Docol gourmet preta*")
            return

        # Cabeçalho do projeto
        st.markdown(f"**{info['nome']}**")
        if info["cliente"] and info["cliente"] != "(definir)":
            st.caption(f"Cliente: {info['cliente']} · {info['emissao']} · {info['versao']}")

        # Tabela com estilo
        try:
            styled = style_preview_dataframe(df)
            st.dataframe(
                styled,
                use_container_width=True,
                height=min(600, 60 + len(df) * 38),
                hide_index=True,
            )
        except Exception:
            # Fallback sem estilo se a versão do pandas/streamlit não suportar
            st.dataframe(df, use_container_width=True, hide_index=True)

        # Rodapé de contagem
        n = len(df)
        n_ap = (df["STATUS"] == "APROVADO").sum()
        n_dec = (df["STATUS"] == "DECIDIR").sum()
        st.caption(f"{n} itens · ✓ {n_ap} aprovados · ? {n_dec} a decidir")


# ─── App principal ───────────────────────────────────────────────────────────

def main():
    # 1. Auth
    if not check_auth():
        return

    # 2. Init
    init_session()
    api_key = get_api_key()

    # 3. Sidebar
    render_sidebar()

    # 4. Layout: 2 colunas
    col_chat, col_preview = st.columns([1, 1], gap="large")

    # 5. Coluna de chat (esquerda)
    with col_chat:
        st.markdown("### 💬 Conversa")

        # Container com altura fixa e scrollbar — evita que a página cresça indefinidamente
        chat_container = st.container(height=580)

        with chat_container:
            # Mensagem de boas-vindas se chat vazio
            if not st.session_state.display_messages:
                with st.chat_message("assistant", avatar="📋"):
                    st.markdown(
                        "Olá! Pronta para montar o memorial.\n\n"
                        "Descreva os itens em linguagem natural — ambiente por ambiente:\n\n"
                        "> *cozinha: cuba Tramontina inox 50x40, torneira Docol gourmet preta, "
                        "cooktop Brastemp indução 5 bocas*"
                    )

            # Histórico de mensagens
            for msg in st.session_state.display_messages:
                avatar = "👤" if msg["role"] == "user" else "📋"
                with st.chat_message(msg["role"], avatar=avatar):
                    st.markdown(msg["content"])

        # Input fora do container — fica fixo abaixo da área de scroll
        user_input = st.chat_input(
            placeholder="Descreva os itens do projeto...",
            disabled=(api_key is None),
        )

    # 6. Coluna de preview (direita) — sempre renderiza com estado atual
    with col_preview:
        st.markdown("### 📊 Planilha ao vivo")
        preview_placeholder = st.container()
        wb_current = get_workbook()
        render_preview(preview_placeholder, wb_current)

    # 7. Processar input (após ambas as colunas serem definidas)
    if user_input and api_key:
        with col_chat:
            with st.chat_message("user", avatar="👤"):
                st.markdown(user_input)
        process_message(user_input, api_key, preview_placeholder)
        # Re-renderiza o app inteiro para que a sidebar (download button, contadores)
        # reflita o workbook atualizado. Sem isso, o botão de download captura
        # os bytes da renderização anterior (planilha vazia).
        st.rerun()

    elif user_input and not api_key:
        st.error(
            "Configure a API key em `.streamlit/secrets.toml` antes de usar.\n"
            "Copie `.streamlit/secrets.toml.example` e preencha com sua chave."
        )


if __name__ == "__main__":
    main()
