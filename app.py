"""
Gerador de Controle de Recebimento de Materiais
App web com Streamlit — funciona em qualquer navegador, sem instalar nada.
"""

import streamlit as st
import pandas as pd
from datetime import date

from extrator import extrair
from gerador_docx import gerar

# ── Configuração da página ─────────────────────────────────────────────────────
st.set_page_config(
    page_title="SEINFRA – Controle de Recebimento",
    page_icon="🏛️",
    layout="wide",
)

# ── Senha de acesso ────────────────────────────────────────────────────────────
SENHA_CORRETA = "Pmm@Seinfra#2025"

def _login_page():
    st.markdown("""
    <style>
        /* Fundo geral */
        [data-testid="stAppViewContainer"] {
            background: linear-gradient(135deg, #0a3d62 0%, #1a5276 60%, #2471a3 100%);
            min-height: 100vh;
        }
        [data-testid="stHeader"] { background: transparent; }

        /* Card central */
        .login-card {
            background: white;
            border-radius: 16px;
            padding: 48px 40px 40px 40px;
            max-width: 420px;
            margin: 80px auto 0 auto;
            box-shadow: 0 20px 60px rgba(0,0,0,0.35);
        }
        .brasao-container {
            text-align: center;
            margin-bottom: 8px;
        }
        .login-titulo {
            text-align: center;
            font-size: 20px;
            font-weight: 700;
            color: #0a3d62;
            margin: 0 0 4px 0;
            line-height: 1.3;
        }
        .login-subtitulo {
            text-align: center;
            font-size: 13px;
            color: #7f8c8d;
            margin: 0 0 28px 0;
        }
        .login-divider {
            border: none;
            border-top: 1px solid #ecf0f1;
            margin: 0 0 24px 0;
        }
        .rodape-login {
            text-align: center;
            font-size: 11px;
            color: #aab7b8;
            margin-top: 28px;
        }

        /* Esconde elementos Streamlit desnecessários na tela de login */
        #MainMenu, footer, [data-testid="stToolbar"] { visibility: hidden; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown('<div class="login-card">', unsafe_allow_html=True)

    # Brasão / ícone
    st.markdown("""
    <div class="brasao-container">
        <div style="
            width:72px; height:72px; border-radius:50%;
            background:linear-gradient(135deg,#0a3d62,#2471a3);
            display:flex; align-items:center; justify-content:center;
            margin:0 auto 16px auto;
            font-size:36px; line-height:72px; text-align:center;
        ">🏛️</div>
    </div>
    <p class="login-titulo">Prefeitura Municipal de Maracanaú</p>
    <p class="login-subtitulo">Sistema de Controle de Recebimento – SEINFRA</p>
    <hr class="login-divider">
    """, unsafe_allow_html=True)

    senha = st.text_input(
        "🔒  Senha de acesso",
        type="password",
        placeholder="Digite a senha...",
        label_visibility="collapsed",
    )

    entrar = st.button("Entrar", type="primary", use_container_width=True)

    if entrar:
        if senha == SENHA_CORRETA:
            st.session_state["autenticado"] = True
            st.rerun()
        else:
            st.error("Senha incorreta. Tente novamente.")

    st.markdown("""
    <p class="rodape-login">Acesso restrito a servidores autorizados da SEINFRA</p>
    </div>
    """, unsafe_allow_html=True)

# ── Controle de sessão ─────────────────────────────────────────────────────────
if not st.session_state.get("autenticado"):
    _login_page()
    st.stop()

# ── Barra superior com logout ──────────────────────────────────────────────────
col_titulo, col_logout = st.columns([8, 1])
with col_titulo:
    st.title("📋 Gerador de Controle de Recebimento")
    st.caption("Prefeitura Municipal de Maracanaú – SEINFRA")
with col_logout:
    st.write("")
    if st.button("🚪 Sair", help="Encerrar sessão"):
        st.session_state["autenticado"] = False
        st.rerun()

st.divider()

# ── Upload do PDF ──────────────────────────────────────────────────────────────
st.header("1. Ordem de Fornecimento (PDF)")
pdf_file = st.file_uploader(
    "Selecione o arquivo PDF da Ordem de Fornecimento",
    type=["pdf"],
    help="Funciona com PDFs digitais e também com PDFs escaneados (imagem).",
)

if not pdf_file:
    st.info("⬆️ Faça upload do PDF para começar.")
    st.stop()

# ── Extração ───────────────────────────────────────────────────────────────────
with st.spinner("📖 Lendo PDF e extraindo dados..."):
    try:
        dados = extrair(pdf_file.read())
    except Exception as e:
        st.error(f"Erro ao ler o PDF: {e}")
        st.stop()

n_itens = len(dados.get("items", []))
st.success(f"✅ PDF lido com sucesso — **{n_itens} item(s)** encontrado(s).")

# ── Exibe / edita dados extraídos ─────────────────────────────────────────────
st.header("2. Dados extraídos (verifique e corrija se necessário)")

col1, col2 = st.columns(2)
with col1:
    dados["numero_of"]  = st.text_input("Nº da Ordem de Fornecimento", dados.get("numero_of",""))
    dados["fornecedor"] = st.text_input("Fornecedor / Contratada", dados.get("fornecedor",""))
    dados["cnpj"]       = st.text_input("CNPJ", dados.get("cnpj",""))
    dados["empenho"]    = st.text_input("Empenho", dados.get("empenho",""))

with col2:
    dados["contrato"]   = st.text_input("Contrato", dados.get("contrato",""))
    dados["pregao"]     = st.text_input("Pregão Digital", dados.get("pregao","10.002/2024"))
    dados["ata"]        = st.text_input("Ata de Reg. de Preços", dados.get("ata","10.004/2024"))
    dados["valor_of"]   = st.text_input("Valor total da Ordem (R$)", dados.get("valor_of",""))

# ── Dados do recebimento ───────────────────────────────────────────────────────
st.header("3. Dados do recebimento")

col3, col4 = st.columns(2)
with col3:
    dados["danfe"] = st.text_input("DANFE Nº", "")
    dados["data_recebimento"] = st.text_input(
        "Data de Recebimento",
        date.today().strftime("%d/%m/%Y"),
    )
with col4:
    dados["local"] = st.text_input(
        "Local de Entrega",
        dados.get("endereco_entrega", "Rua 01, nº 650 – Novo Maracanaú"),
    )
    dados["responsavel"] = st.text_input(
        "Responsável pelo Recebimento",
        "Clayton Rocha Braz",
    )
    dados["cargo"] = st.text_input(
        "Cargo",
        "DIRETOR PÁTIO DE MANUTENÇÃO – SEINFRA",
    )

# ── Tabela de itens ────────────────────────────────────────────────────────────
st.header("4. Itens da Ordem")

if n_itens == 0:
    st.warning(
        "⚠️ Nenhum item foi detectado automaticamente. "
        "Você pode adicioná-los manualmente na tabela abaixo."
    )

_cols = ["codigo", "descricao", "marca", "unidade", "quantidade",
         "valor_unitario", "valor_total"]
_labels = {
    "codigo": "Código",
    "descricao": "Descrição",
    "marca": "Marca",
    "unidade": "Unidade",
    "quantidade": "Quantidade",
    "valor_unitario": "Valor Unitário",
    "valor_total": "Valor Total",
}

df_items = pd.DataFrame(
    dados.get("items", []),
    columns=_cols,
) if dados.get("items") else pd.DataFrame(columns=_cols)

edited = st.data_editor(
    df_items,
    column_config={k: st.column_config.TextColumn(v) for k, v in _labels.items()},
    num_rows="dynamic",
    use_container_width=True,
    height=min(400, 60 + 35 * max(1, len(df_items))),
)

dados["items"] = edited.to_dict("records")

# ── Gerar DOCX ────────────────────────────────────────────────────────────────
st.header("5. Gerar documento")

if not dados.get("danfe"):
    st.warning("⚠️ Preencha o **DANFE Nº** antes de gerar.")

gerar_btn = st.button(
    "✅  Gerar Controle de Recebimento (DOCX)",
    type="primary",
    disabled=not bool(dados.get("danfe")),
)

if gerar_btn:
    with st.spinner("Gerando documento..."):
        try:
            docx_bytes = gerar(dados)
        except Exception as e:
            st.error(f"Erro ao gerar o documento: {e}")
            st.stop()

    nome_arquivo = f"Controle Recebimento OF {dados.get('numero_of','')}.docx"
    st.download_button(
        label="⬇️  Baixar Controle de Recebimento",
        data=docx_bytes,
        file_name=nome_arquivo,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    st.success("Documento gerado! Clique no botão acima para baixar.")
