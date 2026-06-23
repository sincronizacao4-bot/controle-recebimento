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
        /* Remove padding padrão do Streamlit */
        [data-testid="stAppViewContainer"] > .main { padding: 0 !important; }
        [data-testid="stAppViewContainer"] {
            background: linear-gradient(135deg, #1a1a2e 0%, #16213e 40%, #0f3460 70%, #533483 100%);
            min-height: 100vh;
        }
        [data-testid="stHeader"]  { background: transparent !important; }
        [data-testid="stToolbar"] { visibility: hidden; }
        #MainMenu, footer         { visibility: hidden; }

        /* Bolhas decorativas de fundo */
        body::before {
            content: "";
            position: fixed;
            width: 400px; height: 400px;
            background: radial-gradient(circle, rgba(83,52,131,0.5), transparent 70%);
            top: -100px; left: -100px;
            border-radius: 50%;
            pointer-events: none;
        }
        body::after {
            content: "";
            position: fixed;
            width: 350px; height: 350px;
            background: radial-gradient(circle, rgba(15,52,96,0.6), transparent 70%);
            bottom: -80px; right: -80px;
            border-radius: 50%;
            pointer-events: none;
        }

        /* Card glassmorphism */
        .glass-card {
            background: rgba(255,255,255,0.08);
            backdrop-filter: blur(20px);
            -webkit-backdrop-filter: blur(20px);
            border: 1px solid rgba(255,255,255,0.18);
            border-radius: 24px;
            padding: 48px 40px 40px 40px;
            max-width: 400px;
            margin: 60px auto 0 auto;
            box-shadow: 0 25px 45px rgba(0,0,0,0.4),
                        0 0 0 1px rgba(255,255,255,0.05) inset;
        }

        /* Avatar */
        .avatar-ring {
            width: 90px; height: 90px;
            border-radius: 50%;
            background: rgba(255,255,255,0.12);
            border: 2px solid rgba(255,255,255,0.3);
            display: flex; align-items: center; justify-content: center;
            margin: 0 auto 20px auto;
            font-size: 42px;
            box-shadow: 0 8px 32px rgba(0,0,0,0.3);
        }

        /* Textos */
        .login-title {
            text-align: center;
            font-size: 18px;
            font-weight: 700;
            color: #ffffff;
            margin: 0 0 4px 0;
            letter-spacing: 0.3px;
        }
        .login-sub {
            text-align: center;
            font-size: 12px;
            color: rgba(255,255,255,0.55);
            margin: 0 0 32px 0;
            letter-spacing: 0.5px;
            text-transform: uppercase;
        }

        /* Inputs com estilo escuro */
        [data-testid="stTextInput"] input {
            background: rgba(255,255,255,0.08) !important;
            border: 1px solid rgba(255,255,255,0.2) !important;
            border-radius: 12px !important;
            color: #ffffff !important;
            padding: 14px 16px !important;
            font-size: 14px !important;
            transition: border 0.2s;
        }
        [data-testid="stTextInput"] input:focus {
            border: 1px solid rgba(255,255,255,0.5) !important;
            box-shadow: 0 0 0 3px rgba(255,255,255,0.06) !important;
        }
        [data-testid="stTextInput"] input::placeholder { color: rgba(255,255,255,0.4) !important; }
        [data-testid="stTextInput"] label { color: rgba(255,255,255,0.7) !important; font-size:13px !important; }

        /* Botão login */
        [data-testid="stButton"] > button {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
            color: white !important;
            border: none !important;
            border-radius: 12px !important;
            padding: 14px !important;
            font-size: 15px !important;
            font-weight: 700 !important;
            letter-spacing: 1.5px !important;
            text-transform: uppercase !important;
            width: 100% !important;
            margin-top: 8px !important;
            box-shadow: 0 8px 20px rgba(102,126,234,0.4) !important;
            transition: all 0.2s !important;
        }
        [data-testid="stButton"] > button:hover {
            transform: translateY(-2px) !important;
            box-shadow: 0 12px 28px rgba(102,126,234,0.55) !important;
        }

        /* Mensagem de erro */
        [data-testid="stAlert"] {
            background: rgba(231,76,60,0.2) !important;
            border: 1px solid rgba(231,76,60,0.4) !important;
            border-radius: 10px !important;
            color: #ff8a80 !important;
        }

        .rodape-glass {
            text-align: center;
            font-size: 11px;
            color: rgba(255,255,255,0.3);
            margin-top: 24px;
            letter-spacing: 0.5px;
        }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="glass-card">
        <div class="avatar-ring">🏛️</div>
        <p class="login-title">Prefeitura de Maracanaú</p>
        <p class="login-sub">SEINFRA · Controle de Recebimento</p>
    </div>
    """, unsafe_allow_html=True)

    # Inputs dentro do card via CSS overlap
    with st.container():
        st.markdown('<div style="max-width:400px;margin:0 auto;margin-top:-220px;padding:0 40px;">', unsafe_allow_html=True)

        senha = st.text_input(
            "Senha",
            type="password",
            placeholder="••••••••••••••••",
        )
        entrar = st.button("LOGIN", use_container_width=True)

        if entrar:
            if senha == SENHA_CORRETA:
                st.session_state["autenticado"] = True
                st.rerun()
            else:
                st.error("Senha incorreta. Tente novamente.")

        st.markdown('<p class="rodape-glass">Acesso restrito · SEINFRA</p>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

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

# Objeto extraído automaticamente (editável)
dados["objeto"] = st.text_area(
    "📋 Objeto da Ordem (extraído automaticamente — edite se necessário)",
    value=dados.get("objeto", ""),
    height=80,
)

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
    dados["matricula"] = st.text_input("Matrícula", "52680")
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

# ── Validação de valores ───────────────────────────────────────────────────────
def _parse_valor(v: str) -> float:
    """Converte '11.548,15' ou '11548.15' para float."""
    try:
        v = str(v).strip().replace(" ", "")
        if "," in v and "." in v:
            v = v.replace(".", "").replace(",", ".")
        elif "," in v:
            v = v.replace(",", ".")
        return float(v)
    except Exception:
        return 0.0

_total_itens = sum(_parse_valor(i.get("valor_total", 0)) for i in dados["items"])
_valor_of    = _parse_valor(dados.get("valor_of", "0"))

if dados["items"] and _valor_of > 0:
    _diff = abs(_total_itens - _valor_of)
    if _diff > 0.10:
        st.warning(
            f"⚠️ **Atenção:** A soma dos itens (**R$ {_total_itens:,.2f}**) "
            f"não bate com o valor da Ordem (**R$ {_valor_of:,.2f}**). "
            f"Diferença de R$ {_diff:,.2f}. Verifique os valores antes de gerar."
        )
    else:
        st.info(f"✅ Soma dos itens: **R$ {_total_itens:,.2f}** — confere com o valor da Ordem.")

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
