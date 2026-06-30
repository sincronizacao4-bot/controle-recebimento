"""
Gerador de Controle de Recebimento de Materiais
App web com Streamlit — funciona em qualquer navegador, sem instalar nada.
"""

import streamlit as st
import pandas as pd
import base64, os
from datetime import date
from extrator import extrair
from gerador_docx import gerar

_BRASAO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "brasao.png")

def _brasao_b64() -> str:
    try:
        with open(_BRASAO_PATH, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except Exception:
        return ""


# ── Efeito constelação (fundo animado) ────────────────────────────────────────
def _constellation():
    st.markdown("""
    <canvas id="st-cv" style="position:fixed;top:0;left:0;width:100vw;height:100vh;
        pointer-events:none;z-index:9999;"></canvas>
    <script>
    (function launch(){
      const cv = document.getElementById('st-cv');
      if(!cv){ setTimeout(launch, 100); return; }
      const ctx = cv.getContext('2d');
      let W, H, stars=[], mouse={x:-9999,y:-9999};
      const N=130, LINK=140, MR=180;

      function resize(){
        W=cv.width=window.innerWidth;
        H=cv.height=window.innerHeight;
      }
      function init(){
        stars=[];
        for(let i=0;i<N;i++) stars.push({
          x:Math.random()*W, y:Math.random()*H,
          vx:(Math.random()-.5)*.35, vy:(Math.random()-.5)*.35,
          r:Math.random()*1.4+.5
        });
      }
      function draw(){
        ctx.clearRect(0,0,W,H);
        stars.forEach(s=>{
          s.x+=s.vx; s.y+=s.vy;
          if(s.x<0||s.x>W)s.vx*=-1;
          if(s.y<0||s.y>H)s.vy*=-1;
        });
        for(let i=0;i<stars.length;i++){
          for(let j=i+1;j<stars.length;j++){
            let d=Math.hypot(stars[j].x-stars[i].x, stars[j].y-stars[i].y);
            if(d<LINK){
              ctx.beginPath();
              ctx.strokeStyle=`rgba(180,210,255,${.25*(1-d/LINK)})`;
              ctx.lineWidth=.5;
              ctx.moveTo(stars[i].x,stars[i].y);
              ctx.lineTo(stars[j].x,stars[j].y);
              ctx.stroke();
            }
          }
          let d=Math.hypot(mouse.x-stars[i].x, mouse.y-stars[i].y);
          if(d<MR){
            ctx.beginPath();
            ctx.strokeStyle=`rgba(150,200,255,${.85*(1-d/MR)})`;
            ctx.lineWidth=.9;
            ctx.moveTo(stars[i].x,stars[i].y);
            ctx.lineTo(mouse.x,mouse.y);
            ctx.stroke();
          }
        }
        stars.forEach(s=>{
          ctx.beginPath();
          ctx.arc(s.x,s.y,s.r,0,Math.PI*2);
          ctx.fillStyle='rgba(255,255,255,.9)';
          ctx.fill();
        });
        requestAnimationFrame(draw);
      }
      document.addEventListener('mousemove',e=>{mouse.x=e.clientX;mouse.y=e.clientY;});
      window.addEventListener('resize',()=>{resize();init();});
      resize(); init(); draw();
    })();
    </script>
    """, unsafe_allow_html=True)

# ── Configuração da página ─────────────────────────────────────────────────────
st.set_page_config(
    page_title="SEINFRA – Controle de Recebimento",
    page_icon="🏛️",
    layout="wide",
)

# ── Senha de acesso ────────────────────────────────────────────────────────────
USUARIO_CORRETO = "Crislene"
SENHA_CORRETA   = "Pmm@Seinfra#2025"

def _login_page():
    b64 = _brasao_b64()
    logo_src = f"data:image/png;base64,{b64}" if b64 else ""

    st.markdown(f"""
    <style>
      /* ── Fundo ── */
      [data-testid="stAppViewContainer"] {{
        background: linear-gradient(135deg,#1a1a2e 0%,#16213e 40%,#0f3460 70%,#533483 100%) !important;
        min-height: 100vh;
      }}
      [data-testid="stHeader"], [data-testid="stToolbar"] {{ display:none !important; }}
      #MainMenu, footer {{ visibility:hidden; }}

      /* ── Glass card aplicado ao container do Streamlit ── */
      .block-container {{
        max-width: 460px !important;
        margin: 0 auto !important;
        padding: 44px 40px 36px !important;
        background: rgba(255,255,255,0.07) !important;
        backdrop-filter: blur(24px) !important;
        -webkit-backdrop-filter: blur(24px) !important;
        border: 1px solid rgba(255,255,255,0.15) !important;
        border-radius: 28px !important;
        box-shadow: 0 25px 50px rgba(0,0,0,0.5) !important;
        margin-top: 60px !important;
      }}

      /* ── Avatar dourado ── */
      .gold-wrap {{
        position:relative; width:116px; height:116px; margin:0 auto 20px;
      }}
      .gold-spin {{
        position:absolute; inset:0; border-radius:22px;
        background: conic-gradient(
          #FFD700 0%,#FFF5A0 12%,#FFA500 25%,
          transparent 40%,transparent 60%,
          #FFA500 75%,#FFF5A0 88%,#FFD700 100%
        );
        animation: gspin 2.4s linear infinite;
      }}
      @keyframes gspin {{ to {{ transform:rotate(360deg); }} }}
      .gold-dot {{
        position:absolute; width:8px; height:8px; border-radius:50%;
        background:radial-gradient(circle,#FFD700,#FF8C00);
        box-shadow:0 0 10px 3px #FFD70099;
        top:-4px; left:calc(50% - 4px);
        transform-origin:4px calc(58px + 4px);
        animation:gorbit 2.4s linear infinite;
      }}
      .gold-dot.d2 {{ animation-delay:-1.2s; }}
      @keyframes gorbit {{ to {{ transform:rotate(360deg); }} }}
      .gold-inner {{
        position:absolute; inset:4px; border-radius:18px;
        background:rgba(10,20,55,0.92);
        display:flex; align-items:center; justify-content:center; overflow:hidden;
      }}
      .gold-inner img {{ width:80px; height:80px; object-fit:contain; }}

      .lt {{ text-align:center;font-size:19px;font-weight:700;
              color:#fff;margin:0 0 4px;letter-spacing:.3px; }}
      .ls {{ text-align:center;font-size:11px;color:rgba(255,255,255,.5);
              margin:0 0 24px;letter-spacing:.8px;text-transform:uppercase; }}

      /* ── Inputs ── */
      [data-testid="stTextInput"] label {{
        color:rgba(255,255,255,.75) !important; font-size:13px !important; }}
      [data-testid="stTextInput"] input {{
        background:rgba(255,255,255,.08) !important;
        border:1px solid rgba(255,255,255,.2) !important;
        border-radius:12px !important; color:#fff !important; font-size:14px !important;
      }}
      [data-testid="stTextInput"] input:focus {{
        border-color:rgba(255,255,255,.5) !important;
        box-shadow:0 0 0 3px rgba(255,255,255,.06) !important;
      }}
      [data-testid="stTextInput"] input::placeholder {{ color:rgba(255,255,255,.35) !important; }}

      /* ── Botão ── */
      [data-testid="stButton"] > button {{
        background:linear-gradient(135deg,#667eea,#764ba2) !important;
        color:#fff !important; border:none !important;
        border-radius:12px !important; font-size:15px !important;
        font-weight:700 !important; letter-spacing:1.5px !important;
        text-transform:uppercase !important; width:100% !important;
        margin-top:6px !important;
        box-shadow:0 8px 20px rgba(102,126,234,.45) !important;
      }}
      [data-testid="stButton"] > button:hover {{
        transform:translateY(-2px) !important;
        box-shadow:0 12px 28px rgba(102,126,234,.6) !important;
      }}

      /* ── Erro ── */
      [data-testid="stAlert"] {{
        background:rgba(231,76,60,.18) !important;
        border:1px solid rgba(231,76,60,.4) !important;
        border-radius:10px !important; color:#ff8a80 !important;
      }}
      .lf {{ text-align:center;font-size:11px;color:rgba(255,255,255,.25);
              margin-top:18px;letter-spacing:.5px; }}
    </style>

    <div class="gold-wrap">
      <div class="gold-spin"></div>
      <div class="gold-dot"></div>
      <div class="gold-dot d2"></div>
      <div class="gold-inner">
        {"<img src='" + logo_src + "'/>" if logo_src else "🏛️"}
      </div>
    </div>
    <p class="lt">Prefeitura de Maracanaú</p>
    <p class="ls">SEINFRA · Controle de Recebimento</p>
    """, unsafe_allow_html=True)

    _constellation()

    usuario = st.text_input("Usuário", placeholder="Nome de usuário")
    senha   = st.text_input("Senha", type="password", placeholder="••••••••••••••••")
    entrar  = st.button("LOGIN", use_container_width=True)

    if entrar:
        if usuario.strip().lower() == USUARIO_CORRETO.lower() and senha == SENHA_CORRETA:
            st.session_state["autenticado"] = True
            st.rerun()
        else:
            st.error("Usuário ou senha incorretos. Tente novamente.")

    st.markdown('<p class="lf">Acesso restrito · SEINFRA</p>', unsafe_allow_html=True)

# ── Controle de sessão ─────────────────────────────────────────────────────────
if not st.session_state.get("autenticado"):
    _login_page()
    st.stop()

# ── Barra superior com logout ──────────────────────────────────────────────────
_constellation()
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

with st.expander("🔍 DEBUG – texto bruto do OCR (para suporte técnico)"):
    st.text(dados.get("_ocr_text", "(não disponível)"))

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

_items_raw = dados.get("items", [])
_n_suspeitos = sum(1 for i in _items_raw if i.get("_suspeito"))
if _n_suspeitos:
    st.warning(
        f"⚠️ **{_n_suspeitos} item(ns)** ficaram com algum valor (quantidade, "
        f"valor unitário ou total) **não reconhecido pelo OCR**. Revise as linhas "
        f"destacadas na tabela abaixo antes de gerar o documento."
    )

df_items = pd.DataFrame(
    _items_raw,
    columns=_cols + ["_suspeito"],
) if _items_raw else pd.DataFrame(columns=_cols + ["_suspeito"])

edited = st.data_editor(
    df_items,
    column_config={
        **{k: st.column_config.TextColumn(v) for k, v in _labels.items()},
        "_suspeito": st.column_config.CheckboxColumn("⚠️ Revisar", help="Marcado automaticamente quando o OCR não conseguiu ler todos os valores deste item."),
    },
    num_rows="dynamic",
    use_container_width=True,
    height=min(420, 60 + 35 * max(1, len(df_items))),
)

dados["items"] = [
    {k: v for k, v in row.items() if k != "_suspeito"}
    for row in edited.to_dict("records")
]

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

# Validação cruzada por linha: quantidade × valor unitário deve bater com o total
_linhas_incoerentes = []
for idx, item in enumerate(dados["items"]):
    q = _parse_valor(item.get("quantidade", 0))
    vu = _parse_valor(item.get("valor_unitario", 0))
    vt = _parse_valor(item.get("valor_total", 0))
    if q and vu and vt:
        esperado = q * vu
        if abs(esperado - vt) > max(0.05, esperado * 0.01):
            _linhas_incoerentes.append((idx + 1, item.get("descricao", ""), esperado, vt))

if _linhas_incoerentes:
    msg = "⚠️ **Inconsistência matemática encontrada:**\n"
    for n, desc, esp, vt in _linhas_incoerentes:
        msg += f"\n- Item {n} ({desc[:40]}): qtd × vlr unit. = **R$ {esp:,.2f}**, mas total informado é **R$ {vt:,.2f}**"
    st.error(msg)

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
        label="⬇️  Baixar Controle de Recebimento (.docx)",
        data=docx_bytes,
        file_name=nome_arquivo,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )

    st.success("Documento gerado! Baixe o .docx para editar no Word, ou use o preview abaixo para imprimir direto pelo navegador.")

    # ── Preview imprimível direto do navegador (sem precisar abrir o Word) ───
    with st.expander("🖨️  Visualizar e imprimir (sem precisar do Word)", expanded=False):
        _itens_html = "".join(
            f"<tr><td>{i+1}</td><td>{it.get('codigo','')}</td><td style='text-align:left'>{it.get('descricao','')}</td>"
            f"<td>{it.get('marca','')}</td><td>{it.get('unidade','')}</td><td>{it.get('quantidade','')}</td>"
            f"<td>{it.get('valor_unitario','')}</td><td>{it.get('valor_total','')}</td></tr>"
            for i, it in enumerate(dados.get("items", []))
        )
        b64_logo = _brasao_b64()
        _logo_tag = f"<img src='data:image/png;base64,{b64_logo}' style='height:60px'/>" if b64_logo else ""

        preview_html = f"""
        <div id="doc-preview" style="background:#fff;color:#111;padding:24px;border-radius:8px;
            font-family:Arial,sans-serif;font-size:12px;max-width:900px;margin:0 auto;">
          <div style="display:flex;align-items:center;gap:12px;border-bottom:2px solid #1A3F7A;padding-bottom:8px;margin-bottom:12px;">
            {_logo_tag}
            <div>
              <div style="font-weight:bold;font-size:14px;color:#1A3F7A;">Prefeitura Municipal de Maracanaú</div>
              <div style="font-size:10px;color:#333;">Secretaria de Infraestrutura, Mobilidade e Controle Urbano</div>
            </div>
          </div>
          <h2 style="text-align:center;color:#C05000;text-decoration:underline;font-size:15px;">CONTROLE DE RECEBIMENTO DE MATERIAIS</h2>
          <p><b>CONTRATADA:</b> {dados.get('fornecedor','')} &nbsp; <b>CNPJ:</b> {dados.get('cnpj','')}</p>
          <p><b>EMPENHO:</b> {dados.get('empenho','')} &nbsp; <b>DANFE Nº:</b> {dados.get('danfe','')} &nbsp; <b>DATA:</b> {dados.get('data_recebimento','')}</p>
          <p><b>OBJETO:</b> {dados.get('objeto','')}</p>
          <table style="width:100%;border-collapse:collapse;margin-top:8px;" border="1">
            <thead style="background:#1A3F7A;color:#fff;">
              <tr><th>Nº</th><th>Código</th><th>Descrição</th><th>Marca</th><th>Unid.</th><th>Qtd.</th><th>Vlr Unit.</th><th>Vlr Total</th></tr>
            </thead>
            <tbody>{_itens_html}</tbody>
          </table>
          <p style="margin-top:24px;"><b>RESPONSÁVEL:</b> {dados.get('responsavel','')} — Matrícula {dados.get('matricula','')}</p>
          <p>{dados.get('cargo','')}</p>
        </div>
        <style>
          #doc-preview table td, #doc-preview table th {{ border:1px solid #999; padding:4px; text-align:center; }}
          @media print {{
            [data-testid="stSidebar"], [data-testid="stHeader"], [data-testid="stToolbar"],
            .stButton, #MainMenu, footer {{ display:none !important; }}
          }}
        </style>
        <button onclick="window.print()" style="margin-top:12px;padding:10px 22px;background:#1A3F7A;
            color:#fff;border:none;border-radius:8px;font-weight:bold;cursor:pointer;">
            🖨️ Imprimir agora
        </button>
        """
        st.markdown(preview_html, unsafe_allow_html=True)
        st.caption(
            "Este preview é só para impressão rápida (layout simplificado). "
            "Para o documento oficial completo (com marca d'água, rodapé e formatação final), "
            "use o arquivo .docx baixado acima."
        )
