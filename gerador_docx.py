"""
Geração do Controle de Recebimento em DOCX.
"""

import io
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ── Helpers ────────────────────────────────────────────────────────────────────
def _set_bg(cell, hex_color: str):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def _cell(cell, text: str, bold=False, size=9, center=False, white=False):
    para = cell.paragraphs[0]
    para.clear()
    run = para.add_run(str(text))
    run.bold = bold
    run.font.size = Pt(size)
    if white:
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    if center:
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.paragraph_format.space_before = Pt(1)
    para.paragraph_format.space_after = Pt(1)


def _header_row(row, cols: list[str], bg="1F4E79"):
    for i, label in enumerate(cols):
        _set_bg(row.cells[i], bg)
        _cell(row.cells[i], label, bold=True, center=True, white=True)


def _info_table(doc, rows: list[tuple], header: str | None = None,
                col_widths=(Cm(5), Cm(11.5))):
    extra = 1 if header else 0
    tbl = doc.add_table(rows=len(rows) + extra, cols=2)
    tbl.style = "Table Grid"
    for row in tbl.rows:
        for i, w in enumerate(col_widths):
            row.cells[i].width = w

    if header:
        merged = tbl.rows[0].cells[0]
        merged.merge(tbl.rows[0].cells[1])
        _set_bg(merged, "1F4E79")
        _cell(merged, header, bold=True, white=True)

    for i, (label, value) in enumerate(rows):
        r = tbl.rows[i + extra]
        _set_bg(r.cells[0], "D9D9D9")
        _cell(r.cells[0], label, bold=True)
        _cell(r.cells[1], value)

    return tbl


def _spacer(doc, lines=1):
    for _ in range(lines):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)


# ── Geração principal ──────────────────────────────────────────────────────────
def gerar(dados: dict) -> bytes:
    """
    Recebe o dict de dados (da extração + campos do usuário) e
    retorna os bytes do arquivo DOCX.
    """
    doc = Document()

    # Margens estreitas para caber mais itens
    for section in doc.sections:
        section.top_margin    = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin   = Cm(2)
        section.right_margin  = Cm(2)

    # Fonte padrão
    doc.styles["Normal"].font.name = "Arial"
    doc.styles["Normal"].font.size = Pt(9)

    # ── Título ──────────────────────────────────────────────────────────────
    titulo = doc.add_paragraph("CONTROLE DE RECEBIMENTO DE MATERIAIS")
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    titulo.paragraph_format.space_after = Pt(8)
    run = titulo.runs[0]
    run.bold = True
    run.font.size = Pt(13)

    # ── 1. Partes ────────────────────────────────────────────────────────────
    _info_table(doc, [
        ("CONTRATANTE", "Prefeitura Municipal de Maracanaú"),
        ("CONTRATADA",  dados.get("fornecedor", "")),
        ("CNPJ",        dados.get("cnpj", "")),
    ])
    _spacer(doc)

    # ── 2. Fundamentação legal ───────────────────────────────────────────────
    _info_table(doc, [
        ("PREGÃO DIGITAL",       dados.get("pregao",   "10.002/2024")),
        ("ATA DE REG. DE PREÇOS",dados.get("ata",      "10.004/2024")),
        ("CONTRATO",              dados.get("contrato", "")),
        ("EMPENHO",               dados.get("empenho",  "")),
    ], header="FUNDAMENTAÇÃO LEGAL")
    _spacer(doc)

    # ── 3. Objeto ────────────────────────────────────────────────────────────
    _info_table(doc, [
        ("", dados.get("objeto",
            f"Aquisição de material conforme Ordem de Fornecimento "
            f"Nº {dados.get('numero_of', '')}.")),
    ], header="OBJETO", col_widths=(Cm(0.1), Cm(16.4)))
    _spacer(doc)

    # ── 4. Dados do recebimento ───────────────────────────────────────────────
    _info_table(doc, [
        ("ORDEM DE FORNECIMENTO Nº", dados.get("numero_of", "")),
        ("DANFE Nº",                 dados.get("danfe", "")),
        ("DATA DE RECEBIMENTO",      dados.get("data_recebimento", "")),
        ("LOCAL DE ENTREGA",         dados.get("local", "")),
    ], header="DADOS DO RECEBIMENTO")
    _spacer(doc)

    # ── 5. Tabela de itens ────────────────────────────────────────────────────
    items = dados.get("items", [])
    if items:
        sub = doc.add_paragraph("ITENS RECEBIDOS")
        sub.paragraph_format.space_before = Pt(2)
        sub.paragraph_format.space_after  = Pt(2)
        sub.runs[0].bold = True
        sub.runs[0].font.size = Pt(10)

        HDR = ["Nº", "CÓDIGO", "DESCRIÇÃO", "MARCA", "UNIDADE", "QTDE",
               "VLR UNIT.", "VLR TOTAL"]
        WIDTHS = [Cm(0.8), Cm(1.5), Cm(6.0), Cm(2.0), Cm(1.8),
                  Cm(1.4), Cm(1.8), Cm(1.8)]

        tbl = doc.add_table(rows=1 + len(items), cols=len(HDR))
        tbl.style = "Table Grid"

        # Larguras
        for row in tbl.rows:
            for i, w in enumerate(WIDTHS):
                row.cells[i].width = w

        # Cabeçalho
        _header_row(tbl.rows[0], HDR)

        # Dados
        for idx, item in enumerate(items):
            row = tbl.rows[idx + 1]
            bg = "EBF3FB" if idx % 2 == 0 else "FFFFFF"
            for cell in row.cells:
                _set_bg(cell, bg)

            vals = [
                str(idx + 1),
                item.get("codigo", ""),
                item.get("descricao", ""),
                item.get("marca", ""),
                item.get("unidade", ""),
                item.get("quantidade", ""),
                item.get("valor_unitario", ""),
                item.get("valor_total", ""),
            ]
            for i, v in enumerate(vals):
                center = i not in (2,)   # só descrição alinhada à esquerda
                _cell(row.cells[i], v, center=center)

    _spacer(doc)

    # ── 6. Declaração ─────────────────────────────────────────────────────────
    decl = doc.add_paragraph(
        "Declaro, para os devidos fins, que os materiais e/ou serviços objeto deste "
        "documento foram recebidos nesta data, em conformidade com as especificações "
        "técnicas e quantitativas constantes no contrato/ordem de fornecimento, "
        "encontrando-se em perfeitas condições de uso e funcionamento."
    )
    decl.paragraph_format.space_before = Pt(4)
    decl.paragraph_format.space_after  = Pt(16)

    # ── 7. Assinatura ─────────────────────────────────────────────────────────
    resp = dados.get("responsavel", "Clayton Rocha Braz")
    cargo = dados.get("cargo", "DIRETOR PÁTIO DE MANUTENÇÃO – SEINFRA")
    for txt in (resp, cargo):
        p = doc.add_paragraph(txt)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].bold = True
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(0)

    # ── Retorna bytes ─────────────────────────────────────────────────────────
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()
