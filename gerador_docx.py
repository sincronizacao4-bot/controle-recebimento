"""
Geração do Controle de Recebimento em DOCX.
"""

import io
import os
from docx import Document
from docx.shared import Pt, RGBColor, Cm, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement
from lxml import etree

# Caminho do brasão (mesma pasta do script no servidor)
_BRASAO = os.path.join(os.path.dirname(__file__), "brasao.png")


# ── Helpers de célula/tabela ───────────────────────────────────────────────────
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


# ── Cabeçalho com brasão ───────────────────────────────────────────────────────
def _add_header(doc):
    """Cabeçalho com brasão à esquerda e texto da prefeitura à direita."""
    section = doc.sections[0]
    header = section.header
    header.is_linked_to_previous = False

    # Limpa parágrafo padrão do cabeçalho
    for p in header.paragraphs:
        p.clear()

    p = header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(4)

    # Brasão (imagem)
    if os.path.exists(_BRASAO):
        run_img = p.add_run()
        run_img.add_picture(_BRASAO, height=Cm(1.8))
        p.add_run("  ")  # espaço entre brasão e texto

    # Nome da prefeitura
    run_txt = p.add_run("Prefeitura Municipal de Maracanaú\n")
    run_txt.bold = True
    run_txt.font.size = Pt(13)
    run_txt.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

    run_sub = p.add_run("Secretaria de Infraestrutura – SEINFRA")
    run_sub.font.size = Pt(9)
    run_sub.font.color.rgb = RGBColor(0x44, 0x44, 0x44)

    # Linha separadora abaixo do cabeçalho
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), "1F4E79")
    pBdr.append(bottom)
    pPr.append(pBdr)


# ── Rodapé ─────────────────────────────────────────────────────────────────────
def _add_footer(doc):
    """Rodapé com endereço da prefeitura e número de página."""
    section = doc.sections[0]
    footer = section.footer
    footer.is_linked_to_previous = False

    for p in footer.paragraphs:
        p.clear()

    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(0)

    # Linha separadora acima do rodapé
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    top = OxmlElement("w:top")
    top.set(qn("w:val"), "single")
    top.set(qn("w:sz"), "6")
    top.set(qn("w:space"), "4")
    top.set(qn("w:color"), "1F4E79")
    pBdr.append(top)
    pPr.append(pBdr)

    # Texto do endereço
    run_addr = p.add_run(
        "Prefeitura Municipal de Maracanaú  |  "
        "Av. Contorno Sul, s/n – Centro  |  "
        "Maracanaú – CE  |  CEP 61.900-000  |  "
        "www.maracanau.ce.gov.br"
    )
    run_addr.font.size = Pt(7)
    run_addr.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    # Número de página
    p2 = footer.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after = Pt(0)

    run_pg = p2.add_run("Página ")
    run_pg.font.size = Pt(7)
    run_pg.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    # Campo de número de página automático
    fldChar1 = OxmlElement("w:fldChar")
    fldChar1.set(qn("w:fldCharType"), "begin")
    instrText = OxmlElement("w:instrText")
    instrText.text = "PAGE"
    fldChar2 = OxmlElement("w:fldChar")
    fldChar2.set(qn("w:fldCharType"), "end")
    run_num = p2.add_run()
    run_num._r.append(fldChar1)
    run_num._r.append(instrText)
    run_num._r.append(fldChar2)
    run_num.font.size = Pt(7)
    run_num.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    run_of = p2.add_run(" de ")
    run_of.font.size = Pt(7)
    run_of.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    fldChar3 = OxmlElement("w:fldChar")
    fldChar3.set(qn("w:fldCharType"), "begin")
    instrText2 = OxmlElement("w:instrText")
    instrText2.text = "NUMPAGES"
    fldChar4 = OxmlElement("w:fldChar")
    fldChar4.set(qn("w:fldCharType"), "end")
    run_total = p2.add_run()
    run_total._r.append(fldChar3)
    run_total._r.append(instrText2)
    run_total._r.append(fldChar4)
    run_total.font.size = Pt(7)
    run_total.font.color.rgb = RGBColor(0x66, 0x66, 0x66)


# ── Marca d'água ───────────────────────────────────────────────────────────────
def _add_watermark(doc):
    """Adiciona o brasão como marca d'água semitransparente em todas as páginas."""
    if not os.path.exists(_BRASAO):
        return

    # Adiciona a imagem à parte de relacionamentos do documento
    part = doc.part
    image_part = part.new_part(
        "/word/media/watermark.png",
        "image/png",
        open(_BRASAO, "rb").read(),
    )
    rId = part.relate_to(
        image_part,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    )

    # Monta o XML do shape flutuante (marca d'água padrão Word)
    # 7200000 EMU ≈ 8 cm — tamanho central na página
    WMK_XML = f"""
    <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:pPr><w:jc w:val="center"/></w:pPr>
      <w:r>
        <w:rPr><w:noProof/></w:rPr>
        <w:pict>
          <v:shape xmlns:v="urn:schemas-microsoft-com:vml"
                   xmlns:o="urn:schemas-microsoft-com:office:office"
                   xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                   type="#_x0000_t75"
                   style="position:absolute;left:0;top:0;width:180pt;height:180pt;z-index:-251659264;mso-position-horizontal:center;mso-position-vertical:center"
                   o:allowincell="f">
            <v:imagedata r:id="{rId}" o:title="brasao" gain="19661f" blacklevel="22938f"/>
          </v:shape>
        </w:pict>
      </w:r>
    </w:p>"""

    wmk_elem = etree.fromstring(WMK_XML)

    # Insere no cabeçalho (assim aparece em todas as páginas, atrás do conteúdo)
    section = doc.sections[0]
    hdr = section.header
    hdr._element.insert(0, wmk_elem)


# ── Geração principal ──────────────────────────────────────────────────────────
def gerar(dados: dict) -> bytes:
    """
    Recebe o dict de dados (da extração + campos do usuário) e
    retorna os bytes do arquivo DOCX.
    """
    doc = Document()

    # Margens — top maior para não sobrepor o cabeçalho
    for section in doc.sections:
        section.top_margin    = Cm(3.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin   = Cm(2)
        section.right_margin  = Cm(2)
        section.header_distance = Cm(1.2)
        section.footer_distance = Cm(1.0)

    # Fonte padrão
    doc.styles["Normal"].font.name = "Arial"
    doc.styles["Normal"].font.size = Pt(9)

    # Cabeçalho, rodapé e marca d'água
    _add_header(doc)
    _add_footer(doc)
    _add_watermark(doc)

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
        ("PREGÃO DIGITAL",        dados.get("pregao",   "10.002/2024")),
        ("ATA DE REG. DE PREÇOS", dados.get("ata",      "10.004/2024")),
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

        for row in tbl.rows:
            for i, w in enumerate(WIDTHS):
                row.cells[i].width = w

        _header_row(tbl.rows[0], HDR)

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
                center = i not in (2,)
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
