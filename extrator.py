"""
Extração de dados da Ordem de Fornecimento (PDF).

Estratégia dupla:
  1. Texto nativo (PyMuPDF) – funciona para PDFs digitais
  2. OCR com pytesseract – fallback para PDFs escaneados
"""

import re
import fitz  # PyMuPDF
from typing import Optional

# ── OCR: tenta Windows OCR → pytesseract → falha silenciosa ───────────────────
def _ocr_page(page: fitz.Page) -> str:
    """Renderiza a página como imagem e faz OCR."""

    mat = fitz.Matrix(3, 3)   # 216 dpi
    pix = page.get_pixmap(matrix=mat)
    img_bytes = pix.tobytes("png")

    # 1. Windows OCR (disponível nativamente no Windows 10/11)
    try:
        return _windows_ocr(img_bytes)
    except Exception:
        pass

    # 2. pytesseract (precisa do Tesseract instalado — disponível no Streamlit Cloud)
    try:
        import pytesseract
        from PIL import Image
        import io
        img = Image.open(io.BytesIO(img_bytes))
        return pytesseract.image_to_string(img, lang="por")
    except Exception:
        pass

    return ""


def _windows_ocr(img_bytes: bytes) -> str:
    """OCR via API nativa do Windows — não precisa de instalação extra."""
    import asyncio, tempfile, os

    async def _run(path: str) -> str:
        from winrt.windows.media.ocr import OcrEngine
        from winrt.windows.globalization import Language
        from winrt.windows.graphics.imaging import BitmapDecoder
        from winrt.windows.storage import StorageFile

        lang = Language("pt-BR")
        engine = OcrEngine.try_create_from_language(lang)
        file = await StorageFile.get_file_from_path_async(path)
        stream = await file.open_async(0)
        decoder = await BitmapDecoder.create_async(stream)
        bmp = await decoder.get_software_bitmap_async()
        result = await engine.recognize_async(bmp)
        return result.text

    tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
    tmp.write(img_bytes)
    tmp.close()
    try:
        return asyncio.run(_run(tmp.name))
    finally:
        os.unlink(tmp.name)


_USEFUL_WORDS = re.compile(
    r"FORNECEDOR|EMPENHO|ORDEM|VALOR|MARCA|CNPJ|DATA|QUANTIDADE",
    re.IGNORECASE,
)

def _get_page_text(page: fitz.Page) -> str:
    """
    Retorna texto nativo quando ele contém conteúdo útil do documento.
    Cai para OCR quando o texto nativo é vazio ou contém só metadados
    (ex: assinatura digital sem o conteúdo visual da página).
    """
    text = page.get_text("text").strip()
    if text and _USEFUL_WORDS.search(text):
        return text
    return _ocr_page(page)


# ── Normalização ───────────────────────────────────────────────────────────────
def _norm(text: str) -> str:
    text = re.sub(r"[ \t]+", " ", text)    # múltiplos espaços → 1
    text = re.sub(r"\n{3,}", "\n\n", text) # quebras excessivas
    return text.strip()


_DASH = r"[\-–—•‣►·•‣�]"   # vários separadores que o OCR pode gerar


# ── Helpers ────────────────────────────────────────────────────────────────────
def _find(pattern: str, text: str, default: str = "", flags=re.IGNORECASE) -> str:
    m = re.search(pattern, text, flags)
    return m.group(1).strip() if m else default


def _findall(pattern: str, text: str) -> list[str]:
    return re.findall(pattern, text, re.IGNORECASE)


# ── Parser de itens ────────────────────────────────────────────────────────────
def _parse_items(full_text: str) -> list[dict]:
    """
    Suporta dois layouts:

    Layout A (tabela com linha separada por item):
      13012 - CIMENTO - 50 KG
      MARCA   UNIDADE   QUANTIDADE   VALOR UNITARIO   VALOR TOTAL
      POTY    SACO      355,00       R$ 32,53         R$ 11.548,15

    Layout B (OCR desestruturado – tudo numa linha):
      13012 - CIMENTO 50 KG MARCA POTY UNIDADE SACO QUANTIDADE 355,00 ...
      (todos os VALOR UNITARIO no final, todos os VALOR TOTAL no final)
    """

    # ── 1. Localiza cabeçalhos de item ────────────────────────────────────────
    header_re = re.compile(
        rf"\b(\d{{4,6}})\s*{_DASH}\s*(.+?)(?=\n|MARCA|\s+\d{{4,6}}\s*{_DASH}|\Z)",
        re.IGNORECASE,
    )
    starts = []
    for m in header_re.finditer(full_text):
        starts.append((m.start(), m.group(1), m.group(2).strip()))

    if not starts:
        return []

    # ── 2. Coleta listas globais (fallback para Layout B) ─────────────────────
    all_qtd    = _findall(r"QUANTIDADE\s+([\d.,]+)", full_text)
    all_vunit  = _findall(r"VALOR\s+UNIT[A-Z.]*\s+R\$\s*([\d.,]+)", full_text)
    all_vtotal = _findall(r"VALOR\s+TOTAL\s+(?:R\$|RS)?\s*([\d.,'OoIl,]+)", full_text)

    # Limpa OCR comum: "OO" → "00", "l" ou "I" → "1" nos valores
    def clean_val(v: str) -> str:
        v = v.replace("OO", "00").replace("O", "0").replace("l", "1").replace("I", "1")
        v = v.replace(" ", "")
        return v

    all_vtotal = [clean_val(v) for v in all_vtotal]

    items = []
    for idx, (pos, codigo, descricao_raw) in enumerate(starts):
        end_pos = starts[idx + 1][0] if idx + 1 < len(starts) else len(full_text)
        bloco = full_text[pos:end_pos]

        # Limpa descrição
        descricao = re.sub(
            r"\s*(MARCA|UNIDADE|QUANTIDADE|VALOR|Total).*", "",
            descricao_raw, flags=re.IGNORECASE
        ).strip()

        # Marca e unidade buscadas no bloco
        marca   = _find(r"MARCA\s+(\S+)", bloco)
        unidade_raw = _find(r"UNIDADE\s+(\S+(?:\s+\S+){0,3})", bloco)
        # Se unidade começa com dígito → é "UNIDADE 1.0 UNIDADE" → simplifica
        unidade = "SACO" if re.search(r"\bSACO\b", bloco, re.I) else (
                  "UNIDADE" if (not unidade_raw or re.match(r"^\d", unidade_raw))
                  else unidade_raw.split()[0].upper()
        )

        # Quantidade
        qtd = _find(r"QUANTIDADE\s+([\d.,]+)", bloco)
        if not qtd and idx < len(all_qtd):
            qtd = all_qtd[idx]

        # Valor unitário
        v_unit = _find(r"VALOR\s+UNIT[A-Z.]*\s+R\$\s*([\d.,]+)", bloco)
        if not v_unit and idx < len(all_vunit):
            v_unit = all_vunit[idx]

        # Valor total (pega ÚLTIMO match no bloco para evitar pegar do item anterior)
        vtotal_matches = re.findall(
            r"VALOR\s+TOTAL\s+(?:R\$|RS)?\s*([\d.,' OoIl]+)", bloco, re.IGNORECASE
        )
        v_total = clean_val(vtotal_matches[-1]) if vtotal_matches else ""
        if not v_total and idx < len(all_vtotal):
            v_total = all_vtotal[idx]

        items.append({
            "codigo":         codigo,
            "descricao":      descricao,
            "marca":          marca,
            "unidade":        unidade,
            "quantidade":     qtd,
            "valor_unitario": v_unit,
            "valor_total":    v_total,
        })

    return items


# ── Extração principal ─────────────────────────────────────────────────────────
def extrair(pdf_bytes: bytes) -> dict:
    """
    Recebe bytes do PDF e retorna dict com todos os campos extraídos.
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    # Junta texto de TODAS as páginas
    pages_text: list[str] = []
    for page in doc:
        pages_text.append(_get_page_text(page))

    # Página 1 = cabeçalho da ordem; demais = itens + detalhes
    p1 = _norm(pages_text[0]) if pages_text else ""
    full = _norm("\n".join(pages_text))

    # ── Campos do cabeçalho ───────────────────────────────────────────────────
    numero_of = _find(r"ORDEM\s+DE\s+FORNECIMENTO\s*N[º°?P:oO]*\s*[:\s]?\s*(\d+)", full)
    data_of   = _find(r"DATA[:\s]+(\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4})", p1)
    empenho   = _find(r"EMPENHO[:\s]+([A-Z0-9./\-]+)", p1)

    # Fornecedor — linha após "FORNECEDOR:"
    fornecedor = _find(
        r"FORNECEDOR[:\s]+(.+?)(?:\n|CNPJ|ENDERE[Ç C])", p1
    )
    fornecedor = re.sub(r"^FORNECEDOR[:\s]*", "", fornecedor, flags=re.I).strip()

    # CNPJ — aceita com ou sem pontuação
    cnpj = _find(r"CNPJ[:\s]+([\d]{2}[\s.][\d]{3}[\s.][\d]{3}[/\\][\d]{4}[-][\d]{2})", p1)
    if not cnpj:
        cnpj = _find(r"CNPJ[:\s]+([\d./\- ]{14,20})", p1)

    valor_of = _find(r"VALOR\s+DA\s+ORDEM.*?R\$\s*([\d.,]+)", p1)
    prazo    = _find(r"PRAZO\s+DE\s+ENTREGA[:\s]+(\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4})", p1)
    setor    = _find(r"SETOR\s+SOLICITANTE[:\s]+(.+?)(?:\n|ALMOX)", p1)
    almox    = _find(r"ALMOXARIFADO[:\s]+(.+?)(?:\n|ENDERE)", p1)
    endereco = _find(r"ENDERE[CÇ]O\s+DE\s+ENTREGA[:\s]+(.+?)(?:\n|PRAZO)", p1)

    # ── Itens (todas as páginas após a 1ª) ────────────────────────────────────
    # Texto de itens começa na página 2 em diante
    items_text = _norm("\n".join(pages_text[1:])) if len(pages_text) > 1 else ""
    # Se a pág 1 também contiver itens, adiciona
    if re.search(r"VALOR\s+UNIT", p1, re.I):
        items_text = full

    items = _parse_items(items_text if items_text else full)

    return {
        "numero_of":  numero_of,
        "data_of":    data_of,
        "empenho":    empenho,
        "fornecedor": fornecedor,
        "cnpj":       cnpj,
        "valor_of":   valor_of,
        "prazo":      prazo,
        "setor":      setor,
        "almox":      almox,
        "endereco_entrega": endereco,
        "items":      items,
    }
