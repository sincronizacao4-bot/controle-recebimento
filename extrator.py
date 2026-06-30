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

    mat = fitz.Matrix(4, 4)   # 288 dpi — mais nítido para tabelas pequenas
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
        from PIL import Image, ImageOps, ImageFilter
        import io
        img = Image.open(io.BytesIO(img_bytes)).convert("L")  # escala de cinza
        img = ImageOps.autocontrast(img)                       # realça contraste
        img = img.filter(ImageFilter.SHARPEN)                  # realça bordas do texto
        config = "--psm 6"  # assume bloco uniforme de texto (melhor p/ tabelas)
        texto = pytesseract.image_to_string(img, lang="por", config=config)
        if not texto.strip():
            # Segunda tentativa sem pré-processamento, caso o contraste tenha piorado a leitura
            img2 = Image.open(io.BytesIO(img_bytes))
            texto = pytesseract.image_to_string(img2, lang="por")
        return texto
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


_DASH = r"[\-–—•‣►·•‣�|]"   # vários separadores que o OCR pode gerar (incl. pipe)


# ── Helpers ────────────────────────────────────────────────────────────────────
def _find(pattern: str, text: str, default: str = "", flags=re.IGNORECASE) -> str:
    m = re.search(pattern, text, flags)
    return m.group(1).strip() if m else default


def _findall(pattern: str, text: str) -> list[str]:
    return re.findall(pattern, text, re.IGNORECASE)


# ── Parser de itens ────────────────────────────────────────────────────────────
def _parse_items(full_text: str) -> list[dict]:
    """
    Tenta 3 estratégias em cascata para extrair os valores de cada item.

    ESTRATÉGIA 1 — Tabela em 2 linhas (Tesseract/Linux):
        13012 - CIMENTO - 50 KG
        MARCA | UNIDADE | QUANTIDADE  VALOR UNITARIO  VALOR TOTAL   ← cabeçalho
        POTY    SACO      355,00       R$ 32,53        R$ 11.548,15  ← valores

    ESTRATÉGIA 2 — Palavras-chave inline (Windows OCR):
        13012 - CIMENTO 50 KG MARCA POTY UNIDADE SACO QUANTIDADE 355,00
        VALOR UNITARIO R$ 32,53 VALOR TOTAL R$ 11.548,15

    ESTRATÉGIA 3 — Listas globais indexadas por posição:
        Quando os valores aparecem todos agrupados no final do texto.
        Usa o índice do item para pegar o valor correspondente.
    """

    def clean_num(v: str) -> str:
        v = str(v).strip()
        v = v.replace("OO","00").replace("O","0").replace("l","1").replace("I","1")
        v = v.replace("S", "5") if re.fullmatch(r"[\dSO.,]+", v) else v
        return v.replace(" ", "")

    # Padrões tolerantes a abreviações e ruído de OCR
    _RE_QTD    = r"(?:QUANTIDADE|QUANT\.?|QTDE?\.?)\s*[:\s]\s*([\d.,]+)"
    _RE_VUNIT  = r"VA?LO?R\s+UNIT[A-ZÁ-Ú]*\.?\s*[:\s]?\s*(?:R\$)?\s*([\d.,]+)"
    _RE_VTOTAL = r"VA?LO?R\s+TO?TA?L\.?\s*[:\s]?\s*(?:R\$)?\s*([\d.,]+)"

    def money_in(text: str) -> list[str]:
        """Retorna todos os valores R$ encontrados no texto, limpos."""
        return [clean_num(m) for m in re.findall(r"R\$\s*([\d.,]+)", text, re.I)]

    def best_unidade(text: str) -> str:
        if re.search(r"\bSACO\b", text, re.I):    return "SACO"
        if re.search(r"\bLITRO\b", text, re.I):   return "LITRO"
        if re.search(r"\bCX\b", text, re.I):       return "CX"
        if re.search(r"\bKG\b", text, re.I):       return "KG"
        if re.search(r"\bMT?\b", text, re.I):      return "M"
        return "UNIDADE"

    # ── Localiza início de cada item ──────────────────────────────────────────
    item_hdr = re.compile(
        rf"^[ \t]*(\d{{4,6}})\s*{_DASH}\s*(.+?)[ \t]*$",
        re.MULTILINE | re.IGNORECASE,
    )
    starts = [(m.start(), m.group(1), m.group(2).strip())
              for m in item_hdr.finditer(full_text)]

    if not starts:
        return []

    # ── Listas globais (Estratégia 3) ─────────────────────────────────────────
    # Captura todos os valores numéricos após QUANTIDADE / VALOR UNIT / VALOR TOTAL
    g_qtd    = [clean_num(v) for v in re.findall(_RE_QTD, full_text, re.I)]
    g_vunit  = [clean_num(v) for v in re.findall(_RE_VUNIT, full_text, re.I)]
    g_vtotal = [clean_num(v) for v in re.findall(_RE_VTOTAL, full_text, re.I)]
    g_marca  = re.findall(r"MARCA\s*[:\s]\s*([A-Za-zÀ-ú][\w\-À-ú]*)", full_text, re.I)

    items = []
    for idx, (pos, codigo, desc_raw) in enumerate(starts):
        end_pos = starts[idx + 1][0] if idx + 1 < len(starts) else len(full_text)
        bloco   = full_text[pos:end_pos]

        # Descrição limpa
        descricao = re.sub(
            r"\s*(MARCA|UNIDADE|QUANTIDADE|VALOR|Total).*", "",
            desc_raw, flags=re.IGNORECASE
        ).strip()

        marca = unidade = qtd = v_unit = v_total = ""

        # ══ ESTRATÉGIA 1: cabeçalho de tabela + linha de valores ══════════════
        col_hdr = re.search(r"^.*MARCA.*(?:QUANTIDADE|QUANT|QTDE?).*$", bloco, re.MULTILINE | re.I)
        if col_hdr:
            after     = bloco[col_hdr.end():]
            val_lines = [l.strip() for l in after.splitlines() if l.strip()]
            val_line  = val_lines[0] if val_lines else ""

            money = money_in(val_line)
            qtd_m = re.search(r"([\d.,]+)\s+R\$", val_line, re.I)

            if qtd_m:
                qtd = clean_num(qtd_m.group(1))
                pre = val_line[:qtd_m.start()].strip()
            else:
                partes_num = re.findall(r"[\d.,]+", val_line.split("R$")[0])
                qtd = clean_num(partes_num[-1]) if partes_num else ""
                pre = re.sub(r"[\d.,]+\s*$", "", val_line.split("R$")[0]).strip()

            v_unit  = money[0] if len(money) >= 1 else ""
            v_total = money[1] if len(money) >= 2 else ""

            parts = pre.split()
            marca   = parts[0] if parts else ""
            unidade = best_unidade(pre + " " + bloco)

        # ══ ESTRATÉGIA 2: palavras-chave inline no bloco ══════════════════════
        if not qtd:
            qtd = clean_num(_find(_RE_QTD, bloco))
        if not v_unit:
            v_unit = clean_num(_find(_RE_VUNIT, bloco))
        if not v_total:
            # Pega o ÚLTIMO match de VALOR TOTAL no bloco (evita pegar do item anterior)
            vt_all = re.findall(_RE_VTOTAL, bloco, re.I)
            v_total = clean_num(vt_all[-1]) if vt_all else ""
        if not marca:
            marca = _find(r"MARCA\s*[:\s]\s*([A-Za-zÀ-ú][\w\-À-ú]*)", bloco)
        if not unidade:
            unidade = best_unidade(bloco)

        # ══ ESTRATÉGIA 4: valores monetários "soltos" no bloco (sem rótulo legível) ══
        # Quando o OCR engole os rótulos mas mantém os números, usa os valores R$
        # encontrados no bloco, na ordem em que aparecem (unitário antes do total).
        if not v_unit or not v_total:
            soltos = money_in(bloco)
            if not v_unit and len(soltos) >= 1:
                v_unit = soltos[0]
            if not v_total and len(soltos) >= 2:
                v_total = soltos[1]

        # ══ ESTRATÉGIA 3: listas globais indexadas ════════════════════════════
        if not qtd     and idx < len(g_qtd):    qtd    = g_qtd[idx]
        if not v_unit  and idx < len(g_vunit):  v_unit = g_vunit[idx]
        if not v_total and idx < len(g_vtotal): v_total = g_vtotal[idx]
        if not marca   and idx < len(g_marca):  marca  = g_marca[idx]
        if not unidade: unidade = "UNIDADE"

        # Calcula VALOR TOTAL quando o OCR não leu (qtd × vunit)
        if not v_total and qtd and v_unit:
            try:
                def _to_float(s):
                    s = clean_num(s)
                    # Formato brasileiro: 1.234,56
                    if "," in s and "." in s:
                        s = s.replace(".", "").replace(",", ".")
                    elif "," in s:
                        s = s.replace(",", ".")
                    return float(s)
                v_total = f"{_to_float(qtd) * _to_float(v_unit):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            except Exception:
                pass

        # Marca como suspeito qualquer item sem os 3 valores numéricos essenciais
        suspeito = not (qtd and v_unit and v_total)

        items.append({
            "codigo":         codigo,
            "descricao":      descricao,
            "marca":          marca,
            "unidade":        unidade,
            "quantidade":     qtd,
            "valor_unitario": v_unit,
            "valor_total":    v_total,
            "_suspeito":      suspeito,
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

    # Objeto — linha de observações/aquisição
    objeto = _find(
        r"(Aquisi[çc][aã]o\s+.+?)(?:\n[A-Z]{2,}|\Z)", full, flags=re.IGNORECASE | re.DOTALL
    )
    objeto = re.sub(r"\s+", " ", objeto).strip()[:400] if objeto else ""

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
        "objeto":     objeto,
        "items":      items,
        "_ocr_text":  full,
    }
