"""Relatório de revisitas baseado no último turno preenchido."""
from datetime import datetime, time
from io import BytesIO
import re
import unicodedata
from xml.sax.saxutils import escape
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, LongTable, TableStyle


def texto(value):
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def normalizar(value):
    s = unicodedata.normalize("NFKD", texto(value)).encode("ascii", "ignore").decode().upper()
    return re.sub(r"\s+", " ", s).strip()


def nome_coluna(value):
    return re.sub(r"[^A-Z0-9]", "", normalizar(value))


def status_ausente(value):
    return re.fullmatch(r"AUSENTE(?:\s*\d+)?", normalizar(value)) is not None


def turno_preenchido(value):
    return texto(value) not in {"", "-", "–", "—"}


def periodo(value):
    n = normalizar(value)
    if n in {"MANHA", "MATUTINO", "M"}:
        return "Manhã"
    if n in {"TARDE", "VESPERTINO", "T"}:
        return "Tarde"
    if n in {"NOITE", "NOTURNO", "N"}:
        return "Noite"
    if isinstance(value, (datetime, time, pd.Timestamp)):
        hour = value.hour
    elif isinstance(value, (int, float)) and 0 <= value < 1:
        hour = int(value * 24)
    else:
        match = re.fullmatch(r"(\d{1,2})[:hH](\d{2})(?::\d{2})?", texto(value))
        if not match or int(match[1]) > 23 or int(match[2]) > 59:
            return ""
        hour = int(match[1])
    return "Manhã" if hour < 12 else "Tarde" if hour < 18 else "Noite"


def ler_excel(content):
    with pd.ExcelFile(BytesIO(content), engine="openpyxl") as book:
        sheet = next((s for s in book.sheet_names if normalizar(s) == "CAMPO"), book.sheet_names[0])
        raw = pd.read_excel(book, sheet_name=sheet, header=None, dtype=object)
    # Reconhece o cabeçalho mesmo quando há um título acima da tabela.
    header = next((i for i in range(min(30, len(raw))) if {"CHAVEVISITA", "LOGRADOURO", "STATUS"}.issubset({nome_coluna(v) for v in raw.iloc[i]})), None)
    if header is None:
        raise ValueError("Não encontrei as colunas CHAVE_VISITA, LOGRADOURO e STATUS na aba CAMPO ou na primeira aba.")
    df = raw.iloc[header+1:].copy()
    df.columns = [nome_coluna(v) or f"SEM_TITULO_{i}" for i, v in enumerate(raw.iloc[header])]
    if df.columns.duplicated().any():
        raise ValueError("Existem títulos de colunas duplicados na planilha.")
    return df.dropna(how="all"), sheet


def preparar(df):
    df = df.copy()
    df.columns = [nome_coluna(c) for c in df.columns]
    required = {"CHAVEVISITA", "LOGRADOURO", "STATUS"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError("Colunas obrigatórias ausentes: " + ", ".join(sorted(missing)))
    turnos = [c for c in ["TURNO", "TURNO2", "TURNO4", "TURNO6"] if c in df.columns]
    if not turnos:
        raise ValueError("Não encontrei as colunas Turno, Turno2, Turno4 ou Turno6.")
    rows = []
    for _, row in df.iterrows():
        if not status_ausente(row["STATUS"]):
            continue
        numero = next((texto(row[c]) for c in ["N", "NO", "NUMERO"] if c in df.columns and texto(row[c])), "")
        endereco = ", ".join(filter(None, [texto(row["LOGRADOURO"]), numero, texto(row.get("COMPLEMENTO")), texto(row.get("MUNICIPIO"))])) or "Endereço não informado"
        # A ordem das colunas representa a sequência das tentativas, sem usar os dias.
        ultimo = next((row[c] for c in reversed(turnos) if turno_preenchido(row[c])), None)
        anterior = periodo(ultimo)
        proximo = {"Manhã": "Tarde", "Tarde": "Manhã"}.get(anterior)
        tentativa = re.fullmatch(r"AUSENTE\s*(\d+)", normalizar(row["STATUS"]))
        numero_visita = int(tentativa[1]) + 1 if tentativa else None
        realizar = f"Realizar {numero_visita}ª visita" if numero_visita else "Realizar próxima visita"
        acao = f"{realizar} no período da {proximo.lower()}." if proximo else f"{realizar}; confirmar o período da última visita."
        rows.append({"Endereço": endereco, "Última visita": anterior or (texto(ultimo) if ultimo is not None else "Não informado"), "Próxima visita": proximo or "Confirmar", "Recomendação": acao})
    result = pd.DataFrame(rows, columns=["Endereço", "Última visita", "Próxima visita", "Recomendação"])
    if "TRECHO" in df.columns:
        ausentes = df[df["STATUS"].map(status_ausente)]
        origem_trechos = ausentes if not ausentes.empty else df
        result.attrs["trechos"] = list(dict.fromkeys(texto(v) for v in origem_trechos["TRECHO"] if texto(v)))
    return result


def gerar_pdf(result, origem):
    out = BytesIO()
    doc = SimpleDocTemplate(out, pagesize=A4, leftMargin=1.5*cm, rightMargin=1.5*cm, topMargin=1.5*cm, bottomMargin=1.5*cm)
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="Subtitulo", fontName="Helvetica", fontSize=13, leading=17, alignment=1, spaceAfter=6))
    styles.add(ParagraphStyle(name="Celula", fontName="Helvetica", fontSize=9, leading=12, wordWrap="CJK"))
    styles.add(ParagraphStyle(name="Cabecalho", parent=styles["Celula"], fontName="Helvetica-Bold", textColor=colors.white))
    def p(value, style="Celula"):
        return Paragraph(escape(texto(value)), styles[style])
    trechos = result.attrs.get("trechos", [])
    story = [Paragraph("Relatório de Revisitas", styles["Title"])]
    if trechos:
        story.append(p(("Trecho " if len(trechos) == 1 else "Trechos ") + ", ".join(trechos), "Subtitulo"))
    story.extend([Spacer(1, 0.2*cm), p(f"Origem: {origem}"), p(f"Visitas com status Ausente: {len(result)}"), Spacer(1, 0.4*cm)])
    if result.empty:
        story.append(p("Nenhum registro com status Ausente encontrado."))
    else:
        rows = [[p(x, "Cabecalho") for x in ["Endereço", "Última visita", "Próxima visita sugerida"]]]
        for _, row in result.iterrows():
            rows.append([p(row["Endereço"]), p(row["Última visita"]), p(row["Recomendação"])])
        table = LongTable(rows, colWidths=[8.5*cm, 2.8*cm, 6.7*cm], repeatRows=1, hAlign="LEFT")
        table.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,0), colors.HexColor("#146078")), ("VALIGN", (0,0), (-1,-1), "TOP"), ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#f0f5f7")]), ("TOPPADDING", (0,0), (-1,-1), 9), ("BOTTOMPADDING", (0,0), (-1,-1), 9)]))
        story.append(table)
    def footer(canvas, document):
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.grey)
        canvas.drawRightString(19.5*cm, 0.8*cm, f"Página {document.page}")
    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    return out.getvalue()
