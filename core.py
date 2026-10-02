"""Relatório de revisitas com histórico, próximo dia e restrições por trecho."""
from datetime import datetime, time, date, timedelta
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


def ultima_visita(value):
    turno = periodo(value)
    if isinstance(value, (datetime, time, pd.Timestamp)):
        horario = value.strftime("%H:%M")
    elif isinstance(value, (int, float)) and 0 <= value < 1:
        minutos = round(float(value) * 24 * 60)
        horario = f"{minutos // 60 % 24:02d}:{minutos % 60:02d}"
    else:
        match = re.fullmatch(r"(\d{1,2})[:hH](\d{2})(?::\d{2})?", texto(value))
        horario = f"{int(match[1]):02d}:{match[2]}" if match and turno else ""
    if horario and turno:
        return f"{horario} - {turno}"
    return turno or texto(value) or "Não informado"


DIAS = ["segunda-feira", "terça-feira", "quarta-feira", "quinta-feira", "sexta-feira", "sábado", "domingo"]


def dia_visita(value):
    """Aceita datas do Excel, datas brasileiras e nomes dos dias sem inventar datas."""
    if not turno_preenchido(value):
        return None, None
    if isinstance(value, (datetime, date, pd.Timestamp)):
        d = value.date() if isinstance(value, datetime) else value
        return d, d.weekday()
    n = normalizar(value).replace("-FEIRA", "")
    for i, nome in enumerate(DIAS):
        if n == normalizar(nome).replace("-FEIRA", ""):
            return None, i
    if isinstance(value, (int, float)) and 1 <= value <= 100000:
        d = date(1899, 12, 30) + timedelta(days=int(value))
        return d, d.weekday()
    try:
        d = pd.to_datetime(texto(value), dayfirst=True, errors="raise").date()
        return d, d.weekday()
    except (ValueError, TypeError, OverflowError):
        return None, None


def proximo_dia(value):
    d, indice = dia_visita(value)
    if indice is None:
        return "confirmar dia"
    salto = 2 if indice == 5 else 1
    seguinte = (indice + salto) % 7
    return (f"{(d + timedelta(days=salto)):%d/%m/%Y} ({DIAS[seguinte]})"
            if d else DIAS[seguinte])


def numero_trecho(value):
    match = re.fullmatch(r"(?:TRECHO\s*)?(\d+)(?:\.0)?", normalizar(value))
    return int(match[1]) if match else None


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
        visitas = []
        for i, (col_turno, col_dia) in enumerate(zip(
                ["TURNO", "TURNO2", "TURNO4", "TURNO6"],
                ["DIA", "DIA3", "DIA5", "DIA7"]), 1):
            valor, dia = row.get(col_turno), row.get(col_dia)
            if turno_preenchido(valor) or turno_preenchido(dia):
                visitas.append((i, valor, dia))
        ultimo = visitas[-1][1] if visitas else None
        ultimo_dia = visitas[-1][2] if visitas else None
        if not turno_preenchido(ultimo_dia) and isinstance(ultimo, datetime):
            ultimo_dia = ultimo
        anterior = periodo(ultimo)
        proximo = {"Manhã": "Tarde", "Tarde": "Manhã"}.get(anterior)
        tentativa = re.fullmatch(r"AUSENTE\s*(\d+)", normalizar(row["STATUS"]))
        numero_visita = int(tentativa[1]) + 1 if tentativa else (visitas[-1][0] + 1 if visitas else None)
        trecho = numero_trecho(row.get("TRECHO"))
        if trecho in range(15, 20) and anterior == "Noite":
            proximo = "Manhã"
        _, indice_dia = dia_visita(ultimo_dia)
        if indice_dia == 4 and proximo:
            proximo = "Manhã ou Tarde"
        sugestao = "Folder (sugestão)" if numero_visita == 3 else (proximo or "Confirmar período")
        if numero_visita == 3:
            proximo = "Folder"
        elif trecho in range(8, 12) and proximo:
            sugestao += {"Manhã": " (08:00 - 11:59)", "Tarde": " (12:00 - 15:00)",
                         "Manhã ou Tarde": " (08:00 - 11:59 ou 12:00 - 15:00)"}[proximo]
        acao = f"{str(numero_visita) + 'ª' if numero_visita else 'Próxima'} visita = {sugestao}; {proximo_dia(ultimo_dia)}"
        historico = []
        for i, valor, dia in visitas:
            if not turno_preenchido(dia) and isinstance(valor, datetime):
                dia = valor
            d, indice = dia_visita(dia)
            dia_texto = f"{d:%d/%m/%Y} ({DIAS[indice]})" if d else DIAS[indice] if indice is not None else "dia não informado"
            historico.append(f"V{i}: {ultima_visita(valor)} - {dia_texto}")
        rows.append({"Endereço": endereco, "Última visita": ultima_visita(ultimo),
                     "Histórico": "\n".join(historico) or "Não informado",
                     "Próxima visita": proximo or "Confirmar", "Recomendação": acao})
    result = pd.DataFrame(rows, columns=["Endereço", "Última visita", "Histórico", "Próxima visita", "Recomendação"])
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
        return Paragraph(escape(texto(value)).replace("\n", "<br/>"), styles[style])
    trechos = result.attrs.get("trechos", [])
    story = [Paragraph("Relatório de Revisitas", styles["Title"])]
    if trechos:
        story.append(p(("Trecho " if len(trechos) == 1 else "Trechos ") + ", ".join(trechos), "Subtitulo"))
    story.extend([Spacer(1, 0.2*cm), p(f"Origem: {origem}"), p(f"Visitas com status Ausente: {len(result)}"), Spacer(1, 0.4*cm)])
    if result.empty:
        story.append(p("Nenhum registro com status Ausente encontrado."))
    else:
        rows = [[p(x, "Cabecalho") for x in ["Endereço", "Histórico", "Próxima visita sugerida"]]]
        for _, row in result.iterrows():
            rows.append([p(row["Endereço"]), p(row["Histórico"]), p(row["Recomendação"])])
        table = LongTable(rows, colWidths=[7*cm, 5*cm, 6*cm], repeatRows=1, hAlign="LEFT")
        table.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,0), colors.HexColor("#146078")), ("VALIGN", (0,0), (-1,-1), "TOP"), ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#f0f5f7")]), ("TOPPADDING", (0,0), (-1,-1), 9), ("BOTTOMPADDING", (0,0), (-1,-1), 9)]))
        story.append(table)
    def footer(canvas, document):
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.grey)
        canvas.drawRightString(19.5*cm, 0.8*cm, f"Página {document.page}")
    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    return out.getvalue()
