#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from flask import Flask, render_template, request, send_file
import requests, csv, io, os
from datetime import datetime, timedelta
from reportlab.pdfgen import canvas
from pypdf import PdfReader, PdfWriter

app = Flask(__name__)

CARTELLA     = os.path.dirname(os.path.abspath(__file__))
PDF_MODULO   = os.path.join(CARTELLA, "modulo_mnp_fastweb.pdf")
PDF_CESSIONE = os.path.join(CARTELLA, "modulo_cessione_sim.pdf")

SHEET1_URL = "https://docs.google.com/spreadsheets/d/1BGVcUswg_gpKalu2lCUBvN3IBtUG2duKccYPC8m-ZgE/export?format=csv"
SHEET2_URL = "https://docs.google.com/spreadsheets/d/1xKjw9TriDsjIl1dSSsyL0KCbAyu0aXHVFVGw459tbjI/export?format=csv"

NOME_RIVENDITORE   = "Bi Zeta srl"
CODICE_RIVENDITORE = "22847.04L20"
MAX_SIM = 8

PAGE_W, PAGE_H = 842.0, 575.0

COORD = {
    "ragione_sociale":  (195, 515),
    "piva":             (55,  503),
    "dom_via":          (170, 490),
    "dom_num":          (776, 490),
    "dom_comune":       (70,  475),
    "dom_cap":          (700, 475),
    "sede_via":         (170, 464),
    "sede_num":         (776, 464),
    "sede_comune":      (70,  453),
    "sede_cap":         (700, 453),
    "nome":             (75,  414),
    "cognome":          (393, 414),
    "cod_fiscale":      (648, 414),
    "tel_cellulare":    (680, 381),
    "operatore":        (155, 344),
    "data_portabilita": (407, 349),
    "qty_sim":          (778, 344),
    "col_serial_voda":  68,
    "col_intestatario": 241,
    "col_cf_piva":      393,
    "col_serial_imp":   510,
    "col_num_prim":     619,
    "col_num_sec":      673,
    "col_abbonamento":  798,
    "row_y": [267, 250, 234, 217, 200, 184, 167, 150],
    "nome_rivenditore": (455, 150),
    "cod_rivenditore":  (455, 132),
    "data_firma":       (73,  38),
    "data_presunta":    (140, 30),
}


# ─── COORDINATE MODULO CESSIONE SIM (A4 portrait 595×842 pt) ────────────────
COORD_CESSIONE = {
    "luogo_data":       (90,  660),   # riga "(Luogo e Data)"
    "sottoscritto":     (155, 545),   # "Io sottoscritto ___"
    "ragione_delegato": (325, 390),   # dopo "dell'Azienda"
    "piva_delegato":    (85,  371),   # dopo "P IVA"
    "num_x":            192,          # X numeri telefono
    "num_y":            [323, 305, 287, 269],  # Y 4 righe numeri
    "ragione_cessione": (57,  228),   # acconsento — ragione sociale
    "piva_cessione":    (455, 227),   # acconsento — P.IVA
}


def val(d, *keys):
    for k in keys:
        v = d.get(k, "")
        if v and str(v).strip():
            return str(v).strip()
    return ""


def carica_sheet(url):
    resp = requests.get(url, allow_redirects=True, timeout=20)
    resp.raise_for_status()
    righe = list(csv.DictReader(io.StringIO(resp.content.decode("utf-8"))))
    return [r for r in righe if any(v.strip() for v in r.values())]


def _split_via_num(indirizzo):
    parti = indirizzo.rsplit(" ", 1)
    if len(parti) == 2 and parti[1].rstrip(".,/").replace("°", "").isdigit():
        return parti[0], parti[1]
    return indirizzo, ""


def _tronca(testo, max_car):
    return testo[:max_car] + "…" if len(testo) > max_car else testo


def _prossimo_lavorativo(data, giorni):
    d, aggiunti = data, 0
    while aggiunti < giorni:
        d += timedelta(days=1)
        if d.weekday() < 5:
            aggiunti += 1
    return d


def crea_pdf(cliente, sim_list, data_port_str):
    buf = io.BytesIO()
    c   = canvas.Canvas(buf, pagesize=(PAGE_W, PAGE_H))

    def t(key, valore, size=9):
        c.setFont("Helvetica", size)
        c.drawString(*COORD[key], valore)

    def tx(x, y, valore, size=9):
        c.setFont("Helvetica", size)
        c.drawString(x, y, valore)

    rag_soc   = val(cliente, "RAGIONE SOCIALE")
    piva      = val(cliente, "P.IVA").replace(" ", "")
    indirizzo = val(cliente, "INDIRIZZO")
    comune    = val(cliente, "COMUNE")
    cap       = val(cliente, "CAP")
    via, num  = _split_via_num(indirizzo)

    t("ragione_sociale", rag_soc)
    t("piva",            piva)
    t("dom_via",         via);   t("dom_num",    num)
    t("dom_comune",      comune); t("dom_cap",   cap)
    t("sede_via",        via);   t("sede_num",   num)
    t("sede_comune",     comune); t("sede_cap",  cap)
    t("nome",            val(cliente, "NOME"))
    t("cognome",         val(cliente, "COGNOME"))
    t("cod_fiscale",     val(cliente, "CODICE FISCALE"))
    t("tel_cellulare",   val(cliente, "RECAPITO MOBILE", "MOBILE", "RECAPITO"))
    t("operatore",       "1Mobile")
    t("data_portabilita", data_port_str)
    t("qty_sim",         str(len(sim_list)))

    c.setFont("Helvetica", 7)
    for i, sim in enumerate(sim_list[:MAX_SIM]):
        y           = COORD["row_y"][i]
        serial_voda = val(sim, "SIM VODAFONE DEFINITIVA")
        serial_imp  = val(sim, "SIM 1 MOBILE")
        numero      = val(sim, "NUMERO")
        intestat    = val(sim, "H1") or rag_soc

        tx(COORD["col_serial_voda"],  y, serial_voda, size=6)
        tx(COORD["col_intestatario"], y, _tronca(intestat, 28))
        tx(COORD["col_cf_piva"],      y, piva)
        tx(COORD["col_serial_imp"],   y, serial_imp, size=6)
        num_display = numero[1:] if numero.startswith("3") else numero
        tx(COORD["col_num_prim"],     y, num_display)
        tx(COORD["col_abbonamento"],  y, "X", size=8)

    t("nome_rivenditore", NOME_RIVENDITORE)
    t("cod_rivenditore",  CODICE_RIVENDITORE)
    oggi = datetime.today().strftime("%d/%m/%Y")
    t("data_firma",    oggi)
    t("data_presunta", data_port_str)

    c.save()
    buf.seek(0)

    template  = PdfReader(PDF_MODULO)
    overlay_r = PdfReader(buf)
    writer    = PdfWriter()
    pagina    = template.pages[0]
    pagina.merge_page(overlay_r.pages[0])
    writer.add_page(pagina)

    out = io.BytesIO()
    writer.write(out)
    out.seek(0)
    return out


# ─── ROUTES ──────────────────────────────────────────────────────────────────

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html", modulo="mnp")


@app.route("/cerca", methods=["POST"])
def cerca():
    piva = request.form.get("piva", "").strip().replace(" ", "")
    if not piva:
        return render_template("index.html", errore="Inserisci una P.IVA.")
    try:
        clienti = carica_sheet(SHEET1_URL)
        sim_all = carica_sheet(SHEET2_URL)
    except Exception as e:
        return render_template("index.html", errore=f"Errore connessione: {e}")

    risultati = [r for r in clienti if r.get("P.IVA", "").replace(" ", "") == piva]
    if not risultati:
        return render_template("index.html", errore=f"Nessun cliente trovato con P.IVA {piva}.")

    cliente   = risultati[0]
    sim_list  = [r for r in sim_all if r.get("P.IVA", "").replace(" ", "") == piva and r.get("NUMERO", "").strip()]
    data_def  = _prossimo_lavorativo(datetime.today(), 4).strftime("%d/%m/%Y")

    return render_template("cliente.html",
        cliente=cliente, sim_list=sim_list,
        piva=piva, data_default=data_def)


@app.route("/genera", methods=["POST"])
def genera():
    piva       = request.form.get("piva", "").strip()
    data_port  = request.form.get("data_port", "").strip()
    try:
        clienti  = carica_sheet(SHEET1_URL)
        sim_all  = carica_sheet(SHEET2_URL)
    except Exception as e:
        return f"Errore connessione: {e}", 500

    risultati = [r for r in clienti if r.get("P.IVA", "").replace(" ", "") == piva]
    if not risultati:
        return "Cliente non trovato.", 404

    cliente  = risultati[0]
    sim_list = [r for r in sim_all if r.get("P.IVA", "").replace(" ", "") == piva and r.get("NUMERO", "").strip()]

    if not sim_list:
        return "Nessuna SIM trovata per questo cliente.", 404

    gruppi    = [sim_list[i:i+MAX_SIM] for i in range(0, len(sim_list), MAX_SIM)]
    nome_base = val(cliente, "RAGIONE SOCIALE").replace("/", "-").replace("\\", "-").replace(":", " ").strip()
    data_str  = datetime.today().strftime("%Y%m%d")

    if len(gruppi) == 1:
        pdf_buf   = crea_pdf(cliente, gruppi[0], data_port)
        nome_file = f"MNP_{nome_base}_{data_str}.pdf"
        return send_file(pdf_buf, as_attachment=True,
                         download_name=nome_file, mimetype="application/pdf")

    # Più di 8 SIM → ZIP con più PDF
    import zipfile
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w") as zf:
        for idx, gruppo in enumerate(gruppi, 1):
            pdf_buf   = crea_pdf(cliente, gruppo, data_port)
            nome_file = f"MNP_{nome_base}_{data_str}_{idx}di{len(gruppi)}.pdf"
            zf.writestr(nome_file, pdf_buf.read())
    zip_buf.seek(0)
    return send_file(zip_buf, as_attachment=True,
                     download_name=f"MNP_{nome_base}_{data_str}.zip",
                     mimetype="application/zip")


# ─── CESSIONE SIM ─────────────────────────────────────────────────────────────

def crea_pdf_cessione(cliente, sim_list):
    buf = io.BytesIO()
    c   = canvas.Canvas(buf, pagesize=(595, 842))

    nome_cognome = f"{val(cliente, 'NOME')} {val(cliente, 'COGNOME')}".strip()
    ragione      = val(cliente, "RAGIONE SOCIALE")
    piva         = val(cliente, "P.IVA").replace(" ", "")
    oggi         = datetime.today().strftime("%d/%m/%Y")

    def t(key, valore, size=9):
        c.setFont("Helvetica", size)
        c.drawString(*COORD_CESSIONE[key], valore)

    def tx(x, y, valore, size=9):
        c.setFont("Helvetica", size)
        c.drawString(x, y, valore)

    t("luogo_data",       f"Messina, {oggi}")
    t("sottoscritto",     nome_cognome)
    t("ragione_delegato", ragione)
    t("piva_delegato",    piva)
    t("ragione_cessione", ragione)
    t("piva_cessione",    piva)

    numeri = [val(s, "NUMERO") for s in sim_list if val(s, "NUMERO")]
    for i, numero in enumerate(numeri[:4]):
        tx(COORD_CESSIONE["num_x"], COORD_CESSIONE["num_y"][i], numero)

    c.save()
    buf.seek(0)

    template  = PdfReader(PDF_CESSIONE)
    overlay_r = PdfReader(buf)
    writer    = PdfWriter()
    pagina    = template.pages[0]
    pagina.merge_page(overlay_r.pages[0])
    writer.add_page(pagina)

    out = io.BytesIO()
    writer.write(out)
    out.seek(0)
    return out


@app.route("/cessione", methods=["GET"])
def cessione():
    return render_template("index.html", modulo="cessione")


@app.route("/cerca_cessione", methods=["POST"])
def cerca_cessione():
    piva = request.form.get("piva", "").strip().replace(" ", "")
    if not piva:
        return render_template("index.html", modulo="cessione", errore="Inserisci una P.IVA.")
    try:
        clienti = carica_sheet(SHEET1_URL)
        sim_all = carica_sheet(SHEET2_URL)
    except Exception as e:
        return render_template("index.html", modulo="cessione", errore=f"Errore connessione: {e}")

    risultati = [r for r in clienti if r.get("P.IVA", "").replace(" ", "") == piva]
    if not risultati:
        return render_template("index.html", modulo="cessione",
                               errore=f"Nessun cliente trovato con P.IVA {piva}.")

    cliente  = risultati[0]
    sim_list = [r for r in sim_all if r.get("P.IVA", "").replace(" ", "") == piva
                and r.get("NUMERO", "").strip()]

    return render_template("cessione_cliente.html",
                           cliente=cliente, sim_list=sim_list, piva=piva)


@app.route("/genera_cessione", methods=["POST"])
def genera_cessione():
    piva = request.form.get("piva", "").strip()
    try:
        clienti = carica_sheet(SHEET1_URL)
        sim_all = carica_sheet(SHEET2_URL)
    except Exception as e:
        return f"Errore connessione: {e}", 500

    risultati = [r for r in clienti if r.get("P.IVA", "").replace(" ", "") == piva]
    if not risultati:
        return "Cliente non trovato.", 404

    cliente  = risultati[0]
    sim_list = [r for r in sim_all if r.get("P.IVA", "").replace(" ", "") == piva
                and r.get("NUMERO", "").strip()]

    ragione   = val(cliente, "RAGIONE SOCIALE").replace("/", "-").replace("\\", "-").replace(":", " ").strip()
    data_str  = datetime.today().strftime("%Y%m%d")
    pdf_buf   = crea_pdf_cessione(cliente, sim_list)
    nome_file = f"CESSIONE_{ragione}_{data_str}.pdf"

    return send_file(pdf_buf, as_attachment=True,
                     download_name=nome_file, mimetype="application/pdf")


# ─── DICHIARAZIONE CLIENTE ────────────────────────────────────────────────────

PDF_DICHIARAZIONE = os.path.join(CARTELLA, "modulo_dichiarazione_cliente.pdf")

COORD_DICH = {
    "nome_cognome": (170, 697),
    "ragione":      (120, 632),
    "piva":         (180, 592),
    "data":         (230, 322),
}
FONT_DICH     = 30
FONT_DICH_RAG = 16


def crea_pdf_dichiarazione(cliente):
    buf  = io.BytesIO()
    c    = canvas.Canvas(buf, pagesize=(595, 842))

    nome  = f"{val(cliente, 'NOME')} {val(cliente, 'COGNOME')}".strip()
    rag   = val(cliente, "RAGIONE SOCIALE")
    piva  = val(cliente, "P.IVA").replace(" ", "")
    oggi  = datetime.today().strftime("%d/%m/%Y")

    c.setFont("Helvetica", FONT_DICH_RAG)
    c.drawString(*COORD_DICH["nome_cognome"], nome)
    c.drawString(*COORD_DICH["ragione"],      rag)
    c.setFont("Helvetica", FONT_DICH)
    c.drawString(*COORD_DICH["piva"],         piva)
    c.drawString(*COORD_DICH["data"],         oggi)

    c.save(); buf.seek(0)

    template  = PdfReader(PDF_DICHIARAZIONE)
    overlay_r = PdfReader(buf)
    writer    = PdfWriter()
    pagina    = template.pages[0]
    pagina.merge_page(overlay_r.pages[0])
    writer.add_page(pagina)

    out = io.BytesIO()
    writer.write(out)
    out.seek(0)
    return out


@app.route("/dichiarazione", methods=["GET"])
def dichiarazione():
    return render_template("index.html", modulo="dichiarazione")


@app.route("/cerca_dichiarazione", methods=["POST"])
def cerca_dichiarazione():
    piva = request.form.get("piva", "").strip().replace(" ", "")
    if not piva:
        return render_template("index.html", modulo="dichiarazione", errore="Inserisci una P.IVA.")
    try:
        clienti = carica_sheet(SHEET1_URL)
    except Exception as e:
        return render_template("index.html", modulo="dichiarazione", errore=f"Errore connessione: {e}")

    risultati = [r for r in clienti if r.get("P.IVA", "").replace(" ", "") == piva]
    if not risultati:
        return render_template("index.html", modulo="dichiarazione",
                               errore=f"Nessun cliente trovato con P.IVA {piva}.")

    cliente = risultati[0]
    return render_template("dichiarazione_cliente.html", cliente=cliente, piva=piva)


@app.route("/genera_dichiarazione", methods=["POST"])
def genera_dichiarazione():
    piva = request.form.get("piva", "").strip()
    try:
        clienti = carica_sheet(SHEET1_URL)
    except Exception as e:
        return f"Errore connessione: {e}", 500

    risultati = [r for r in clienti if r.get("P.IVA", "").replace(" ", "") == piva]
    if not risultati:
        return "Cliente non trovato.", 404

    cliente   = risultati[0]
    pdf_buf   = crea_pdf_dichiarazione(cliente)
    ragione   = val(cliente, "RAGIONE SOCIALE").replace("/", "-").replace("\\", "-").replace(":", " ").strip()
    nome_file = f"DICHIARAZIONE_{ragione}_{datetime.today().strftime('%Y%m%d')}.pdf"

    return send_file(pdf_buf, as_attachment=True,
                     download_name=nome_file, mimetype="application/pdf")


if __name__ == "__main__":
    app.run(debug=False, host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
