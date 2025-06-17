from flask import Flask, render_template, request, send_file
import pandas as pd
import random
import fitz  # PyMuPDF
import io
from datetime import datetime

app = Flask(__name__)

# Lade die Fragen aus der Excel-Datei
def lade_fragen():
    df_dict = pd.read_excel("Checkliste.xlsx", sheet_name=None, engine="openpyxl")
    fragen = []
    for kategorie, df in df_dict.items():
        for frage in df.iloc[:, 0].dropna():
            fragen.append((kategorie, frage))
    return fragen

@app.route("/", methods=["GET", "POST"])
def formular():
    if request.method == "POST":
        datum = request.form.get("datum")
        bereich = request.form.get("bereich")
        fuehrungskraft = request.form.get("fuehrungskraft")

        fragen = lade_fragen()
        ausgewaehlt = random.sample(fragen, 10)

        antworten = []
        for i, (kategorie, frage) in enumerate(ausgewaehlt):
            antwort = request.form.get(f"antwort_{i}")
            bemerkung = request.form.get(f"bemerkung_{i}")
            antworten.append((kategorie, frage, antwort, bemerkung))

        # Speichere in Excel-Übersicht
        eintrag = {
            "Datum": datum,
            "Bereich": bereich,
            "Führungskraft": fuehrungskraft,
            "Zeitstempel": datetime.now().isoformat()
        }
        for i, (kat, frage, antw, bemerkung) in enumerate(antworten):
            eintrag[f"Frage_{i+1}"] = frage
            eintrag[f"Antwort_{i+1}"] = antw
            eintrag[f"Bemerkung_{i+1}"] = bemerkung
        try:
            df_alt = pd.read_excel("antworten_uebersicht.xlsx", engine="openpyxl")
            df_neu = pd.concat([df_alt, pd.DataFrame([eintrag])], ignore_index=True)
        except FileNotFoundError:
            df_neu = pd.DataFrame([eintrag])
        df_neu.to_excel("antworten_uebersicht.xlsx", index=False)

        # PDF generieren
        pdf_stream = io.BytesIO()
        doc = fitz.open()
        seite = doc.new_page()

        text = f"Safety-Checkliste für Führungskräfte\n\nDatum: {datum}\nBereich: {bereich}\nFührungskraft: {fuehrungskraft}\n\n"
        for i, (kat, frage, antw, bemerkung) in enumerate(antworten):
            text += f"{i+1}. [{kat}] {frage}\nAntwort: {antw or '-'}\nBemerkung: {bemerkung or '-'}\n\n"

        seite.insert_text((50, 50), text, fontsize=11)
        doc.save(pdf_stream)
        pdf_stream.seek(0)
        return send_file(pdf_stream, as_attachment=True, download_name="Safety_Checkliste.pdf", mimetype="application/pdf")

    else:
        fragen = lade_fragen()
        ausgewaehlt = random.sample(fragen, 10)
        return render_template("form.html", fragen=ausgewaehlt)

if __name__ == "__main__":
    app.run(debug=True)
