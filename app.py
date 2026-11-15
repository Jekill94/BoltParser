import os
import re
import openpyxl
from flask import Flask, render_template, request, send_file
import pdfplumber

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/estrai', methods=['POST'])
def estrai():
    if 'pdf_files' not in request.files:
        return "Errore: Nessun file caricato."
    
    files = request.files.getlist('pdf_files')
    parola_chiave = request.form.get('filtro', '').strip()
    
    if not files or files[0].filename == '':
        return "Nessun file selezionato."
    
    righe_strutturate = []
    
    for file in files:
        if file and file.filename.endswith('.pdf'):
            percorso_pdf = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(percorso_pdf)
            
            try:
                with pdfplumber.open(percorso_pdf) as pdf:
                    for indice, pagina in enumerate(pdf.pages):
                        num_pagina = indice + 1
                        testo_pagina = pagina.extract_text()
                        
                        if testo_pagina:
                            for riga in testo_pagina.split('\n'):
                                riga_pulita = riga.strip()
                                if not riga_pulita:
                                    continue
                                
                                riga_lower = riga_pulita.lower()
                                
                                # --- FILTRO ANTI-SPOPOLAMENTO / SCARTO INTELLIGENTE ---
                                # Esclude intestazioni, piè di pagina, metadati del PDF o righe descrittive fittizie
                                if any(parola in riga_lower for parola in ["pagina", "elettronica", "descrizione articolo", "p.iva", "codice fiscale", "spett.le", "fattura n"]):
                                    continue
                                    
                                if "test" in riga_lower and not any(ch.isdigit() for ch in riga_lower):
                                    continue

                                if parola_chiave and parola_chiave.lower() not in riga_lower:
                                    continue

                                # 1. Pattern principale per righe con quantità esplicita (es: Descrizione 10 50,00 EUR 500,00 EUR)
                                match = re.search(r'^(.*?)\s+(\d+)\s+([\d\.,]+)\s*([A-Za-z€$]+)\s+([\d\.,]+)\s*([A-Za-z€$]+)?$', riga_pulita)
                                
                                if match:
                                    desc = match.group(1).strip()
                                    qta = match.group(2).strip()
                                    pz = match.group(3).strip()
                                    val_pz = match.group(4).strip()
                                    tot = match.group(5).strip()
                                    val_tot = match.group(6).strip() if match.group(6) else val_pz
                                    
                                    # Pulizia ulteriore da eventuali residui di intestazione nel nome della descrizione
                                    if "fattura" in desc.lower() or len(desc) < 2:
                                        continue
                                        
                                    righe_strutturate.append((
                                        file.filename, num_pagina, 
                                        qta, desc, 
                                        pz, val_pz, 
                                        tot, val_tot
                                    ))
                                else:
                                    # 2. Pattern alternativo per righe senza quantità esplicita a sinistra (quantità default = 1)
                                    match_alt = re.search(r'^(.*?)\s+([\d\.,]+)\s*([A-Za-z€$]+)\s+([\d\.,]+)\s*([A-Za-z€$]+)?$', riga_pulita)
                                    if match_alt:
                                        desc = match_alt.group(1).strip()
                                        
                                        # Scarta se è una descrizione troppo corta o contiene parole chiave di intestazione
                                        if len(desc) < 3 or "fattura" in desc.lower() or "technova" in desc.lower():
                                            continue
                                            
                                        qta = "1"
                                        pz = match_alt.group(2).strip()
                                        val_pz = match_alt.group(3).strip()
                                        tot = match_alt.group(4).strip()
                                        val_tot = match_alt.group(5).strip() if match_alt.group(5) else val_pz
                                        
                                        righe_strutturate.append((
                                            file.filename, num_pagina, 
                                            qta, desc, 
                                            pz, val_pz, 
                                            tot, val_tot
                                        ))
                                            
            except Exception as e:
                print(f"Errore file {file.filename}: {str(e)}")
            
            if os.path.exists(percorso_pdf):
                os.remove(percorso_pdf)
                
    if not righe_strutturate:
        return "Nessun dato strutturato estratto dai file."

    # Creazione del Database Excel finale pulito e ordinato
    nome_file_xlsx = "database_contabile_definitivo.xlsx"
    percorso_xlsx = os.path.abspath(os.path.join(OUTPUT_FOLDER, nome_file_xlsx))
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Contabilità Analitica"
    
    # Intestazioni di colonna professionali
    ws.append([
        "Nome File PDF", "Pagina", "Quantità", 
        "Descrizione / Voce", "Prezzo Unitario", 
        "Valuta", "Totale Riga", "Valuta Totale"
    ])
    
    for row in righe_strutturate:
        filename, pagina, qta, desc, pz, val_pz, tot, val_tot = row
        
        # Conversione in formato numerico nativo per consentire calcoli e formule su Excel
        qta_num = int(qta) if qta.isdigit() else qta
        
        pz_num = pz
        if pz:
            try:
                pz_num = float(pz.replace('.', '').replace(',', '.'))
            except ValueError:
                pass
                
        tot_num = tot
        if tot:
            try:
                tot_num = float(tot.replace('.', '').replace(',', '.'))
            except ValueError:
                pass
                
        ws.append([filename, pagina, qta_num, desc, pz_num, val_pz, tot_num, val_tot])
        
    wb.save(percorso_xlsx)
    
    return send_file(percorso_xlsx, as_attachment=True, download_name=nome_file_xlsx)

if __name__ == '__main__':
    app.run(debug=True)