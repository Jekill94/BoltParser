import streamlit as st
import pdfplumber
import openpyxl
import pandas as pd
import re
import os
import io

st.set_page_config(page_title="BoltParser - Excel Converter", page_icon="⚡", layout="centered")

st.title("⚡ BoltParser")
st.markdown("Estrai i dati dalle tue fatture PDF e convertili in un database Excel ordinato e pronto per i calcoli.")

# Filtro opzionale
parola_chiave = st.text_input("Filtra per parola chiave (facoltativo):", placeholder="Es. Totale, Importo, EUR...")

# Caricamento file multipli
uploaded_files = st.file_uploader("Seleziona uno o più file PDF", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    if st.button("Genera Database Globale Excel", type="primary"):
        righe_strutturate = []
        
        with st.spinner("Elaborazione dei PDF in corso..."):
            for file in uploaded_files:
                try:
                    # Legge il file direttamente dalla memoria di Streamlit
                    with pdfplumber.open(file) as pdf:
                        for indice, pagina in enumerate(pdf.pages):
                            num_pagina = indice + 1
                            testo_pagina = pagina.extract_text()
                            
                            if testo_pagina:
                                for riga in testo_pagina.split('\n'):
                                    riga_pulita = riga.strip()
                                    if not riga_pulita:
                                        continue
                                    
                                    riga_lower = riga_pulita.lower()
                                    
                                    # Filtri di scarto per intestazioni/piè di pagina
                                    if any(parola in riga_lower for parola in ["pagina", "elettronica", "descrizione articolo", "p.iva", "codice fiscale", "spett.le", "fattura n"]):
                                        continue
                                        
                                    if "test" in riga_lower and not any(ch.isdigit() for ch in riga_lower):
                                        continue

                                    if parola_chiave and parola_chiave.lower() not in riga_lower:
                                        continue

                                    # Pattern riga con quantità
                                    match = re.search(r'^(.*?)\s+(\d+)\s+([\d\.,]+)\s*([A-Za-z€$]+)\s+([\d\.,]+)\s*([A-Za-z€$]+)?$', riga_pulita)
                                    
                                    if match:
                                        desc = match.group(1).strip()
                                        qta = match.group(2).strip()
                                        pz = match.group(3).strip()
                                        val_pz = match.group(4).strip()
                                        tot = match.group(5).strip()
                                        val_tot = match.group(6).strip() if match.group(6) else val_pz
                                        
                                        if "fattura" in desc.lower() or len(desc) < 2:
                                            continue
                                            
                                        righe_strutturate.append((file.name, num_pagina, qta, desc, pz, val_pz, tot, val_tot))
                                    else:
                                        # Pattern riga senza quantità esplicita (default = 1)
                                        match_alt = re.search(r'^(.*?)\s+([\d\.,]+)\s*([A-Za-z€$]+)\s+([\d\.,]+)\s*([A-Za-z€$]+)?$', riga_pulita)
                                        if match_alt:
                                            desc = match_alt.group(1).strip()
                                            if len(desc) < 3 or "fattura" in desc.lower() or "technova" in desc.lower():
                                                continue
                                                
                                            qta = "1"
                                            pz = match_alt.group(2).strip()
                                            val_pz = match_alt.group(3).strip()
                                            tot = match_alt.group(4).strip()
                                            val_tot = match_alt.group(5).strip() if match.group(5) else val_pz
                                            
                                            righe_strutturate.append((file.name, num_pagina, qta, desc, pz, val_pz, tot, val_tot))
                                                
                except Exception as e:
                    st.error(f"Errore nel file {file.name}: {str(e)}")
                    
        if righe_strutturate:
            # Creazione del dataframe e conversione numerica
            df_data = []
            for row in righe_strutturate:
                filename, pagina, qta, desc, pz, val_pz, tot, val_tot = row
                
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
                        
                df_data.append([filename, pagina, qta_num, desc, pz_num, val_pz, tot_num, val_tot])
                
            df = pd.DataFrame(df_data, columns=[
                "Nome File PDF", "Pagina", "Quantità", 
                "Descrizione / Voce", "Prezzo Unitario", 
                "Valuta", "Totale Riga", "Valuta Totale"
            ])
            
            st.success(f"Estrazione completata con successo! Trovate {len(df)} voci contabili.")
            
            # Mostra un'anteprima interattiva nella pagina
            st.dataframe(df)
            
            # Esportazione in Excel in memoria
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Contabilità Analitica')
            excel_data = output.getvalue()
            
            # Pulsante di download diretto
            st.download_button(
                label="📥 Scarica Database Excel (.xlsx)",
                data=excel_data,
                file_name="database_contabile_definitivo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Nessun dato strutturato trovato nei file selezionati.")
