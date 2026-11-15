# ⚡ BoltParser

**BoltParser** è un'applicazione web leggera e intelligente progettata per semplificare la contabilità aziendale. Converte file PDF multipli (come fatture e preventivi) in un unico database Excel strutturato e pronto per i calcoli.

---

## 🎯 A cosa serve?
Gestire fatture in formato PDF può richiedere molto tempo se fatto a mano. BoltParser automatizza il lavoro sporco: legge i documenti, filtra le informazioni inutili, separa i testi dai valori numerici e restituisce un foglio di calcolo pulito e ordinato.

---

## ✨ Funzionalità principali
* **Caricamento Multiplo:** Seleziona e carica intere serie di file PDF in un solo colpo.
* **Separazione Chirurgica dei Dati:** Divide automaticamente la **Quantità**, la **Descrizione**, il **Prezzo Unitario** e il **Totale Riga** in colonne distinte.
* **Filtro Intelligente:** Scarta autonomamente intestazioni, piè di pagina e scorie testuali (come "Pagina 1 di 1" o dati anagrafici superflui).
* **Pronto per i Calcoli:** Converte i dati numerici nel formato nativo di Excel, permettendo l'uso immediato di formule e somme automatiche (es. `=SOMMA(...)`).
* **Filtro per Parola Chiave:** Cerca e isola rapidamente solo le voci di spesa desiderate.

---

## 🛠️ Tecnologie utilizzate
* [Python](https://www.python.org/) - Linguaggio di programmazione principale.
* [Streamlit](https://streamlit.io/) - Framework per l'interfaccia grafica web interattiva.
* [pdfplumber](https://github.com/jsvine/pdfplumber) - Libreria avanzata per l'estrazione geometrica e testuale dai PDF.
* [openpyxl](https://openpyxl.readthedocs.io/) & [Pandas](https://pandas.pydata.org/) - Gestione ed esportazione dei file tabulari in formato `.xlsx`.

---

## 🚀 Come avviarlo in locale
Se desideri eseguire l'applicazione direttamente sul tuo computer:

1. Clona o scarica questo repository.
2. Installa le dipendenze richieste aprendo il terminale nella cartella del progetto:
   ```bash
   pip install -r requirements.txt
