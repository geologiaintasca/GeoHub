# SismaSite

Applicazione web per l'analisi sismica preliminare di un sito secondo le Norme Tecniche per le Costruzioni 2018 (NTC 2018): calcola i parametri di pericolosità sismica di base interpolandoli dal reticolo di riferimento ufficiale a 10.751 nodi, genera gli spettri di risposta e supporta lo screening preliminare di pendii e liquefazione.

🔗 **Demo online:** https://geologiaintasca.github.io/GeoHub/

## Prerequisiti

Per usare l'applicazione:
- Un browser moderno (Chrome, Firefox, Edge, Safari)
- Un server web locale per lo sviluppo, ad es. Python 3 (già incluso in macOS/Linux, scaricabile da [python.org](https://python.org) su Windows) — necessario perché i browser bloccano il caricamento di `grid-data.json` se la pagina viene aperta come file locale

Per rigenerare `grid-data.json` da una nuova fonte dati (facoltativo):
- Python 3.10+
- Librerie `pandas`, `xlrd` (per file `.xls`) e/o `openpyxl` (per file `.xlsx`)

## Installazione

```bash
git clone https://github.com/geologiaintasca/GeoHub.git
cd GeoHub
```

Nessuna build o installazione di dipendenze è necessaria per usare l'app: è HTML/CSS/JavaScript puro, eseguito interamente nel browser.

## Guida all'uso

Avviare un server locale nella cartella del progetto:

```bash
python3 -m http.server 8000
```

e aprire [http://localhost:8000/](http://localhost:8000/) nel browser.

Per rigenerare `grid-data.json` a partire da un file ufficiale del reticolo di riferimento (Excel o TXT/CSV):

```bash
pip install pandas xlrd openpyxl
python3 convert_grid.py spettri2008.xls grid-data.json
```

## Tecnologie utilizzate

- HTML5 / CSS3 / JavaScript (vanilla, nessun framework)
- [Leaflet.js](https://leafletjs.com/) — mappa interattiva per la selezione del sito
- [Chart.js](https://www.chartjs.org/) — spettri di risposta e grafico di disaggregazione
- Python 3 (`pandas`, `xlrd`/`openpyxl`) — script di conversione dati (`convert_grid.py`)

## 📄 Licenza

Questo progetto è distribuito sotto licenza **MIT**. Consulta il file [LICENSE](LICENSE) per i dettagli completi.

> **Nota sui dati:** i dati del reticolo di riferimento sismico e i parametri NTC 2018 sono elaborati a partire da fonti ufficiali INGV / Ministero delle Infrastrutture e dei Trasporti.
