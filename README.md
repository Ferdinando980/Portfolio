# Portfolio di cybersecurity / Cybersecurity portfolio

Sito statico con i case study del Google Cybersecurity Certificate, in italiano e in inglese.
A static site with the case studies from the Google Cybersecurity Certificate, in Italian and English.

## Cosa contiene / What it contains

- 20 case study, ciascuno con la pagina web e il PDF scaricabile (tranne l'inventario dei beni, che contiene nomi resi anonimi).
- Profilo e pagina Progetti.
- Tutti gli esercizi sono su scenari simulati del corso. Nessun sistema reale è stato toccato.

## Come si genera / How to build

Servono Python con i pacchetti `python-docx`, `lxml` e `pillow`.

```bash
python tools/build_site.py
```

Il sito pronto si trova nella cartella `docs/`, senza server e senza risorse esterne.
L'anteprima in locale si apre con `python -m http.server 8765` dentro `docs/`.

## Struttura / Layout

| Cartella | Contenuto |
|----------|-----------|
| `docs/` | Il sito generato, da pubblicare (GitHub Pages dalla cartella `docs`). |
| `tools/` | Il generatore e i testi: `profile.json`, `projects.json`, `site.config.json`, `cases_data.py`. |
| `tools/assets/` | Stile e grafica. |

I testi dei case study non sono in questo repository: il generatore li legge dai file Word della cartella `Security Audit/01_Case-Studies`.

## Privacy

Nel sito non ci sono telefono, indirizzo, data di nascita né email. La foto originale (`tools/photo/`) è esclusa da `.gitignore`.
