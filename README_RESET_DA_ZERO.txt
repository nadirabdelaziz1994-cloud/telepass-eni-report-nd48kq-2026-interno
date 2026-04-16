RESET DA ZERO - GITHUB
=====================

1) Crea un repository VUOTO su GitHub.

2) Carica TUTTO il contenuto di questa cartella nel repository.
   Se la cartella nascosta .github non si carica:
   - fai Add file > Create new file
   - nome file: .github/workflows/aggiorna-dashboard.yml
   - incolla il contenuto del file:
     WORKFLOW_DA_COPIARE_SE_NON_VEDI_CARTELLA_NASCOSTA.yml

3) In Settings > Pages imposta:
   - Source: Deploy from a branch
   - Branch: main
   - Folder: /docs

4) Metti i file qui:
   - input/lista -> 1 file Lista PDV
   - input/anagrafica -> 1 file Anagrafica
   - input/report_settimanali -> tutti i file settimana
   - input/gare -> il file Excel della gara aggiornata
   - FILE_UTILI -> pdf/excel utili da scaricare nel sito

5) Ogni volta che fai Commit, GitHub ricompila il sito da solo.

6) Il sito finale si apre dal link GitHub Pages.

NOTE IMPORTANTI
---------------
- Non toccare docs/index.html: si aggiorna da solo.
- Non rinominare template_dashboard.html
- Non rinominare aggiorna_dashboard.py
- Il layout di questa versione è quello approvato, con:
  * classifica senza colonne Wxx
  * gara letta da file input/gare
  * filtri attivi anche nella pagina gara
  * export Excel/PDF della sola gara quando sei nella pagina Gara PDV
