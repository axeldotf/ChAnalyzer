# ChAnalyzer

**ChAnalyzer** è un tool Python sviluppato in collaborazione con **Selektra Italia Srl** per l'analisi e la reportistica di copertura delle reti mobili. Il software elabora dati da file Excel per generare report dettagliati sulla copertura dei vari operatori (TIM, Vodafone, Wind3, Iliad) e tecnologie (GSM, UMTS, LTE).

## Funzionalità principali

- **Selezione e analisi automatica di file Excel** contenenti misurazioni di segnale mobile.
- **Mappatura automatica** degli operatori e delle tecnologie in base ai valori di ARFCN o Channel.
- **Calcolo delle metriche di copertura** per ogni operatore e tecnologia.
- **Formattazione condizionale dei dati** per evidenziare il livello di copertura con colori distintivi.
- **Generazione di report Excel** contenenti sia i dati analizzati per singolo piano, sia un riepilogo complessivo.

## Installazione

Per utilizzare il software, assicurarsi di avere Python installato e installare le seguenti librerie:

```sh
pip install pandas tqdm xlsxwriter
```

## Utilizzo

Eseguire lo script principale e seguire le istruzioni a schermo:

```sh
python script.py
```

### Passaggi interattivi:
1. Selezionare il file Excel di output.
2. Aggiungere uno o più file di input contenenti i dati di misurazione.
3. Specificare il nome del piano per ogni file selezionato.
4. Il software analizzerà i dati e creerà un report riepilogativo.

## Output

Alla fine del processo, il software genererà un file Excel contenente:
- **Un foglio per ogni piano selezionato**, con i dati elaborati.
- **Un foglio "Riepilogo"**, con una sintesi della copertura per ogni tecnologia e operatore.

## Autori

Questo software è stato sviluppato da **Alessandro Frullo** in collaborazione con **Selektra Italia Srl**.

