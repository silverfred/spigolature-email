# Spigolature

Questa repository invia automaticamente un capitolo random di `spigolature.docx` via email.

## Orari di invio

Il workflow è configurato per inviare al massimo 2 email al giorno, in ora italiana:

- una tra le 05:00 e le 06:00;
- una tra le 18:00 e le 19:00.

GitHub Actions schedula i cron in UTC. Per gestire correttamente sia ora solare sia ora legale italiana, `cron-trigger.yml` lancia alcuni tentativi nelle finestre UTC compatibili; lo script Python controlla poi l'ora locale `Europe/Rome` e impedisce doppioni tramite `invii_email.txt`.

## File principali

- `.github/workflows/cron-trigger.yml`: attiva il workflow negli orari programmati.
- `.github/workflows/daily_email.yml`: installa le dipendenze, esegue lo script e salva le cronologie.
- `send_random_chapter.py`: estrae un capitolo random non ancora inviato e manda la mail.
- `cronologia.txt`: tiene traccia dei capitoli già inviati, così non si ripetono finché non sono finiti.
- `invii_email.txt`: tiene traccia delle finestre giornaliere già usate, così non partono più email nella stessa fascia.
- `requirements.txt`: dipendenze Python.

## Test manuale

Da GitHub puoi usare **Actions → Random chapter email → Run workflow**.

Il test manuale invia una mail anche fuori dalle fasce orarie, perché serve a verificare rapidamente che secrets, SMTP e script funzionino.
