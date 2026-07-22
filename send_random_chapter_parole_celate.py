import os
import re
import ssl
import random
import smtplib
import requests
from pathlib import Path
from email.message import EmailMessage
from zoneinfo import ZoneInfo
from datetime import datetime

from docx import Document


# Regex per riconoscere una riga che contiene SOLO un numero arabo
ARABIC_NUM_RE = re.compile(r"^\d+$")


# Nome del file DOCX.
DOCX_PATH = Path(os.getenv("DOCX_PATH", "parole_celate.docx"))

# File in cui vengono salvati i capitoli già inviati
HISTORY_PATH = Path(os.getenv("HISTORY_PATH", "cronologia_parole_celate.txt"))

# Numero atteso di capitoli nel documento (71 + 82 = 153)
EXPECTED_CHAPTER_COUNT = int(os.getenv("EXPECTED_CHAPTER_COUNT", "153"))

# Configurazione Gmail SMTP
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 465

# GitHub API
GITHUB_OWNER = "silverfred"
GITHUB_REPO = "spigolature-email"
GITHUB_WORKFLOW = "daily_email_parole_celate.yml"


def should_run_now() -> bool:
    """
    Su GitHub Actions:
    - se il workflow viene lanciato manualmente, invia sempre;
    - se parte da cron, invia solo alle ore locali previste.

    Questo serve perché GitHub Actions usa UTC,
    mentre noi vogliamo inviare alle 6:00 e alle 18:00 ora italiana.
    """
    if os.getenv("GITHUB_EVENT_NAME") == "workflow_dispatch":
        return True

    enforce = os.getenv("ENFORCE_LOCAL_HOURS", "false").lower() == "true"

    if not enforce:
        return True

    timezone = os.getenv("LOCAL_TZ", "Europe/Rome")
    allowed_hours_raw = os.getenv("ALLOWED_LOCAL_HOURS", "6,18")

    allowed_hours = {
        int(hour.strip())
        for hour in allowed_hours_raw.split(",")
        if hour.strip()
    }

    now = datetime.now(ZoneInfo(timezone))

    return now.hour in allowed_hours


def clean_paragraph_text(paragraph) -> str:
    """
    Estrae il testo da un paragrafo del DOCX.

    Gestisce:
    - paragrafi vuoti;
    - righe vuote;
    - paragrafi con soli spazi.
    """
    text = paragraph.text.strip()
    return text


def load_chapters(docx_path: Path) -> list[dict]:
    """
    Carica il DOCX e estrae i 153 capitoli numerati (1-71 nella prima parte, 1-82 nella seconda).
    Salta il primo paragrafo (titolo) e l'ultimo (corsivo).
    """
    if not docx_path.exists():
        raise FileNotFoundError(f"File DOCX non trovato: {docx_path}")

    doc = Document(str(docx_path))

    # Estrai tutti i paragrafi
    all_paragraphs = doc.paragraphs

    if not all_paragraphs:
        raise ValueError("Il documento non contiene paragrafi.")

    # Salta il primo paragrafo (titolo iniziale)
    paragraphs = all_paragraphs[1:-1]  # Esclude primo e ultimo

    chapter_starts = []

    # Trova tutti gli indici dove inizia un numero arabo
    for para_index, para in enumerate(paragraphs):
        text = clean_paragraph_text(para).strip()
        normalized_text = text.strip()

        if ARABIC_NUM_RE.fullmatch(normalized_text):
            chapter_starts.append((normalized_text, para_index))

    if len(chapter_starts) != EXPECTED_CHAPTER_COUNT:
        raise ValueError(
            f"Numero capitoli inatteso: trovati {len(chapter_starts)}, "
            f"attesi {EXPECTED_CHAPTER_COUNT}."
        )

    chapters = []

    for i, (chapter_num, start_index) in enumerate(chapter_starts):
        if i + 1 < len(chapter_starts):
            end_index = chapter_starts[i + 1][1]  # FIX: Prendi l'indice [1], non il numero [0]
        else:
            end_index = len(paragraphs)

        chapter_parts = []

        # Raccogli i paragrafi dal numero fino al prossimo numero
        for para in paragraphs[start_index + 1:end_index]:
            text = clean_paragraph_text(para)

            if text:
                chapter_parts.append(text)

        chapter_text = "\n\n".join(chapter_parts).strip()

        if chapter_text:
            chapters.append(
                {
                    "number": chapter_num,
                    "text": chapter_text,
                }
            )

    return chapters


def load_history(history_path: Path) -> set[str]:
    """
    Legge cronologia_parole_celate.txt e restituisce l'insieme dei capitoli già inviati.
    Se il file non esiste, restituisce un insieme vuoto.
    """
    if not history_path.exists():
        return set()

    sent_chapters = set()

    with history_path.open("r", encoding="utf-8") as file:
        for line in file:
            value = line.strip()

            if value:
                sent_chapters.add(value)

    return sent_chapters


def append_to_history(history_path: Path, chapter_number: str) -> None:
    """
    Aggiunge a cronologia_parole_celate.txt il numero del capitolo appena inviato.
    """
    with history_path.open("a", encoding="utf-8") as file:
        file.write(chapter_number.strip() + "\n")


def choose_random_chapter(chapters: list[dict], history_path: Path) -> dict:
    """
    Sceglie un capitolo random tra quelli non ancora inviati.

    Quando tutti i capitoli sono stati inviati,
    svuota cronologia_parole_celate.txt e ricomincia da capo.
    """
    sent_chapters = load_history(history_path)

    available_chapters = [
        chapter
        for chapter in chapters
        if chapter["number"] not in sent_chapters
    ]

    if not available_chapters:
        print("Tutti i capitoli sono già stati inviati. Resetto la cronologia.")
        history_path.write_text("", encoding="utf-8")
        available_chapters = chapters

    return random.choice(available_chapters)


def get_required_env(name: str) -> str:
    """
    Legge una variabile d'ambiente obbligatoria.
    Se manca, blocca lo script con un errore chiaro.
    """
    value = os.getenv(name)

    if not value:
        raise EnvironmentError(f"Variabile d'ambiente mancante: {name}")

    return value


def generate_trigger_link() -> str:
    """
    Genera un link che trigghera il workflow manualmente.
    
    Il link usa l'API di GitHub per dispatchare il workflow.
    Quando cliccato, trigghera il workflow senza feedback visibile.
    """
    workflow_token = get_required_env("WORKFLOW_TOKEN")
    
    # URL base per triggerare il workflow via API
    base_url = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/actions/workflows/{GITHUB_WORKFLOW}/dispatches"
    
    # Aggiungiamo il token come parametro (workaround per link cliccabile)
    # In realtà useremo una soluzione più elegante con curl nel link
    link = f"{base_url}?token={workflow_token}"
    
    return link


def send_email(subject: str, body: str, trigger_link: str = "") -> None:
    """
    Invia l'email usando Gmail SMTP SSL.
    Richiede queste variabili d'ambiente:

    - SENDER_EMAIL
    - APP_PASSWORD
    - RECEIVER_EMAIL
    """
    sender_email = get_required_env("SENDER_EMAIL")
    app_password = get_required_env("APP_PASSWORD").replace(" ", "")
    receiver_email = get_required_env("RECEIVER_EMAIL")

    # Se c'è il trigger link, aggiungilo al body
    if trigger_link:
        body += f"\n\n---\n🔄 Clicca qui per un nuovo estratto:\n{trigger_link}"

    message = EmailMessage()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject
    message.set_content(body, subtype="plain", charset="utf-8")

    context = ssl.create_default_context()

    with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, context=context) as server:
        server.login(sender_email, app_password)
        server.send_message(message)


def main() -> None:
    """
    Funzione principale:
    - controlla se deve inviare ora;
    - carica i capitoli dal DOCX;
    - sceglie un capitolo non ancora inviato;
    - genera il link per triggerare un nuovo capitolo;
    - manda l'email;
    - aggiorna cronologia_parole_celate.txt.
    """
    if not should_run_now():
        print("Ora locale non prevista per l'invio. Esco senza inviare.")
        return

    chapters = load_chapters(DOCX_PATH)

    selected_chapter = choose_random_chapter(chapters, HISTORY_PATH)

    chapter_number = selected_chapter["number"]
    chapter_text = selected_chapter["text"]

    subject = f"Parole Celate — Capitolo {chapter_number}"

    body = (
        f"Capitolo {chapter_number}\n\n"
        f"{chapter_text}\n\n"
        "---\n"
        "Invio automatico."
    )

    # Genera il link per triggerare un nuovo capitolo
    try:
        trigger_link = generate_trigger_link()
    except EnvironmentError:
        print("WORKFLOW_TOKEN non trovato, email senza link.")
        trigger_link = ""

    send_email(subject, body, trigger_link)

    append_to_history(HISTORY_PATH, chapter_number)

    print(f"Inviato capitolo {chapter_number}.")


if __name__ == "__main__":
    main()
