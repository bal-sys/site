import imaplib
import email
from email.header import decode_header
import os
import re
import csv
import getpass
from tqdm import tqdm
from dotenv import load_dotenv
from datetime import datetime, timedelta

# Charger les variables d'environnement du fichier .env
load_dotenv()

# Récupérer les informations de connexion depuis le fichier .env
username = os.getenv("EMAIL_USER")
imap_server = os.getenv("IMAP_SERVER")


# Fonction pour se connecter à la boîte mail
def connect_to_email(username, password, imap_server):
    try:
        imap = imaplib.IMAP4_SSL(imap_server)
        imap.login(username, password)
        imap.select("inbox")
        print("Connexion réussie")
        return imap
    except Exception as e:
        print(f"Erreur de connexion : {e}")
        return None


# Demander le mot de passe de manière sécurisée et essayer de se connecter
imap = None
while imap is None:
    password = getpass.getpass("Entrez votre mot de passe : ")
    imap = connect_to_email(username, password, imap_server)
    if imap is None:
        print("Mot de passe incorrect. Veuillez réessayer.")


# Fonction pour rechercher et extraire les emails non lus avec un objet spécifique
def fetch_emails_with_subject(imap, subject_keyword):
    try:
        status, messages = imap.search(
            None, f'(UNSEEN SUBJECT "{subject_keyword}")')
        email_ids = messages[0].split()
        emails = []
        for email_id in email_ids:
            status, msg_data = imap.fetch(email_id, "(RFC822)")
            msg = email.message_from_bytes(msg_data[0][1])
            subject, encoding = decode_header(msg["Subject"])[0]
            if isinstance(subject, bytes):
                subject = subject.decode(encoding if encoding else "utf-8")
            from_ = msg["From"]
            date_ = msg["Date"]
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        body = part.get_payload(decode=True).decode()
                        break
            else:
                body = msg.get_payload(decode=True).decode()
            emails.append((email_id, from_, subject, date_, body))
        return emails
    except Exception as e:
        print(f"Erreur lors de la récupération des emails : {e}")
        return []


# Fonction pour nettoyer le corps du message des balises HTML
def clean_html(raw_html):
    clean_re = re.compile('<.*?>')
    clean_text = re.sub(clean_re, '', raw_html)
    return clean_text


# Fonction pour extraire les informations du corps du message
def extract_data_from_body(body):
    data = {"Email": None, "Montant": None, "Durée": None, "Période": None}

    # Nettoyer le corps de l'email
    body = clean_html(body)

    # Modifiez les expressions régulières selon le format exact
    email_pattern = r"Email:\s*([\S]+)"
    montant_pattern = r"Montant:\s*(\d+(\.\d+)?)"
    duree_pattern = r"Durée:\s*(\d+)"
    periode_pattern = r"periode:\s*(Mois|Année)"

    email_match = re.search(email_pattern, body)
    montant_match = re.search(montant_pattern, body)
    duree_match = re.search(duree_pattern, body)
    periode_match = re.search(periode_pattern, body)

    data["Email"] = email_match.group(1).strip() if email_match else None
    data["Montant"] = montant_match.group(1).strip() if montant_match else None
    data["Durée"] = duree_match.group(1).strip() if duree_match else None
    data["Période"] = periode_match.group(1).strip() if periode_match else None

    return data


# Fonction pour sauvegarder les données dans un fichier CSV
def save_to_csv(file_path, data):
    with open(file_path, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(["Email", "Montant", "Durée", "Période", "Date"])
        for record in data:
            writer.writerow([
                record["Email"], record["Montant"], record["Durée"],
                record["Période"], record["Date"]
            ])


if imap:
    # Liste des sujets à traiter
    subjects = ["Nouvelle demande FR", "Nouvelle demande DE"]

    for subject_keyword in subjects:
        emails = fetch_emails_with_subject(imap, subject_keyword)

        # Dictionnaire pour stocker les enregistrements uniques avec le montant le plus élevé
        unique_records = {}

        # Utiliser tqdm pour afficher la barre de progression
        for email_id, from_, subject, date_, body in tqdm(
                emails,
                desc=f"Traitement des emails pour {subject_keyword}",
                unit="email"):
            extracted_data = extract_data_from_body(body)
            email_address = extracted_data["Email"]
            montant = float(
                extracted_data["Montant"]) if extracted_data["Montant"] else 0
            duree = extracted_data["Durée"]
            periode = extracted_data["Période"]

            # Convertir la date en format datetime
            try:
                email_date = datetime.strptime(date_,
                                               "%a, %d %b %Y %H:%M:%S %z")
            except ValueError:
                # Gérer les erreurs de format de date si nécessaire
                continue

            start_date = email_date + timedelta(days=90)  # Ajouter 3 mois
            formatted_date = start_date.strftime("%d/%m/%Y")

            # Ajouter la date au dictionnaire des données
            extracted_data["Date"] = formatted_date

            if email_address:
                if email_address in unique_records:
                    existing_record = unique_records[email_address]
                    existing_montant = float(
                        existing_record["Montant"]
                    ) if existing_record["Montant"] else 0
                    if montant > existing_montant:
                        unique_records[email_address] = extracted_data
                else:
                    unique_records[email_address] = extracted_data

        # Convertir les enregistrements uniques en liste pour sauvegarde
        records_list = list(unique_records.values())

        # Déterminer le chemin du fichier CSV en fonction du sujet
        if subject_keyword == "Nouvelle demande FR":
            csv_file_path = "Demande_FR.csv"
        elif subject_keyword == "Nouvelle demande DE":
            csv_file_path = "Demande_DE.csv"
        else:
            continue

        # Sauvegarder les enregistrements dans le fichier CSV
        save_to_csv(csv_file_path, records_list)

        # Afficher les enregistrements uniques ou un message d'absence
        if records_list:
            for data in records_list:
                print(f"Email: {data['Email']}")
                print(f"Montant: {data['Montant']}")
                print(f"Durée: {data['Durée']}")
                print(f"Période: {data['Période']}")
                print(f"Date: {data['Date']}")
                print("=" * 50)
            print(
                f"Traitement terminé avec succès pour le sujet : {subject_keyword}."
            )
        else:
            print(
                f"Aucune nouvelle demande pour le sujet : {subject_keyword}.")

    # Fermer la connexion
    imap.logout()
