import smtplib
import csv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header
from dotenv import load_dotenv
import os
from tqdm import tqdm
import getpass

# Charger les variables d'environnement du fichier .env
load_dotenv()

# Récupérer les informations de connexion depuis le fichier .env
smtp_server = os.getenv("SMTP_SERVER")
smtp_port = int(os.getenv("SMTP_PORT"))
smtp_user = os.getenv("SMTP_USER")


# Fonction pour lire les données depuis le fichier CSV
def read_from_csv(file_path):
    records = []
    with open(file_path, mode='r', newline='', encoding='utf-8') as file:
        reader = csv.DictReader(file)
        for row in reader:
            records.append(row)
    return records


# Fonction pour envoyer un email
def send_email(to_address, subject, body, server):
    try:
        # Créer le message
        msg = MIMEMultipart()
        msg['From'] = smtp_user
        msg['To'] = to_address
        msg['Subject'] = Header(subject, 'utf-8')

        # Ajouter le corps du message
        msg.attach(MIMEText(body, 'html', 'utf-8'))

        # Envoi du message
        server.sendmail(smtp_user, to_address, msg.as_string())
    except Exception as e:
        print(f"Erreur lors de l'envoi de l'email à {to_address} : {e}")


# Lire les enregistrements depuis les fichiers CSV
files = ["calcul_DE.csv", "calcul_FR.csv"]
subjects = {
    "calcul_DE.csv": "Angebot für Ihr Darlehen",
    "calcul_FR.csv": "Proposition de prêt"
}
bodies = {
    "calcul_DE.csv":
    """
    <html>
    <body>
        <p>Sehr geehrte/r Kunde/in,</p>
        <p>Hier sind die Details Ihrer Anfrage:</p>
        <ul>
            <li><strong>Betrag:</strong> {Montant} EUR</li>
            <li><strong>Dauer:</strong> {Durée}</li>
            <li><strong>Jahreszins:</strong> {Taux_annuel}%</li>
            <li><strong>Beginn des Kredits:</strong> {Date_de_début}</li>
            <li><strong>Monatliche Rate:</strong> {Mensualité_du_crédit} EUR</li>
            <li><strong>Gesamtzahl der Raten:</strong> {Total_des_mensualités} EUR</li>
        </ul>
        <p>Mit freundlichen Grüßen,</p>
        <p>Ihr Team</p>
    </body>
    </html>
    """,
    "calcul_FR.csv":
    """
    <html>
    <body>
        <p>Bonjour,</p>
        <p>Voici les détails de votre demande :</p>
        <ul>
            <li><strong>Montant:</strong> {Montant} EUR</li>
            <li><strong>Durée:</strong> {Durée}</li>
            <li><strong>Taux annuel:</strong> {Taux_annuel}%</li>
            <li><strong>Date de début:</strong> {Date_de_début}</li>
            <li><strong>Mensualité du crédit:</strong> {Mensualité_du_crédit} EUR</li>
            <li><strong>Total des mensualités:</strong> {Total_des_mensualités} EUR</li>
        </ul>
        <p>Cordialement,</p>
        <p>Votre équipe</p>
    </body>
    </html>
    """
}

# Demander le mot de passe de manière sécurisée
smtp_password = getpass.getpass("Entrez votre mot de passe : ")

# Connexion au serveur SMTP
try:
    with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
        server.login(smtp_user, smtp_password)

        # Envoyer des emails pour chaque fichier
        for file in files:
            records = read_from_csv(file)
            subject = subjects[file]
            body_template = bodies[file]

            # Afficher la barre de progression
            for record in tqdm(records,
                               desc=f"Envoi des emails pour {file}",
                               unit="email"):
                email_address = record.get("Email", "")
                body = body_template.format(
                    Montant=record.get("Montant", ""),
                    Durée=record.get("Durée", ""),
                    Taux_annuel=record.get("Taux_annuel", ""),
                    Date_de_début=record.get("Date_de_début", ""),
                    Mensualité_du_crédit=record.get("Mensualité_du_crédit",
                                                    ""),
                    Total_des_mensualités=record.get("Total_des_mensualités",
                                                     ""))
                send_email(email_address, subject, body, server)

    print("Tous les emails ont été envoyés.")
except Exception as e:
    print(f"Erreur lors de la connexion au serveur  : {e}")
