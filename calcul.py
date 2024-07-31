import csv
from datetime import datetime, timedelta
import math

# Fonction pour lire les données depuis le fichier CSV
def read_from_csv(file_path):
    records = []
    with open(file_path, mode='r', newline='', encoding='utf-8') as file:
        reader = csv.DictReader(file)
        for row in reader:
            records.append(row)
    return records

# Fonction pour calculer le tableau d'amortissement
def calculate_amortization(principal, duration_months, annual_rate):
    monthly_rate = annual_rate / 12 / 100
    if monthly_rate == 0:
        monthly_payment = principal / duration_months
    else:
        monthly_payment = (principal * monthly_rate) / (1 - math.pow(1 + monthly_rate, -duration_months))
    total_payment = monthly_payment * duration_months
    return monthly_payment, total_payment

# Fonction pour sauvegarder les résultats dans un fichier CSV
def save_calculations_to_csv(file_path, calculations):
    with open(file_path, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(["Email", "Montant", "Durée", "Date de début", "Taux annuel", "Mensualité du crédit", "Total des mensualités"])
        for calc in calculations:
            writer.writerow([
                calc["Email"], calc["Montant"], calc["Durée"],
                calc["Date de début"], calc["Taux annuel"], calc["Mensualité du crédit"], calc["Total des mensualités"]
            ])

# Fonction principale pour traiter un fichier de demande
def process_file(input_file_path, output_file_path):
    # Lire les enregistrements depuis le fichier CSV d'entrée
    records = read_from_csv(input_file_path)

    # Taux d'intérêt annuel
    annual_interest_rate = 2.0

    # Liste pour les résultats de calcul
    results = []

    for record in records:
        try:
            email = record["Email"]
            montant = float(record["Montant"])
            duree = int(record["Durée"])
            periode = record["Période"]

            # Convertir la date de début du remboursement en format datetime
            request_date = datetime.strptime(record["Date"], "%d/%m/%Y")
            start_date = request_date + timedelta(days=90)  # Ajouter 3 mois
            start_date_formatted = start_date.strftime("%d/%m/%Y")

            # Calculer la durée en mois
            if periode == "Année":
                duration_months = duree * 12
            elif periode == "Mois":
                duration_months = duree
            else:
                raise ValueError(f"Période non valide : {periode}")

            # Calcul des mensualités
            monthly_payment, total_payment = calculate_amortization(montant, duration_months, annual_interest_rate)

            results.append({
                "Email": email,
                "Montant": montant,
                "Durée": f"{duree} {periode}",
                "Date de début": start_date_formatted,
                "Taux annuel": annual_interest_rate,
                "Mensualité du crédit": round(monthly_payment, 2),
                "Total des mensualités": round(total_payment, 2)
            })

        except ValueError as ve:
            print(f"Erreur de conversion des données : {ve}")
        except Exception as e:
            print(f"Erreur lors du traitement des enregistrements : {e}")

    # Sauvegarder les résultats dans le fichier CSV de sortie
    save_calculations_to_csv(output_file_path, results)

# Traiter les fichiers de demande
process_file("Demande_DE.csv", "calcul_DE.csv")
process_file("Demande_FR.csv", "calcul_FR.csv")

# Messages de confirmation
print("Calculs d'amortissement pour Demande_DE terminés avec succès.")
print("Calculs d'amortissement pour Demande_FR terminés avec succès.")
