import locale
import logging
import multiprocessing
import os
from datetime import datetime

import pandas as pd

import csv

# Configuration du logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Chemin du répertoire contenant les fichiers CSV de décès
directory_path = './deces'

# Vérification de l'existence du répertoire
if not os.path.exists(directory_path):
    logging.error(f"Le répertoire spécifié {directory_path} n'existe pas.")
    exit(1)
# Noms de famille à rechercher (codés en dur pour l'exemple)
family_names = ["Detronde", "Marchand", "Vert", "Roturier", "Grandi"]

# Chemins des fichiers de référence pour les cantons et les communes
canton_path = 'csv/canton_2022.csv'
communes_path = 'csv/communes1943_2022.csv'

# Lecture des fichiers de référence pour les cantons et les communes
canton_df = pd.read_csv(canton_path, delimiter=',', dtype=str)
communes_df = pd.read_csv(communes_path, delimiter=',', dtype=str)

# Création de dictionnaires pour la correspondance des codes avec les noms
canton_dict = dict(zip(canton_df['BURCENTRAL'].astype(str), canton_df['NCC']))
communes_dict = dict(zip(communes_df['COM'].astype(str), communes_df['NCC']))

# Définir la locale en français
locale.setlocale(locale.LC_TIME, 'fr_FR')  # Sur certains systèmes, utilisez 'fr_FR.utf8'


# Fonction pour lire un fichier CSV en détectant le délimiteur
def read_csv_with_delimiter(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            sniffer = csv.Sniffer()
            delimiter = str(sniffer.sniff(file.readline()).delimiter)
            file.seek(0)  # Reset file read position
            return pd.read_csv(file_path, delimiter=delimiter, low_memory=False)
    except Exception as e:
        print(f"Erreur lors de la lecture du fichier {file_path}: {e}")
        return None


def safe_parse_date(date_str, format_date='%Y%m%d'):
    try:
        return datetime.strptime(date_str, format_date)
    except ValueError:
        return None


# Fonction pour transformer une ligne de données
def transform_row(row):
    transformed = {}
    try:
        # Diviser nomprenom en Nom et Prénom
        parts = row['nomprenom'].split('*')
        nom = parts[0].strip() if len(parts) > 0 else ''  # Utilisez strip() pour supprimer les espaces inutiles
        prenom = parts[1].replace('/', '') if len(parts) > 1 else ''

        # Convertir sexe
        sexe = 'H' if row['sexe'] == 1 else 'F'

        # Convertir datenaiss et datedeces en format de date lisible
        date_naiss = safe_parse_date(str(row['datenaiss']))
        date_naiss_str = date_naiss.strftime('%a. %d %b. %Y') if date_naiss is not None else ''

        date_deces = safe_parse_date(str(row['datedeces']))
        date_deces_str = date_deces.strftime('%a. %d %b. %Y') if date_deces is not None else ''

        # Calculer l'âge uniquement si les deux dates sont valides
        age = ''
        if date_naiss is not None and date_deces is not None:
            age = int((date_deces - date_naiss).days / 365.25)

        # Extraire le département depuis actedeces
        dep = str(row['lieudeces'])[:2]  # Les 2 premiers chiffres pour le département

        # Essayer de trouver le lieu de décès dans le dictionnaire des cantons
        lieu_deces = canton_dict.get(str(row['lieudeces']), '')

        # Si pas trouvé dans les cantons, chercher dans le dictionnaire des communes
        if not lieu_deces:
            lieu_deces = communes_dict.get(str(row['lieudeces']), '')

        transformed = {'Nom': nom, 'Prénom': prenom, 'Sexe': sexe, 'Date Naiss': date_naiss_str,
                       'Lieu Naiss': row['lieunaiss'], 'Comm Naiss': row['commnaiss'], 'Pays Naiss': row['paysnaiss'],
                       'Age': age, 'Date Deces': date_deces_str, 'Dep': dep, 'Lieu Deces': lieu_deces}

        return pd.Series(transformed)
    except Exception as e:
        print(f"Erreur : {e}")
        # En cas d'erreur, retournez une série de None pour chaque colonne transformée
        return pd.Series([None] * len(transformed))


# Fonction pour créer un fichier Excel formaté
def create_excel_file(data, name):
    # Nettoyage des données: Remplacer NaN par une chaîne vide ou une autre valeur par défaut
    data = data.fillna('')

    # Trier les données par date de décès
    data['Date Deces'] = pd.to_datetime(data['Date Deces'], format='%a. %d %b. %Y', errors='coerce')
    data = data.sort_values(by='Date Deces', ascending=True)
    # Formater la date au format souhaité (ven.. 06 août. 2021)
    data['Date Deces'] = data['Date Deces'].dt.strftime('%a. %d %b. %Y')

    # Éliminer les doublons basés sur le nom complet (Nom + Prénom)
    data = data.drop_duplicates(subset=['Nom', 'Prénom'], keep='first')

    # Chemin du fichier Excel
    excel_path = f'xlsx/noms {name}.xlsx'

    # Utilisation de 'with' pour la gestion automatique des ressources
    with pd.ExcelWriter(excel_path, engine='xlsxwriter',
                        engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
        # Votre code pour écrire dans le fichier Excel

        data.to_excel(writer, index=False, sheet_name='Sheet1')

        # Mise en forme avec xlsxwriter
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        # Configurations supplémentaires pour l'orientation en paysage
        worksheet.set_landscape()
        worksheet.fit_to_pages(1, 1)  # Ajuster la mise en page pour s'adapter à une page en largeur et en hauteur
        header_format = workbook.add_format(
            {'bg_color': '#780373', 'bold': True, 'font_color': 'white', 'align': 'center',
             # Centre le texte horizontalement
             'valign': 'vcenter',  # Centre le texte verticalement
             'border': 1  # Ajout de bordures
             })
        cell_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
        grey_format = workbook.add_format({'bg_color': '#D3D3D3', 'align': 'center', 'valign': 'vcenter', 'border': 1})

        # Mise en forme des en-têtes
        for col_num, value in enumerate(data.columns.values):
            worksheet.write(0, col_num, value, header_format)

        # Ajustement de la taille des colonnes et mise en forme des cellules
        for col_num, column in enumerate(data.columns):
            column_length = max(data[column].astype(str).map(len).max(), len(column)) + 2
            worksheet.set_column(col_num, col_num, column_length, cell_format)

        # Griser une ligne sur deux et centrer le texte
        for row_num in range(1, len(data) + 1):
            format_to_use = grey_format if row_num % 2 == 0 else cell_format
            for col_num in range(len(data.columns)):
                worksheet.write(row_num, col_num, data.iloc[row_num - 1, col_num], format_to_use)


# Traitement principal
def process_name(family_name):
    logging.info(f"Début du traitement pour le nom : {family_name}")

    all_data = []  # Utilisez une liste pour stocker les DataFrames individuels

    # Lecture de chaque fichier CSV dans le répertoire
    for file in os.listdir(directory_path):
        if file.endswith(".csv"):
            file_path = os.path.join(directory_path, file)

            # Vérification de l'existence du fichier
            if not os.path.exists(file_path):
                logging.error(f"Le fichier spécifié {file_path} n'existe pas.")
                continue

            df = read_csv_with_delimiter(file_path)

            if df is not None:
                # Séparation du nom et du prénom, puis comparaison stricte du nom
                filtered_data = df[df['nomprenom'].str.split('*').str[0].str.strip().str.lower() == family_name.lower()]

                # Transformation des données
                transformed_data = filtered_data.apply(transform_row, axis=1, result_type='expand')

                # Ajout des données transformées à la liste
                all_data.append(transformed_data)

    # Concaténation de tous les DataFrames de la liste en un seul DataFrame
    if all_data:
        # Avant de concaténer, filtrer les DataFrames vides ou tout-NA
        all_data = [df for df in all_data if not df.empty and not df.isna().all().all()]
        all_data = pd.concat(all_data, ignore_index=True) if all_data else pd.DataFrame()

        columns_to_keep = ['Nom', 'Prénom', 'Sexe', 'Date Naiss', 'Lieu Naiss', 'Comm Naiss', 'Pays Naiss', 'Age',
                           'Date Deces', 'Dep', 'Lieu Deces']

        all_data = all_data[columns_to_keep]

        # Création du fichier Excel si des données sont trouvées
        if not all_data.empty:
            create_excel_file(all_data, family_name)
            logging.info(f"Fichier Excel créé pour : {family_name}")
        else:
            logging.info(f"Aucune donnée trouvée pour : {family_name}")

    logging.info(f"Fin du traitement pour le nom : {family_name}")


if __name__ == "__main__":
    with multiprocessing.Pool(processes=4) as pool:  # Vous pouvez ajuster le nombre de processus selon votre système
        pool.map(process_name, family_names)
