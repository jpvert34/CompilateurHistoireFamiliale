# Projet de Traitement des Données de Décès

Ce projet consiste en un script Python qui analyse et traite des données issues de fichiers CSV contenant des informations sur des décès. Le script extrait des informations spécifiques, transforme les données, et génère un fichier Excel formaté pour chaque nom de famille spécifié dans la liste de recherche.

## Fonctionnalités

- Lecture de fichiers CSV depuis un répertoire spécifié.
- Filtrage des données basées sur des noms de famille pré-définis.
- Transformation des données incluant la conversion des dates de naissance et de décès.
- Génération de fichiers Excel avec des données nettoyées et formatées.
- Log des événements pour un suivi facile des opérations et des erreurs.

## Prérequis

Pour exécuter ce script, vous aurez besoin de Python 3.6 ou plus récent. Les bibliothèques suivantes doivent être installées :

- pandas
- openpyxl
- xlsxwriter
- csv
- logging

## Installation

1. Clonez le dépôt sur votre machine locale en utilisant la commande suivante :
   ```bash
   git clone https://github.com/jpvert34/CompilateurHistoireFamiliale.git
2. Installez les dépendances nécessaires :
   ```bash
    pip install pandas openpyxl xlsxwriter
3. Utilisation
   Pour lancer le script, naviguez dans le répertoire du projet et exécutez :
   ```bash
   python script_deces.py
Assurez-vous que le répertoire contenant les fichiers CSV est correctement spécifié dans le script sous directory_path.

4. Contribution
Les contributions à ce projet sont bienvenues. Si vous souhaitez contribuer, veuillez forker le dépôt, créer une nouvelle branche pour vos modifications, et soumettre une pull request pour examen.
