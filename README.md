# Projet-Stage-Parisot-Industrie

# Importateur de codes vers SQL Server

Cette application Python permet d’importer, visualiser, modifier et exporter des codes à 13 chiffres dans une base de données SQL Server via une interface graphique intuitive.

---

## Fonctionnalités

- Import de fichiers texte (.txt) par dialogue ou glisser-déposer
- Validation automatique des codes (longueur, numérique, doublons)
- Classement automatique dans deux tables SQL selon le préfixe
- Affichage paginé (50 codes par page) avec navigation
- Modification et suppression de codes directement dans l’interface
- Export des codes vers un fichier texte
- Barre de progression et journal des erreurs

---

## Prérequis

- Python 3.x
- SQL Server (ex: SQLEXPRESS)
- ODBC Driver 17 for SQL Server
- Modules Python :

```bash
pip install pyodbc tkinterdnd2

