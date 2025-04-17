
# Description

Ce script Python permet de modifier l’auteur des métadonnées des fichiers Microsoft Office (.docx, .xlsx, .pptx) présents dans un répertoire donné. Il utilise COM (Component Object Model) pour interagir avec les fichiers via `win32com.client`.

# Fonctionnalités 

- Parcours d’un dossier spécifié.
- Vérification du type de fichier (Word, Excel, PowerPoint).
- Modification de la propriété intégrée "Author".
- Enregistrement et fermeture automatique des fichiers après modification.
- Gestion des erreurs en cas d’échec du traitement d’un fichier.

# Prérequis

Avant d'exécuter ce script, assurez-vous d'avoir installé **Python 3.x** et les dépendances listées dans `requirements.txt`.

> [!TIP]
> Installation des dépendances : `pip install -r requirements.txt`

# Utilisation

1. **Modifier le répertoire cible**

Ouvrez le fichier `author_change.py` et ajustez la variable `directory` pour pointer vers le dossier contenant vos fichiers Office.

4. **Définir le nouvel auteur**

Modifiez la valeur de `new_author` dans le script.

5. **Exécuter le script**

Lancez le script avec :
`python author_change.py`

> [!WARNING]
> 1. Fonctionne uniquement sous Windows (utilisation de win32com.client).
> 2. Compatible avec Word (.docx), Excel (.xlsx) et PowerPoint (.pptx).
> 3. Nécessite que Microsoft Office soit installé sur la machine exécutant le script.

# Auteur

Créé par Zizouillo - 2025
