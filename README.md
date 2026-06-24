# XLcomparator

Comparateur de fichiers Excel (`.xls` / `.xlsx`) avec interface **Streamlit**.

## Fonctionnalités

- Upload d'un fichier **référent** et d'un fichier **à comparer** (xls/xlsx)
- Comparaison feuille par feuille, cellule par cellule
- Affichage d'un tableau interactif des différences
- Filtrage des résultats par feuille
- Export des différences en :
  - **XML** — pour injection dans une base de données
  - **DOCX** — rapport Word
  - **XLSX** — tableau Excel

## Prérequis

- Python 3.10+
- Les dépendances listées dans `requirements.txt`

## Installation

```bash
pip install -r requirements.txt
```

## Lancement

```bash
streamlit run app.py
```

L'application s'ouvre automatiquement dans votre navigateur à l'adresse `http://localhost:8501`.

## Structure du projet

```
XLcomparator/
├── app.py            # Interface Streamlit
├── comparator.py     # Moteur de comparaison et exports
├── requirements.txt  # Dépendances Python
└── README.md
```
