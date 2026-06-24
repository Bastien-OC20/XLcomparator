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

---

## Ce que le comparateur analyse

### Périmètre de l'analyse

Le comparateur charge **l'intégralité** de chaque feuille des deux fichiers Excel, sans en-tête implicite : chaque ligne, y compris la première, est traitée comme une ligne de données ordinaire.

### 1. Feuilles (onglets)

L'ensemble des feuilles présentes dans les deux fichiers est collecté. Pour chaque feuille :

- Si elle existe dans les **deux** fichiers → son contenu est comparé.
- Si elle n'existe que dans **l'un** des deux fichiers → les cellules du fichier qui la possède sont comparées à des cellules vides de l'autre, ce qui génère autant de différences qu'il y a de cellules non vides.

### 2. Titres de colonnes

Les titres de colonnes (première ligne d'une feuille) **sont inclus dans la comparaison** au même titre que n'importe quelle autre cellule. Si un en-tête a été renommé ou modifié entre les deux fichiers, cela apparaît comme une différence sur la ligne 1 de la colonne concernée.

### 3. Colonnes comparées

Les colonnes sont identifiées par leur **lettre Excel** (A, B, C, …, Z, AA, AB, …). Le comparateur aligne les deux feuilles sur la forme maximale (nombre de lignes × nombre de colonnes) en complétant les cellules manquantes par une valeur vide. Toutes les colonnes présentes dans au moins l'un des deux fichiers sont donc couvertes.

### 4. Valeurs présentes et différences détectées

Chaque cellule est convertie en chaîne de caractères avant comparaison. Une **différence** est enregistrée dès que la valeur de la cellule `(ligne, colonne)` du fichier de référence ne correspond pas à celle du fichier comparé. Pour chaque différence, le comparateur retient :

| Champ | Description |
|---|---|
| **Feuille** | Nom de l'onglet Excel concerné |
| **Ligne** | Numéro de ligne (base 1, comme dans Excel) |
| **Colonne** | Lettre de colonne Excel (A, B, C…) |
| **Valeur référence** | Contenu de la cellule dans le fichier de référence |
| **Valeur comparée** | Contenu de la cellule dans le fichier à comparer |

### Ce que le comparateur ne prend pas en compte

- La **mise en forme** des cellules (couleur, police, bordures, etc.)
- Les **formules** : seule la valeur calculée (ou le texte brut) est comparée
- Les **métadonnées** du classeur (auteur, date de création, propriétés du document)

---

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
