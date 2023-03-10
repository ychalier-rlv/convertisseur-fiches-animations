# Convertisseur de fiches d'animations

Regroupe et génère des fiches d'animations en Markdown prêtes à publier sur le CMS [Grav](https://getgrav.org/) à partir de fichiers DOCX.

## Installation

Requiert [Python](https://www.python.org/). Il faut ensuite installer les modules requis :

```console
pip install -r requirements.txt
```

## Utilisation

Exécuter le fichier *main.py* en donnant en argument le chemin vers un dossier.

```console
python main.py ~/Documents/
```

Tous les documents nommés *Fiche animation.docx* sous ce répetoire, en incluant les sous-dossiers, seront traités. Pour chacun, le script génère un dossier au nom standardisé, contenant un descriptif en Markdown ainsi que les différentes ressources voisines du fichier DOCX source. Ce dossier peut directement être téléversé sur un serveur utilisant le CMS Grav, dans le dossier `/user/pages/`, afin d'être disponible en ligne.

### Structure DOCX

Afin d'être correctement analysés, les fichiers DOCX doivent suivre une certaine syntaxe :

- Renseigner un titre avec le style *Titre*.
- Renseigner les métadonnées en utilisant une ligne pour le nom, en gras, puis une ligne pour la valeur. Les noms acceptés sont :
  - Thématiques
  - Durée
  - Participants
  - Public
  - Prérequis
  - Matériel
- Renseigner des parties avec le style *Titre 1*. Ces parties doivent contenir au moins une intitulée *Déroulé*.
- Dans la partie *Déroulée*, chaque étape commence par un *Titre 2*. Ce titre peut également contenir une durée, au format `Titre (XX min)`.
- Pour définir un paragraphe comme étant du code, utiliser un style nommé *Code*.

Seuls quelques mises en forme sont supportées : gras, italique, liste, liens hypertextes.