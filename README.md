# DOCX Parser

Regroupe et génère des fiches d'animations en Markdown prêtes à publier sur le CMS [Grav](https://getgrav.org/) à partir de fichiers DOCX.

## Installation

Requiert [Python](https://www.python.org/). Il faut ensuite installer les modules requis :

```console
pip install -r requirements.txt
```

## Utilisation

Exécuter le fichier *main.py* en donnant en argument le chemin vers un dossier. Tous les documents nommés *Fiche animation.docx* sous ce répetoire (en incluant les sous-dossiers) seront traités. Pour chacun, le script génère un dossier au nom standardisé, contenant un descriptif en Markdown ainsi que les différentes ressources voisines du fichier DOCX source. Ce dossier peut directement être téléversé sur un serveur utilisant le CMS Grav, dans le dossier `/user/pages/`, afin d'être disponible en ligne.