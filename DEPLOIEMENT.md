# Déploiement en ligne — version Celeragens (URL publique gratuite)

Ce dossier (`celeragens_app/`) est une version **autonome** de l'application,
identique en fonctionnement à la version Karbonpath mais avec l'identité
visuelle Celeragens (logo, nom, couleurs teal/turquoise). Elle contient sa
propre copie de `rapport_core.py` : vous pouvez la déployer indépendamment,
sans dépendre des autres fichiers du dossier `CLAUDE`.

## Fichiers nécessaires (déjà prêts dans ce dossier)

- `app.py` — l'application web (charte graphique Celeragens)
- `rapport_core.py` — logique métier (copie identique à la version Karbonpath)
- `generate_rapport.py` — script en ligne de commande (facultatif pour le web)
- `requirements.txt` — dépendances Python
- `celeragens_logo.png` — logo affiché en en-tête de l'app
- `.streamlit/config.toml` — thème de couleurs (teal Celeragens #008080)

## Option A — dépôt GitHub dédié (le plus simple)

### 1. Créer un compte GitHub (si vous n'en avez pas)
https://github.com/signup

### 2. Créer un nouveau dépôt
- "New repository", nom au choix (ex. `rapport-rh-celeragens`)
- Visibilité **Private** recommandée (aucune donnée RH n'est jamais stockée
  dans le dépôt — seul le code y va, les fichiers sont uploadés en mémoire à
  chaque utilisation)

### 3. Uploader le contenu de ce dossier en conservant la structure
"Add file" → "Upload files", glissez le **contenu du dossier `celeragens_app/`**
(pas le dossier lui-même, son contenu) de façon à obtenir à la racine du
dépôt :

```
app.py
rapport_core.py
generate_rapport.py
requirements.txt
celeragens_logo.png
.streamlit/config.toml
```

`.streamlit/config.toml` ne peut pas être glissé isolément dans un
sous-dossier via "Upload files" : utilisez "Add file" → "Create new file",
tapez `.streamlit/config.toml` comme nom (GitHub crée le dossier
automatiquement), et collez-y le contenu fourni.

Validez ("Commit changes").

### 4. Déployer sur Streamlit Community Cloud
https://share.streamlit.io → connectez-vous avec GitHub → "New app" →
sélectionnez le dépôt, branche `main`, main file path `app.py` → "Deploy".

Vous obtenez une URL publique du type `https://VOTRE-APP.streamlit.app`.

## Option B — même dépôt GitHub que la version Karbonpath

Si vous préférez un seul dépôt pour les deux versions : uploadez ce dossier
`celeragens_app/` tel quel (avec son propre sous-dossier) à côté des fichiers
Karbonpath existants à la racine. Sur share.streamlit.io, créez ensuite une
**deuxième** app ("New app") sur le même dépôt, en indiquant cette fois le
main file path `celeragens_app/app.py`. Chaque app obtient sa propre URL
publique et son propre thème (chacune lit le `.streamlit/config.toml` de son
propre dossier).

## Mettre à jour l'application

Toute modification de `app.py` ou `rapport_core.py` dans ce dossier :
re-uploadez le fichier modifié dans le dépôt GitHub correspondant. Streamlit
Community Cloud redéploie automatiquement à chaque commit.

## Confidentialité des données

Les fichiers sources sont uploadés en mémoire pendant la session de
l'utilisateur (jamais écrits sur disque côté serveur, jamais stockés dans le
dépôt GitHub). Gardez le dépôt en **Private** si les données sont sensibles,
et activez si besoin la protection par mot de passe / liste d'utilisateurs
autorisés (réglages de l'app sur share.streamlit.io → "Settings" →
"Sharing").

## Alternative : lancer l'app en local (sans déploiement)

```bash
cd celeragens_app
pip install -r requirements.txt
streamlit run app.py
```
