# Vigie NPE — Démonstrateur CelerAgens

Application de démonstration illustrant les cas d'usage 1 (contrôle de planification) et 2 (contrôle continu du réel) d'un agent IA de contrôle interne du nombre de personnel d'encadrement (NPE) dans les crèches Babilou, dans le cadre de la réforme du chèque-service accueil (CSA) au Luxembourg.

L'application elle-même (`vigie-npe.html`) est une page HTML/CSS/JS autonome ; `app.py` est un habillage Streamlit minimal qui l'intègre telle quelle.

## Lancer en local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Déployer sur Streamlit Community Cloud

1. Pousser ce dépôt sur GitHub (déjà fait si vous lisez ceci depuis le repo).
2. Aller sur [share.streamlit.io](https://share.streamlit.io), se connecter avec le compte GitHub propriétaire du dépôt.
3. **New app** → choisir ce dépôt `ppcoulon/CELERAGENS`, la branche (`main`), et `vigie-npe/app.py` comme fichier principal (ce dossier cohabite avec d'autres applications du même dépôt — chaque app Streamlit Cloud pointe vers son propre sous-dossier).
4. Déployer — Streamlit Cloud installe automatiquement `requirements.txt` et publie l'application avec une URL `https://<nom>.streamlit.app`.

## Mettre à jour l'application

Modifier `vigie-npe.html` (et `app.py` si besoin), commiter et pousser sur `main` : Streamlit Cloud redéploie automatiquement à chaque push.
