# TextEnhancer AI

Utilitaire Windows 11 qui améliore n'importe quel texte sélectionné via un LLM (OpenRouter), en un raccourci clavier — avec compilation automatique de l'exécutable et mises à jour automatiques via GitHub.

## Ce que ça fait

1. Tu sélectionnes du texte dans n'importe quelle application (Word, Outlook, Notion, navigateur…).
2. Tu presses ton raccourci (**Alt+Q** par défaut, modifiable).
3. Une fenêtre d'aperçu s'ouvre : tu choisis un **style** (preset) et le **modèle**, le résultat est généré.
4. Tu peux **éditer** le résultat, le **regénérer**, puis le **coller** (en vrai texte enrichi, gras réel) ou le **copier**.

Tout est paramétrable : clé API, modèle par défaut, presets (instructions), raccourci, démarrage automatique, dépôt GitHub pour les mises à jour.

---

## Installation pour l'utilisateur final (le plus simple)

**Tu n'as pas besoin de Python.** Une fois le dépôt en place (voir plus bas), l'exe est compilé par GitHub et publié dans les *Releases*.

1. Va sur la page **Releases** du dépôt GitHub.
2. Télécharge **`TextEnhancerAI.exe`**.
3. Double-clique dessus → une icône bleue apparaît dans la barre des tâches (près de l'horloge, parfois sous le chevron `^`).
4. Au 1er lancement, la fenêtre **Paramètres** s'ouvre : colle ta clé API OpenRouter, vérifie le modèle et le dépôt GitHub, enregistre.

> SmartScreen peut afficher un avertissement (exe non signé) : *Informations complémentaires → Exécuter quand même*.

---

## Mettre le projet sur GitHub

Le dépôt contient :

```
main.py                     # l'application
requirements.txt            # dépendances Python
build_exe.bat               # compilation locale (optionnel)
run.bat                     # lancement en mode dev (optionnel)
.github/workflows/build.yml # compilation automatique de l'exe (GitHub Actions)
.gitignore  LICENSE  README.md
```

### Étapes

1. **Crée un dépôt public** sur GitHub (ex. `text-enhancer-ai`). Ne coche pas « Add a README ».
2. Dans `main.py`, renseigne ton dépôt (ligne `GITHUB_REPO = "OWNER/REPO"`), par ex. `julienrayess/text-enhancer-ai`. *(Optionnel : tu peux aussi le mettre seulement dans Paramètres → Dépôt GitHub.)*
3. Pousse les fichiers. Depuis le dossier du projet, en remplaçant l'URL :

   ```bash
   git init
   git add .
   git commit -m "Initial commit - TextEnhancer AI"
   git branch -M main
   git remote add origin https://github.com/OWNER/REPO.git
   git push -u origin main
   ```

   *(Pas à l'aise avec git ? Tu peux aussi glisser-déposer tous les fichiers via le bouton « Add file → Upload files » sur la page du dépôt. Pense à inclure le dossier `.github`.)*

---

## Publier une version (et générer l'exe)

L'exe est fabriqué par GitHub Actions. Pour publier la **v1.0.0** :

```bash
git tag v1.0.0
git push origin v1.0.0
```

GitHub compile alors `TextEnhancerAI.exe` sur une machine Windows et crée automatiquement une **Release** avec l'exe en téléchargement. (Tu peux suivre la progression dans l'onglet **Actions**.)

Tu peux aussi lancer la compilation à la main : onglet **Actions → Build Windows EXE → Run workflow** ; l'exe est alors disponible en *artifact*.

> ⚠️ Le numéro de tag doit correspondre à `APP_VERSION` dans `main.py`. Pour une nouvelle version : change `APP_VERSION` (ex. `1.0.1`), commit, puis `git tag v1.0.1 && git push origin v1.0.1`.

---

## Mises à jour automatiques

Au démarrage, l'app interroge GitHub. Si une version plus récente existe :

- une fenêtre te **propose** de mettre à jour (tu gardes le contrôle) ;
- si tu acceptes, l'app **télécharge le nouvel exe, se ferme, se remplace et redémarre** toute seule ;
- tes réglages (clé API, presets) sont dans `%APPDATA%\TextEnhancerAI\` → **conservés** à chaque mise à jour.

Tu peux aussi forcer une vérif : clic droit sur l'icône → **Vérifier les mises à jour**. Et désactiver la vérif au démarrage dans les Paramètres.

> L'auto-update ne fonctionne que sur l'exe compilé (pas en mode script `main.py`), et nécessite un dépôt **public** avec au moins une Release publiée.

---

## Alternative : compiler / lancer en local (sans GitHub)

- **Compiler l'exe toi-même** : installe [Python 3.9+](https://www.python.org/downloads/) (coche *Add Python to PATH*), puis double-clique sur **`build_exe.bat`**. L'exe apparaît dans `dist\TextEnhancerAI.exe`.
- **Lancer sans compiler (mode dev)** : double-clique sur **`run.bat`**.

---

## Prérequis

- **Windows 10/11**
- Une **clé API OpenRouter** ([openrouter.ai/keys](https://openrouter.ai/keys)), avec du crédit pour les modèles payants.
- Pour *développer/compiler en local* uniquement : Python 3.9+.

---

## Réglages stockés

`%APPDATA%\TextEnhancerAI\config.json` (clé API, presets, raccourci, dépôt).
En cas de plantage : `crash.log` dans le même dossier.

---

## Dépannage

- **Le raccourci ne capture rien dans une app lancée en administrateur** : lance aussi `TextEnhancerAI.exe` en administrateur.
- **Le gras n'apparaît pas** : utilise « Coller (texte enrichi) ». Dans un éditeur 100 % texte (Bloc-notes), utilise « brut ».
- **Erreur OpenRouter 401/402** : clé invalide ou crédit insuffisant.
- **L'auto-update dit « dépôt non configuré »** : renseigne `OWNER/REPO` dans Paramètres → Dépôt GitHub, et vérifie qu'une Release existe.
- **GitHub Actions échoue** : ouvre l'onglet Actions pour lire le log ; le plus souvent une dépendance manquante dans `requirements.txt`.
