# Règles de travail Claude Code — REP_PEDA

## 1. Git pull obligatoire au démarrage

Avant toute intervention sur un fichier, exécuter :

```bash
git pull origin main
```

Ne jamais modifier un fichier sans avoir synchronisé avec main au préalable.

## 2. Commits atomiques et conventionnels

Chaque modification distincte = un commit séparé. Prefixes obligatoires :

- `feat:` — nouvelle fonctionnalité
- `fix:` — correction de bug
- `refactor:` — restructuration sans changement de comportement
- `docs:` — documentation uniquement

Exemple : `git commit -m "feat: ajout du module de suivi des élèves"`

## 3. Push direct sur main

GitHub Pages ne sert que `main` — impossible de tester une branche en ligne avant de merger. On travaille donc directement sur `main` :

```bash
git push origin main
```

L'historique git est le filet de sécurité. En cas de problème : `git log --oneline index.html` pour identifier le bon commit, puis `git show <hash>:index.html > index.html` pour restaurer.
