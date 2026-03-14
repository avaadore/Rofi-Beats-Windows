# Publish to GitHub

This document covers a clean GitHub release flow for this Windows rewrite.

## 1) Create a new repository on GitHub
- Create a normal new repository under your account or organization.
- Do not use GitHub's `Fork` button.
- Keep `LICENSE` as GPL-3.0 unless you later decide to re-evaluate licensing.

## 2) Point this local repo to the new GitHub repository
If `origin` still points at the original project, rename it first:

```powershell
git remote rename origin reference
git remote add origin https://github.com/<YOUR_USER>/<YOUR_REPO>.git
```

If this repo has no `origin` yet, just add the new one:

```powershell
git remote add origin https://github.com/<YOUR_USER>/<YOUR_REPO>.git
```

## 3) Commit and push

```powershell
git add .
git commit -m "Prepare public Windows release"
git branch -M main
git push -u origin main
```

## 4) Optional: keep the old remote only as a reference

```powershell
git remote -v
```

You can keep the renamed remote for historical reference, or remove it later if you no longer need it.
