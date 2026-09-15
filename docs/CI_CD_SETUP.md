# CI/CD Setup — GitHub Actions Workflow

## 📋 Overview

Ce workflow automatise la génération du Master Test Report chaque fois qu'une config.json est poussée.

## 🚀 Quick Start

### 1. Ajouter une config

```bash
cat > config/config_myproject.json << JSON
{
  "project": { "name": "MyProject", "environment": "web" },
  "deliverable": "cas-test"
}
JSON
```

### 2. Push

```bash
git add config/config_myproject.json
git commit -m "test: add config for MyProject"
git push origin main
```

### 3. Récupérer le rapport

- Aller à **Actions → [workflow run] → Artifacts**
- Télécharger `test-reports`
- Ouvrir `master_report_*.xlsx`

## 🔄 Triggers

Le workflow se déclenche sur :
- ✅ Push de `config*.json`
- ✅ Pull Request avec config
- ✅ Manual trigger (GitHub UI)
- ✅ Tags (Release)

## 📊 Résultat

```
master_report_{project}_{timestamp}.xlsx
├─ Onglet 1: Cas de test + Priorisation + Résultats
├─ Onglet 2: Couverture
├─ Onglet 3: Priorisation
├─ Onglet 4: Rapport
└─ Onglet 5: Dashboard
```

Pour plus de détails : voir [PROJECT_STRUCTURE.md](PROJECT_STRUCTURE.md)
