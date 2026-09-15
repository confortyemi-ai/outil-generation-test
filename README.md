# Outil de Génération de Documents de Test Management

Automatisation complète de la génération de documents de test management (plans, cas de test, rapports) avec **GitHub Actions CI/CD**.

## 🚀 Quick Start

### 1. Ajouter une configuration

Crée un fichier `config/config_myproject.json` :

```json
{
  "project": {
    "name": "My Project",
    "environment": "web",
    "methodology": "agile"
  },
  "deliverable": "cas-test"
}
```

### 2. Push et automatisation

```bash
git add config/config_myproject.json
git commit -m "test: add config for MyProject"
git push origin main
```

### 3. Résultat

✅ GitHub Actions génère automatiquement le rapport  
✅ Artifact disponible dans Actions → Artifacts  
✅ Télécharge `master_report_*.xlsx`

## 📁 Structure

```
.github/workflows/           # Workflows CI/CD
scripts/                     # Scripts Python (gen_*.py)
config/                      # Configurations JSON
reports/                     # Outputs générés (ne pas commit)
docs/                        # Documentation
```

## 📖 Documentation

- [CI/CD Setup](docs/CI_CD_SETUP.md)
- [Project Structure](docs/PROJECT_STRUCTURE.md)
- [Config Schema](docs/CONFIG_SCHEMA.md)

## 🛠️ Technologies

- **Python 3.11** — openpyxl, python-docx
- **GitHub Actions** — CI/CD automation
- **Excel** — master_report.xlsx (5 onglets)

## 📊 Artefacts générés

| Nom | Format | Contenu |
|-----|--------|---------|
| master_report_*.xlsx | Excel | Cas de test + Priorisation + Résultats + Couverture + Dashboard |

## ✅ Status

- ✅ Workflow GitHub Actions configuré
- ✅ Scripts Python opérationnels
- ✅ CI/CD prêt à l'emploi

## 📝 Licence

MIT
