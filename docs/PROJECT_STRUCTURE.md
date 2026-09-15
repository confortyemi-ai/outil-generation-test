# Project Structure

```
outil-generation-test/
├── .github/workflows/
│   └── generate-test-report.yml
├── scripts/
│   ├── gen_master_report.py
│   ├── gen_plan_test.py
│   ├── gen_test_cases.py
│   ├── gen_coverage.py
│   ├── gen_prioritization.py
│   ├── gen_report.py
│   └── gen_dashboard.py
├── config/
│   ├── config_template.json
│   ├── config_example_regulatory.json
│   └── config_example_commia.json
├── reports/           (gitignored)
├── docs/
│   ├── CI_CD_SETUP.md
│   ├── PROJECT_STRUCTURE.md
│   └── CONFIG_SCHEMA.md
├── .gitignore
├── requirements.txt
└── README.md
```

## 🔑 Fichiers clés

| Fichier | Rôle |
|---------|------|
| `.github/workflows/generate-test-report.yml` | Trigger GitHub Actions |
| `scripts/gen_master_report.py` | Script principal |
| `config/*.json` | Configurations par projet |
| `reports/` | Outputs (ne pas commit) |

Pour exécuter manuellement :
```bash
python scripts/gen_master_report.py config/config_myproject.json
```
