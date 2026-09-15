# Mode Jira - Générer les cas de test à partir des User Stories

## 🎯 Workflow

```
1. Connexion Jira (via .env)
2. Sélectionner un Epic
3. Sélectionner les User Stories
4. Claude IA génère les cas de test automatiquement
5. Nomenclature v4: [TC][FONCTION][TYPE]-NN
6. Créer les cas dans Jira (optionnel)
```

## 🔧 Configuration .env

Ajoute ces lignes à `.env` :

```
JIRA_URL=https://votre-instance.atlassian.net
JIRA_EMAIL=votre.email@company.com
JIRA_API_TOKEN=votre_token_api_jira
```

### Comment générer ton token Jira ?

1. Va à https://id.atlassian.com/manage-profile/security/api-tokens
2. Crée un nouveau token
3. Copie le token dans `.env`

## 🚀 Utilisation

```bash
python scripts/config_generator_ia.py

# Choix 2 pour "Jira Mode"
# Sélectionner Epic → Sélectionner Stories
# Claude génère les cas → Optionnel: créer dans Jira
```

## 📊 Nomenclature des cas (v4)

```
[TC][FONCTION][TYPE]-NN

Exemple:
  TC_LOGIN_NOMINAL-001
  TC_LOGIN_ERROR-002
  TC_CHECKOUT_BOUNDARY-001

Types:
  - NOMINAL: Cas heureux
  - BOUNDARY: Cas limites
  - ERROR: Cas d'erreur
```

## 🔗 Fichiers connexes

- `api_jira_wrapper.py` — Wrapper Jira API
- `gen_test_cases_from_jira_us.py` — Générateur IA depuis US
- `create_xray_tests.py` — Créer les cas dans Jira
