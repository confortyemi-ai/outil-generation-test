# Configuration Schema

## Format JSON

```json
{
  "project": {
    "name": "string",
    "environment": "web|api|mobile|desktop",
    "methodology": "agile|waterfall|cycle-en-v",
    "language": "fr|en",
    "testLevels": ["unit|integration|system|acceptance"]
  },
  "deliverable": "plan-test|cas-test",
  "testPlan": {
    "scope": "string",
    "standards": ["IEEE 829", "ISO/IEC 29119"],
    "risks": ["string"]
  },
  "testCases": {
    "functionName": "string",
    "userStory": "string",
    "includeCases": {
      "nominal": boolean,
      "boundary": boolean,
      "error": boolean
    },
    "testTypes": ["fonctionnel", "régression"],
    "priority": "Critique|Haute|Moyenne|Basse"
  }
}
```

## Exemples

Voir : `config/config_example_*.json`
