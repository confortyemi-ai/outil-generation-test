#!/usr/bin/env python3
"""
gen_test_cases_from_jira_us.py — Générer les cas de test à partir des User Stories Jira

Workflow :
1. Récupère les Stories d'un Epic Jira
2. Claude IA génère les cas de test À PARTIR des Stories
3. Utilise nomenclature v4: [TC][FONCTION][SOUS-FONCTION]-NN
4. Crée les cas dans Jira via create_xray_tests.py
"""

import os
import json
import subprocess
import sys
from pathlib import Path
from dotenv import load_dotenv
from api_jira_wrapper import JiraAPI

load_dotenv()

import anthropic

class TestCaseGeneratorFromStories:
    """Générer les cas de test à partir des US Jira"""
    
    def __init__(self):
        self.claude_api_key = os.getenv('CLAUDE_API_KEY')
        if not self.claude_api_key:
            raise ValueError("❌ CLAUDE_API_KEY manquant dans .env")
        
        self.client = anthropic.Anthropic(api_key=self.claude_api_key)
        self.jira = JiraAPI()
        self.test_cases = []
    
    def generate_test_cases_from_stories(self, stories: list, epic_name: str) -> list:
        """Générer les cas de test à partir des Stories Jira"""
        
        # Préparer les stories pour Claude
        stories_text = "\n".join([
            f"- [{s['key']}] {s['summary']}\n  {s.get('description', '')}"
            for s in stories
        ])
        
        prompt = f"""Tu es un expert QA spécialisé dans la création de cas de test ISO/IEC 29119.

À partir des User Stories suivantes, génère les cas de test correspondants.

Projet/Epic: {epic_name}

User Stories:
{stories_text}

Pour chaque Story, génère :
1. Cas nominaux (happy path)
2. Cas limites (boundary conditions)
3. Cas d'erreur (error cases)

Format de réponse (JSON uniquement, pas d'autre texte) :
[
  {{
    "storyKey": "STORY-001",
    "functionName": "Extraction depuis la Story",
    "testCases": [
      {{
        "title": "Titre du cas",
        "description": "Description claire",
        "preconditions": "État du système avant",
        "steps": [
          {{"step": 1, "action": "Action 1", "expectedResult": "Résultat attendu 1"}},
          {{"step": 2, "action": "Action 2", "expectedResult": "Résultat attendu 2"}}
        ],
        "type": "fonctionnel",
        "priority": "Critique|Haute|Moyenne|Basse"
      }}
    ]
  }}
]

Nomenclature:
- ID: [TC][FONCTION][TYPE]-NN (ex: TC_LOGIN_NOMINAL-001, TC_LOGIN_ERROR-001)
- FONCTION: Extraite de la Story
- TYPE: NOMINAL, BOUNDARY, ERROR
- NN: Numéro séquentiel

Réponds UNIQUEMENT avec le JSON, pas d'explication."""

        print("\n⏳ Claude IA génère les cas de test...")
        
        response = self.client.messages.create(
            model="claude-opus-4-1",
            max_tokens=4000,
            messages=[{"role": "user", "content": prompt}]
        )
        
        # Extraire le JSON
        response_text = response.content[0].text
        
        # Nettoyer les markdown backticks
        if "```json" in response_text:
            response_text = response_text.split("```json")[1].split("```")[0]
        elif "```" in response_text:
            response_text = response_text.split("```")[1].split("```")[0]
        
        try:
            test_cases_data = json.loads(response_text)
            print(f"✅ {len(test_cases_data)} groupe(s) de cas générés")
            return test_cases_data
        except json.JSONDecodeError as e:
            print(f"❌ Erreur JSON: {e}")
            return []
    
    def format_for_xray(self, test_cases_data: list) -> dict:
        """Formater les cas pour create_xray_tests.py"""
        
        formatted = {
            "project": {
                "name": "Test Project",
                "environment": "web",
                "methodology": "agile",
                "language": "fr",
                "testLevels": ["system", "acceptance"]
            },
            "testCases": []
        }
        
        tc_count = 1
        for group in test_cases_data:
            story_key = group.get('storyKey', 'UNKNOWN')
            func_name = group.get('functionName', 'Unknown')
            
            for tc in group.get('testCases', []):
                # Construire l'ID en nomenclature v4
                type_code = "NOMINAL" if "nominal" in tc['title'].lower() else \
                           "BOUNDARY" if "limite" in tc['title'].lower() else "ERROR"
                tc_id = f"TC_{func_name.replace(' ', '_').upper()}_{type_code}-{tc_count:03d}"
                
                formatted_tc = {
                    "id": tc_id,
                    "title": tc['title'],
                    "module": func_name,
                    "userStory": story_key,
                    "type": tc.get('type', 'fonctionnel'),
                    "level": "system",
                    "priority": tc.get('priority', 'Moyenne'),
                    "preconditions": tc.get('preconditions', ''),
                    "data": tc.get('data', ''),
                    "steps": [
                        f"{s['step']}. {s['action']} → {s['expectedResult']}"
                        for s in tc.get('steps', [])
                    ],
                    "expectedResult": tc.get('expectedResult', 'Voir étapes'),
                    "status": "Non exécuté"
                }
                
                formatted["testCases"].append(formatted_tc)
                tc_count += 1
        
        return formatted
    
    def run(self):
        """Workflow complet"""
        try:
            print("\n" + "="*70)
            print("🤖 TEST CASE GENERATOR - Mode Jira User Stories")
            print("="*70)
            
            # Étape 1 : Récupérer les Stories
            project_key = input("\n📍 Quel est votre projet Jira (ex: XSP)? ").strip().upper()
            
            print(f"\n⏳ Récupération des Epics du projet {project_key}...")
            epics = self.jira.search_epics(project_key)
            
            if not epics:
                print("❌ Aucun Epic trouvé")
                return
            
            print(f"\n✅ {len(epics)} Epic(s) :\n")
            for i, epic in enumerate(epics, 1):
                print(f"  {i}. {epic['key']} — {epic['summary']}")
            
            epic_choice = int(input("\nQuel Epic (numéro)? ")) - 1
            selected_epic = epics[epic_choice]
            
            print(f"\n⏳ Récupération des Stories de {selected_epic['key']}...")
            stories = self.jira.search_stories_in_epic(selected_epic['key'])
            
            if not stories:
                print("❌ Aucune Story trouvée")
                return
            
            print(f"\n✅ {len(stories)} Story(ies) :\n")
            for i, story in enumerate(stories, 1):
                print(f"  {i}. {story['key']} — {story['summary']}")
            
            # Étape 2 : Générer les cas
            test_cases_data = self.generate_test_cases_from_stories(
                stories, 
                selected_epic['summary']
            )
            
            if not test_cases_data:
                print("❌ Erreur lors de la génération")
                return
            
            # Étape 3 : Formater pour Xray
            formatted_data = self.format_for_xray(test_cases_data)
            
            # Étape 4 : Sauvegarder
            config_dir = Path("config")
            config_dir.mkdir(exist_ok=True)
            
            timestamp = __import__('datetime').datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"config_xray_{selected_epic['key'].lower()}_{timestamp}.json"
            filepath = config_dir / filename
            
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(formatted_data, f, indent=2, ensure_ascii=False)
            
            print(f"\n✅ Config sauvegardée : {filepath}")
            print(f"📊 Total : {len(formatted_data['testCases'])} cas de test générés")
            
            # Étape 5 : Créer les cas dans Jira (optionnel)
            create_jira = input("\n📤 Créer les cas dans Jira/Xray ? (oui/non): ").strip().lower()
            if create_jira == 'oui':
                print("\n⏳ Création des cas dans Jira...")
                # Appeler create_xray_tests.py si disponible
                if Path("scripts/create_xray_tests.py").exists():
                    result = subprocess.run(
                        [sys.executable, "scripts/create_xray_tests.py", str(filepath)],
                        capture_output=True,
                        text=True
                    )
                    if result.returncode == 0:
                        print("✅ Cas créés dans Jira !")
                    else:
                        print(f"⚠️  {result.stderr}")
                else:
                    print("⚠️  create_xray_tests.py non trouvé")
            
            print("\n" + "="*70)
            print("✅ Workflow complété !")
            print("="*70)
        
        except Exception as e:
            print(f"\n❌ Erreur : {e}")
            import traceback
            traceback.print_exc()
            sys.exit(1)


if __name__ == "__main__":
    generator = TestCaseGeneratorFromStories()
    generator.run()
