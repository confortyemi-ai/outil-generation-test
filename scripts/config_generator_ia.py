#!/usr/bin/env python3
"""
config_generator_ia.py — Interactive AI Agent for Test Management (Manual + Jira)

Usage:
    python scripts/config_generator_ia.py

Modes:
    - Manual: Chat IA pour générer config + Excel
    - Jira: Récupère User Stories → Génère cas de test → Crée dans Jira
"""

import os
import json
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

CLAUDE_API_KEY = os.getenv('CLAUDE_API_KEY')
GITHUB_TOKEN = os.getenv('GITHUB_TOKEN')
GITHUB_USERNAME = os.getenv('GITHUB_USERNAME', 'confortyemi-ai')
JIRA_URL = os.getenv('JIRA_URL')
JIRA_EMAIL = os.getenv('JIRA_EMAIL')
JIRA_API_TOKEN = os.getenv('JIRA_API_TOKEN')

if not CLAUDE_API_KEY:
    print("❌ Erreur : CLAUDE_API_KEY non trouvée dans .env")
    sys.exit(1)

import anthropic


class TestConfigGenerator:
    """Agent IA pour générer les configurations de test management"""
    
    def __init__(self, mode='manual'):
        self.client = anthropic.Anthropic(api_key=CLAUDE_API_KEY)
        self.mode = mode
        self.config = {
            "project": {
                "name": "",
                "environment": "",
                "methodology": "",
                "language": "fr",
                "testLevels": []
            },
            "deliverable": "cas-test",
            "testCases": {
                "functionName": "",
                "userStory": "",
                "includeCases": {
                    "nominal": True,
                    "boundary": True,
                    "error": True
                },
                "testTypes": ["fonctionnel"],
                "priority": "Critique"
            }
        }
        self.conversation = []
    
    def chat_with_ia(self, user_message):
        """Converser avec Claude"""
        self.conversation.append({
            "role": "user",
            "content": user_message
        })
        
        response = self.client.messages.create(
            model="claude-opus-4-1",
            max_tokens=1000,
            system="""Tu es un expert QA spécialisé dans la génération de documents de test management.

Ton rôle :
1. Poser des questions intelligentes et ciblées
2. Collecter les informations pour générer une configuration
3. Être conversationnel et professionnel
4. Résumer et confirmer avant de générer

Champs à couvrir :
- Nom du projet
- Environnement (web/api/mobile/desktop)
- Méthodologie (agile/waterfall/cycle-en-v)
- Type de livrable (plan-test/cas-test)
- Fonction à tester
- User Story
- Priorité

Réponds toujours en français. Sois concis.""",
            messages=self.conversation
        )
        
        assistant_message = response.content[0].text
        self.conversation.append({
            "role": "assistant",
            "content": assistant_message
        })
        
        return assistant_message
    
    def start_manual_mode(self):
        """Mode Manuel : Chat IA normal"""
        print("\n" + "="*70)
        print("🤖 MANUAL MODE - Configuration Interactive")
        print("="*70)
        print("\nType 'exit' pour quitter, 'config' pour voir la config")
        print("="*70 + "\n")
        
        first_message = self.chat_with_ia(
            "Commençons ! Peux-tu démarrer par me poser des questions pour créer une configuration ?"
        )
        print(f"🤖 Agent: {first_message}\n")
        
        while True:
            try:
                user_input = input("👤 Vous: ").strip()
                
                if user_input.lower() == 'exit':
                    print("\n❌ Annulé.")
                    sys.exit(0)
                
                if user_input.lower() == 'config':
                    print("\n📋 Configuration:")
                    print(json.dumps(self.config, indent=2, ensure_ascii=False))
                    continue
                
                if not user_input:
                    continue
                
                response = self.chat_with_ia(user_input)
                print(f"\n🤖 Agent: {response}\n")
                
                if "confirmer" in response.lower() or "résumé" in response.lower():
                    confirm = input("\n👤 D'accord ? (oui/non): ").strip().lower()
                    if confirm == 'oui':
                        print("\n✅ Config finalisée !")
                        self.parse_conversation_to_config()
                        break
            
            except KeyboardInterrupt:
                print("\n\n❌ Annulé.")
                sys.exit(0)
    
    def start_jira_mode(self):
        """Mode Jira : Récupère Stories et génère cas"""
        print("\n" + "="*70)
        print("📊 JIRA MODE - Generate from User Stories")
        print("="*70)
        
        # Vérifier les infos Jira
        if not all([JIRA_URL, JIRA_EMAIL, JIRA_API_TOKEN]):
            print("\n❌ Infos Jira manquantes dans .env")
            print("   Ajoute: JIRA_URL, JIRA_EMAIL, JIRA_API_TOKEN")
            sys.exit(1)
        
        print("\n✅ Infos Jira trouvées")
        print(f"   URL: {JIRA_URL}")
        print(f"   Email: {JIRA_EMAIL}")
        
        # Lancer gen_test_cases_from_jira_us.py
        print("\n⏳ Lancement du générateur Jira...")
        
        try:
            result = subprocess.run(
                [sys.executable, "scripts/gen_test_cases_from_jira_us.py"],
                capture_output=False,
                text=True
            )
            sys.exit(result.returncode)
        
        except FileNotFoundError:
            print("❌ gen_test_cases_from_jira_us.py non trouvé")
            sys.exit(1)
    
    def parse_conversation_to_config(self):
        """Extraire les infos de la conversation"""
        extraction_prompt = f"""
        Basé sur cette conversation, extrais les infos de config.
        
        Conversation:
        {json.dumps(self.conversation, indent=2, ensure_ascii=False)}
        
        Retourne UNIQUEMENT un JSON (pas d'autre texte):
        {{
            "projectName": "",
            "environment": "web|api|mobile|desktop",
            "methodology": "agile|waterfall|cycle-en-v",
            "deliverable": "plan-test|cas-test",
            "functionName": "",
            "userStory": "",
            "priority": "Critique|Haute|Moyenne|Basse"
        }}
        """
        
        try:
            response = self.client.messages.create(
                model="claude-opus-4-1",
                max_tokens=500,
                messages=[{"role": "user", "content": extraction_prompt}]
            )
            
            extracted_json = response.content[0].text
            if "```json" in extracted_json:
                extracted_json = extracted_json.split("```json")[1].split("```")[0]
            elif "```" in extracted_json:
                extracted_json = extracted_json.split("```")[1].split("```")[0]
            
            extracted = json.loads(extracted_json)
            
            self.config["project"]["name"] = extracted.get("projectName", "Test Project")
            self.config["project"]["environment"] = extracted.get("environment", "web")
            self.config["project"]["methodology"] = extracted.get("methodology", "agile")
            self.config["deliverable"] = extracted.get("deliverable", "cas-test")
            self.config["testCases"]["functionName"] = extracted.get("functionName", "")
            self.config["testCases"]["userStory"] = extracted.get("userStory", "US-001")
            self.config["testCases"]["priority"] = extracted.get("priority", "Critique")
        
        except Exception as e:
            print(f"⚠️  Erreur extraction: {e}")
    
    def save_config(self):
        """Sauvegarder la config"""
        config_dir = Path("config")
        config_dir.mkdir(exist_ok=True)
        
        project_slug = self.config["project"]["name"].replace(" ", "_").lower()
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"config_{project_slug}_{timestamp}.json"
        filepath = config_dir / filename
        
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, indent=2, ensure_ascii=False)
        
        print(f"✅ Config sauvegardée : {filepath}")
        return filepath
    
    def generate_report(self):
        """Générer l'Excel"""
        print("\n⏳ Génération du rapport...")
        
        config_file = self.save_config()
        
        try:
            result = subprocess.run(
                [sys.executable, "scripts/gen_master_report.py", str(config_file)],
                capture_output=True,
                text=True,
                timeout=30
            )
            
            if result.returncode == 0:
                print("✅ Rapport généré !")
                reports_dir = Path("reports")
                if reports_dir.exists():
                    reports = sorted(reports_dir.glob("master_report_*.xlsx"), 
                                   key=lambda x: x.stat().st_mtime, reverse=True)
                    if reports:
                        latest_report = reports[0]
                        print(f"📊 {latest_report}")
                        return latest_report
            else:
                print(f"❌ Erreur : {result.stderr}")
                return None
        
        except subprocess.TimeoutExpired:
            print("❌ Timeout")
            return None
        except Exception as e:
            print(f"❌ Erreur : {e}")
            return None
    
    def push_to_github(self, config_file):
        """Pousser sur GitHub"""
        if not GITHUB_TOKEN:
            return
        
        confirm = input("\n📤 Pousser sur GitHub ? (oui/non): ").strip().lower()
        if confirm != 'oui':
            return
        
        try:
            import base64
            import requests
            
            with open(config_file, 'r', encoding='utf-8') as f:
                content = f.read()
            
            content_b64 = base64.b64encode(content.encode()).decode()
            
            url = f"https://api.github.com/repos/{GITHUB_USERNAME}/outil-generation-test/contents/config/{config_file.name}"
            headers = {
                "Authorization": f"token {GITHUB_TOKEN}",
                "Content-Type": "application/json"
            }
            data = {
                "message": f"test: add config for {self.config['project']['name']}",
                "content": content_b64
            }
            
            response = requests.put(url, headers=headers, json=data)
            
            if response.status_code == 201:
                print("✅ Poussé sur GitHub !")
            else:
                print(f"❌ Erreur: {response.status_code}")
        
        except ImportError:
            subprocess.run([sys.executable, "-m", "pip", "install", "requests"], 
                         capture_output=True)
            self.push_to_github(config_file)
        except Exception as e:
            print(f"❌ Erreur : {e}")
    
    def run(self):
        """Exécuter"""
        try:
            if self.mode == 'manual':
                self.start_manual_mode()
                report_file = self.generate_report()
                
                if report_file:
                    print(f"\n✅ Rapport : {report_file}")
                    config_file = sorted(Path("config").glob("config_*_*.json"), 
                                       key=lambda x: x.stat().st_mtime, reverse=True)[0]
                    self.push_to_github(config_file)
            
            elif self.mode == 'jira':
                self.start_jira_mode()
        
        except Exception as e:
            print(f"\n❌ Erreur : {e}")
            sys.exit(1)


def main():
    """Menu principal"""
    print("\n" + "="*70)
    print("🎯 TEST MANAGEMENT CONFIG GENERATOR")
    print("="*70)
    print("\nQuel mode voulez-vous ?")
    print("  1. Manual - Chat IA normal")
    print("  2. Jira - Récupérer les User Stories et générer les cas")
    print("="*70)
    
    choice = input("\nChoix (1 ou 2): ").strip()
    
    if choice == '1':
        generator = TestConfigGenerator(mode='manual')
        generator.run()
    elif choice == '2':
        generator = TestConfigGenerator(mode='jira')
        generator.run()
    else:
        print("❌ Choix invalide")
        sys.exit(1)


if __name__ == "__main__":
    main()
