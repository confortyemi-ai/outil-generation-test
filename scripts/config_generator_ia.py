#!/usr/bin/env python3
"""
config_generator_ia.py — Interactive AI Agent for Test Management Config Generation

Usage:
    python scripts/config_generator_ia.py

Features:
    - Chat interactif avec Claude IA
    - Génère config JSON automatiquement
    - Lance gen_master_report.py
    - Génère rapport Excel
    - Push optionnel sur GitHub
    - Téléchargement du fichier
"""

import os
import json
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv

# Charger les variables .env
load_dotenv()

# Récupérer les clés
CLAUDE_API_KEY = os.getenv('CLAUDE_API_KEY')
GITHUB_TOKEN = os.getenv('GITHUB_TOKEN')
GITHUB_USERNAME = os.getenv('GITHUB_USERNAME', 'confortyemi-ai')

# Vérifier les clés
if not CLAUDE_API_KEY:
    print("❌ Erreur : CLAUDE_API_KEY non trouvée dans .env")
    sys.exit(1)

import anthropic

class TestConfigGenerator:
    """Agent IA pour générer les configurations de test management"""
    
    def __init__(self):
        self.client = anthropic.Anthropic(api_key=CLAUDE_API_KEY)
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
        """Envoyer un message à Claude et obtenir une réponse"""
        self.conversation.append({
            "role": "user",
            "content": user_message
        })
        
        response = self.client.messages.create(
            model="claude-opus-4-1",
            max_tokens=1000,
            system="""Tu es un expert QA (Quality Assurance) spécialisé dans la génération de documents de test management.

Ton rôle :
1. Poser des questions intelligentes et ciblées pour comprendre les besoins de test
2. Collecter les informations pour générer une configuration de test management
3. Être conversationnel et professionnel
4. Résumer les réponses et confirmer la configuration avant la génération

Champs à couvrir :
- Nom du projet
- Environnement (web/api/mobile/desktop)
- Méthodologie (agile/waterfall/cycle-en-v)
- Type de livrable (plan-test/cas-test)
- Fonction à tester
- User Story / exigence
- Priorité des tests

Réponds toujours en français. Sois concis et pratique.""",
            messages=self.conversation
        )
        
        assistant_message = response.content[0].text
        self.conversation.append({
            "role": "assistant",
            "content": assistant_message
        })
        
        return assistant_message
    
    def start_conversation(self):
        """Démarrer la conversation avec l'utilisateur"""
        print("\n" + "="*70)
        print("🤖 AGENT IA - Test Management Config Generator")
        print("="*70)
        print("\nBonjour ! Je suis votre assistant QA. Je vais vous poser quelques")
        print("questions pour générer votre configuration de test management.")
        print("\nType 'exit' pour quitter, 'config' pour voir la config générée")
        print("="*70 + "\n")
        
        # Premier message de l'agent
        first_message = self.chat_with_ia(
            "Commençons ! Peux-tu démarrer par me poser des questions pour créer une configuration de test management ?"
        )
        print(f"🤖 Agent: {first_message}\n")
        
        # Boucle de conversation
        while True:
            try:
                user_input = input("👤 Vous: ").strip()
                
                if user_input.lower() == 'exit':
                    print("\n❌ Génération annulée.")
                    sys.exit(0)
                
                if user_input.lower() == 'config':
                    print("\n📋 Configuration générée :")
                    print(json.dumps(self.config, indent=2, ensure_ascii=False))
                    continue
                
                if not user_input:
                    continue
                
                # Converser avec l'agent
                response = self.chat_with_ia(user_input)
                print(f"\n🤖 Agent: {response}\n")
                
                # Vérifier si l'agent demande de confirmer
                if "confirmer" in response.lower() or "résumé" in response.lower():
                    confirm = input("\n👤 Êtes-vous d'accord ? (oui/non): ").strip().lower()
                    if confirm == 'oui':
                        print("\n✅ Configuration finalisée !")
                        self.parse_conversation_to_config()
                        break
                    elif confirm == 'non':
                        print("\nD'accord, continuons...")
            
            except KeyboardInterrupt:
                print("\n\n❌ Génération annulée.")
                sys.exit(0)
    
    def parse_conversation_to_config(self):
        """Extraire les infos de la conversation et mettre à jour la config"""
        # Demander à Claude d'extraire les infos
        extraction_prompt = f"""
        Basé sur la conversation suivante, extrais les informations de configuration pour un système de test management.
        
        Conversation:
        {json.dumps(self.conversation, indent=2, ensure_ascii=False)}
        
        Retourne UNIQUEMENT un objet JSON valide (pas d'autre texte) avec cette structure:
        {{
            "projectName": "nom du projet",
            "environment": "web|api|mobile|desktop",
            "methodology": "agile|waterfall|cycle-en-v",
            "deliverable": "plan-test|cas-test",
            "functionName": "fonction à tester",
            "userStory": "ID user story",
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
            # Nettoyer les markdown backticks si présents
            if "```json" in extracted_json:
                extracted_json = extracted_json.split("```json")[1].split("```")[0]
            elif "```" in extracted_json:
                extracted_json = extracted_json.split("```")[1].split("```")[0]
            
            extracted = json.loads(extracted_json)
            
            # Mettre à jour la config
            self.config["project"]["name"] = extracted.get("projectName", "Test Project")
            self.config["project"]["environment"] = extracted.get("environment", "web")
            self.config["project"]["methodology"] = extracted.get("methodology", "agile")
            self.config["deliverable"] = extracted.get("deliverable", "cas-test")
            self.config["testCases"]["functionName"] = extracted.get("functionName", "")
            self.config["testCases"]["userStory"] = extracted.get("userStory", "US-001")
            self.config["testCases"]["priority"] = extracted.get("priority", "Critique")
            
        except Exception as e:
            print(f"⚠️  Erreur lors de l'extraction: {e}")
            print("Utilisation de valeurs par défaut...")
    
    def save_config(self):
        """Sauvegarder la config JSON"""
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
        """Générer le rapport avec gen_master_report.py"""
        print("\n⏳ Génération du rapport Excel...")
        
        config_file = self.save_config()
        
        try:
            result = subprocess.run(
                [sys.executable, "scripts/gen_master_report.py", str(config_file)],
                capture_output=True,
                text=True,
                timeout=30
            )
            
            if result.returncode == 0:
                print("✅ Rapport généré avec succès !")
                # Trouver le fichier généré
                reports_dir = Path("reports")
                if reports_dir.exists():
                    reports = sorted(reports_dir.glob("master_report_*.xlsx"), 
                                   key=lambda x: x.stat().st_mtime, reverse=True)
                    if reports:
                        latest_report = reports[0]
                        print(f"📊 Fichier : {latest_report}")
                        print(f"📍 Chemin complet : {latest_report.absolute()}")
                        return latest_report
            else:
                print(f"❌ Erreur lors de la génération : {result.stderr}")
                return None
        
        except subprocess.TimeoutExpired:
            print("❌ La génération a pris trop de temps")
            return None
        except Exception as e:
            print(f"❌ Erreur : {e}")
            return None
    
    def push_to_github(self, config_file):
        """Pousser la config sur GitHub"""
        if not GITHUB_TOKEN:
            print("\n⚠️  Token GitHub non trouvé dans .env")
            return
        
        confirm = input("\n📤 Voulez-vous pousser cette config sur GitHub ? (oui/non): ").strip().lower()
        if confirm != 'oui':
            return
        
        print("\n⏳ Envoi vers GitHub...")
        
        try:
            import base64
            import requests
            
            repo_name = "outil-generation-test"
            
            with open(config_file, 'r', encoding='utf-8') as f:
                content = f.read()
            
            content_b64 = base64.b64encode(content.encode()).decode()
            
            url = f"https://api.github.com/repos/{GITHUB_USERNAME}/{repo_name}/contents/config/{config_file.name}"
            
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
                print("✅ Config poussée sur GitHub !")
                print(f"📍 https://github.com/{GITHUB_USERNAME}/{repo_name}/blob/main/config/{config_file.name}")
                print(f"\n🚀 Workflow GitHub Actions déclenché !")
                print(f"📊 Voir les artifacts : https://github.com/{GITHUB_USERNAME}/{repo_name}/actions")
            else:
                print(f"❌ Erreur GitHub : {response.status_code}")
                print(response.text)
        
        except ImportError:
            print("⚠️  'requests' non installé. Installation...")
            subprocess.run([sys.executable, "-m", "pip", "install", "requests"], 
                         capture_output=True)
            self.push_to_github(config_file)
        except Exception as e:
            print(f"❌ Erreur : {e}")
    
    def run(self):
        """Exécuter le workflow complet"""
        try:
            # Conversation avec l'agent IA
            self.start_conversation()
            
            # Générer le rapport
            report_file = self.generate_report()
            
            if report_file:
                print(f"\n✅ Succès ! Votre rapport est prêt :")
                print(f"   📊 {report_file}")
                
                # Ouvrir le fichier (optionnel)
                try:
                    import platform
                    if platform.system() == 'Darwin':  # macOS
                        subprocess.run(['open', str(report_file)])
                    elif platform.system() == 'Windows':
                        os.startfile(str(report_file))
                    elif platform.system() == 'Linux':
                        subprocess.run(['xdg-open', str(report_file)])
                except:
                    pass
                
                # Proposer de pusher sur GitHub
                config_file = Path("config").glob("config_*_*.json")
                config_file = sorted(config_file, key=lambda x: x.stat().st_mtime, reverse=True)[0] if config_file else None
                if config_file:
                    self.push_to_github(config_file)
            
            print("\n" + "="*70)
            print("✅ Processus terminé !")
            print("="*70)
        
        except Exception as e:
            print(f"\n❌ Erreur : {e}")
            sys.exit(1)


if __name__ == "__main__":
    generator = TestConfigGenerator()
    generator.run()
