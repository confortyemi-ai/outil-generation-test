#!/usr/bin/env python3
"""
api_jira_wrapper.py — Wrapper pour Jira API

Récupère :
- Les Epics du projet
- Les User Stories d'un Epic
- Les détails d'une Story
"""

import os
import requests
import json
from typing import List, Dict
from dotenv import load_dotenv

load_dotenv()

class JiraAPI:
    """Client pour Jira API"""
    
    def __init__(self):
        self.url = os.getenv('JIRA_URL')
        self.email = os.getenv('JIRA_EMAIL')
        self.token = os.getenv('JIRA_API_TOKEN')
        
        if not all([self.url, self.email, self.token]):
            raise ValueError("❌ JIRA_URL, JIRA_EMAIL ou JIRA_API_TOKEN manquant dans .env")
        
        self.headers = {
            'Authorization': f'Basic {self._encode_auth()}',
            'Content-Type': 'application/json'
        }
    
    def _encode_auth(self) -> str:
        """Encoder email:token en base64"""
        import base64
        creds = f"{self.email}:{self.token}"
        return base64.b64encode(creds.encode()).decode()
    
    def search_epics(self, project_key: str) -> List[Dict]:
        """Récupérer les Epics d'un projet"""
        # Pour Jira Cloud, utiliser issuetype = Epic
        jql = f'project = "{project_key}" AND issuetype = Epic ORDER BY created DESC'
        
        response = requests.get(
            f'{self.url}/rest/api/3/search',
            headers=self.headers,
            params={
                'jql': jql,
                'maxResults': 100,
                'expand': 'changelog'
            }
        )
        
        if response.status_code != 200:
            print(f"❌ Erreur Jira: {response.status_code}")
            print(f"📌 Message: {response.text}")
            return []
        
        issues = response.json().get('issues', [])
        return [
            {
                'key': issue['key'],
                'summary': issue['fields']['summary'],
                'name': issue['fields']['summary']
            }
            for issue in issues
        ]
    
    def search_stories_in_epic(self, epic_key: str) -> List[Dict]:
        """Récupérer les User Stories d'un Epic"""
        # Pour Jira Cloud, utiliser "Epic Link" = epic_key
        jql = f'"Epic Link" = {epic_key} AND issuetype = Story ORDER BY created DESC'
        
        response = requests.get(
            f'{self.url}/rest/api/3/search',
            headers=self.headers,
            params={
                'jql': jql,
                'maxResults': 100,
                'expand': 'changelog'
            }
        )
        
        if response.status_code != 200:
            print(f"❌ Erreur Jira: {response.status_code}")
            print(f"📌 Message: {response.text}")
            return []
        
        issues = response.json().get('issues', [])
        return [
            {
                'key': issue['key'],
                'summary': issue['fields']['summary'],
                'description': issue['fields'].get('description', {}).get('content', [{}])[0].get('content', [{}])[0].get('text', '') if issue['fields'].get('description') else ''
            }
            for issue in issues
        ]
    
    def get_story_details(self, story_key: str) -> Dict:
        """Récupérer les détails d'une Story"""
        response = requests.get(
            f'{self.url}/rest/api/3/issues/{story_key}',
            headers=self.headers
        )
        
        if response.status_code != 200:
            print(f"❌ Erreur Jira: {response.status_code}")
            return {}
        
        issue = response.json()
        
        # Extraire la description
        description = ""
        if issue['fields'].get('description'):
            for block in issue['fields']['description'].get('content', []):
                if block['type'] == 'paragraph':
                    for content in block.get('content', []):
                        if content['type'] == 'text':
                            description += content.get('text', '')
        
        return {
            'key': issue['key'],
            'summary': issue['fields']['summary'],
            'description': description,
            'priority': issue['fields'].get('priority', {}).get('name', 'Medium'),
            'assignee': issue['fields'].get('assignee', {}).get('displayName', 'Unassigned'),
            'labels': issue['fields'].get('labels', [])
        }
    
    def get_epic_by_story(self, story_key: str) -> Dict:
        """Récupérer l'Epic parent d'une Story"""
        response = requests.get(
            f'{self.url}/rest/api/3/issues/{story_key}',
            headers=self.headers
        )
        
        if response.status_code != 200:
            return {}
        
        issue = response.json()
        
        # Chercher le lien "Epic Link" ou "Parent Epic"
        epic_link = None
        for field_name, field_value in issue['fields'].items():
            if 'epic' in field_name.lower() or field_name == 'customfield_10014':
                if isinstance(field_value, dict) and 'key' in field_value:
                    epic_link = field_value['key']
                    break
        
        if not epic_link:
            return {}
        
        # Récupérer les détails de l'Epic
        epic_response = requests.get(
            f'{self.url}/rest/api/3/issues/{epic_link}',
            headers=self.headers
        )
        
        if epic_response.status_code != 200:
            return {}
        
        epic = epic_response.json()
        return {
            'key': epic['key'],
            'summary': epic['fields']['summary'],
            'name': epic['fields']['summary']
        }


def main():
    """Test interactif"""
    try:
        jira = JiraAPI()
        print("✅ Connexion Jira réussie!\n")
        
        # Demander le projet
        project_key = input("Quel est votre projet Jira (ex: XSP)? ").strip().upper()
        
        # Récupérer les Epics
        print(f"\n⏳ Récupération des Epics du projet {project_key}...")
        epics = jira.search_epics(project_key)
        
        if not epics:
            print("❌ Aucun Epic trouvé")
            return
        
        print(f"\n✅ {len(epics)} Epic(s) trouvé(s) :\n")
        for i, epic in enumerate(epics, 1):
            print(f"  {i}. {epic['key']} — {epic['summary']}")
        
        # Choisir un Epic
        choice = int(input("\nQuel Epic (numéro)? ")) - 1
        selected_epic = epics[choice]
        
        # Récupérer les Stories
        print(f"\n⏳ Récupération des Stories de {selected_epic['key']}...")
        stories = jira.search_stories_in_epic(selected_epic['key'])
        
        if not stories:
            print("❌ Aucune Story trouvée")
            return
        
        print(f"\n✅ {len(stories)} Story(ies) trouvée(s) :\n")
        for i, story in enumerate(stories, 1):
            print(f"  {i}. {story['key']} — {story['summary']}")
        
        # Choisir une Story
        choice = int(input("\nQuelle Story (numéro)? ")) - 1
        selected_story = stories[choice]
        
        # Détails
        print(f"\n⏳ Récupération des détails de {selected_story['key']}...")
        details = jira.get_story_details(selected_story['key'])
        
        print(f"\n✅ Détails :")
        print(json.dumps(details, indent=2, ensure_ascii=False))
    
    except Exception as e:
        print(f"❌ Erreur : {e}")


if __name__ == "__main__":
    main()
