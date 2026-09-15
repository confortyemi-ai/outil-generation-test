#!/usr/bin/env python3
"""
gen_test_cases.py — Generate test cases from config.json to Excel (ISO/IEC 29119)

Usage:
    python gen_test_cases.py config_commia_example.json
    python gen_test_cases.py config.json --output my_test_cases.xlsx

Output:
    - test_cases_{project_slug}_{timestamp}.xlsx
    - Columns: ISO/IEC 29119 standard fields
    - Sheets: Test Cases, Coverage Summary, Dashboard
"""

import json
import argparse
import sys
from pathlib import Path
from datetime import datetime
from typing import Dict, List
import re

# Try openpyxl; fallback to csv if not available
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    OPENPYXL_AVAILABLE = True
except ImportError:
    print("⚠ openpyxl not installed. Install with: pip install openpyxl")
    OPENPYXL_AVAILABLE = False


class TestCaseGenerator:
    """Generate test cases from config JSON to Excel."""
    
    # ISO/IEC 29119 standard columns
    COLUMNS_ISO = [
        "ID du cas de test",
        "Titre",
        "Module / Fonctionnalité",
        "User Story / Exigence",
        "Type de test",
        "Niveau de test",
        "Priorité",
        "Préconditions",
        "Données de test",
        "Étape 1",
        "Étape 2",
        "Étape 3",
        "Résultat attendu",
        "Résultat obtenu",
        "Statut",
        "Commentaires",
        "Auteur",
        "Date de création",
        "Version"
    ]
    
    # Color scheme (hex)
    COLORS = {
        "header": "003366",      # Dark blue
        "critical": "FFE0E0",     # Light red
        "high": "FFE4CC",         # Light orange
        "medium": "FFFACD",       # Light yellow
        "low": "E2EFDA",          # Light green
        "pass": "C6EFCE",         # Green
        "fail": "FFC7CE"          # Red
    }
    
    def __init__(self, config_path: str):
        """Initialize generator with config file."""
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = json.load(f)
        
        self.project_name = self.config.get("project", {}).get("name", "Project")
        self.slug = re.sub(r'[^\w]+', '_', self.project_name.lower())[:20]
        self.user_story = self.config.get("testCases", {}).get("userStory", "Unknown")
        self.priority = self.config.get("testCases", {}).get("priority", "moyenne")
        self.test_data = self.config.get("testCases", {}).get("testData", "")
        self.execution_type = self.config.get("testCases", {}).get("executionType", "manuel")
    
    def generate_test_cases(self) -> List[Dict]:
        """Generate sample test cases based on config."""
        cases = []
        
        # Determine test case types to cover
        types_to_cover = self.config.get("testCases", {}).get("typesCases", ["nominaux", "limites", "erreur"])
        
        if "nominaux" in types_to_cover:
            cases.extend(self._generate_nominal_cases())
        if "limites" in types_to_cover:
            cases.extend(self._generate_boundary_cases())
        if "erreur" in types_to_cover:
            cases.extend(self._generate_error_cases())
        
        return cases
    
    def _generate_nominal_cases(self) -> List[Dict]:
        """Generate nominal (happy path) test cases."""
        return [
            {
                "id": f"TC_{self.slug}_NOM_001",
                "title": f"Exécution nominale — {self.user_story}",
                "module": "Cas nominal",
                "user_story": self.user_story,
                "test_type": "Fonctionnel",
                "test_level": "Système",
                "priority": self.priority,
                "preconditions": "Données de test valides disponibles; système en état nominal",
                "test_data": self.test_data if self.test_data else "Données standard",
                "steps": [
                    "1. Accéder au système",
                    "2. Charger les données de test",
                    "3. Exécuter le scénario principal",
                    "4. Vérifier le résultat"
                ],
                "expected": "Pas d'erreur; résultat conforme aux attentes",
                "obtained": "À remplir lors de l'exécution",
                "status": "Non exécuté",
                "comments": "Cas de test fondamental",
                "author": "Test Generator",
                "created": datetime.now().strftime("%Y-%m-%d"),
                "version": "1.0"
            }
        ]
    
    def _generate_boundary_cases(self) -> List[Dict]:
        """Generate boundary (limite) test cases."""
        return [
            {
                "id": f"TC_{self.slug}_LIM_001",
                "title": "Valeur limite inférieure",
                "module": "Cas limite",
                "user_story": self.user_story,
                "test_type": "Fonctionnel",
                "test_level": "Système",
                "priority": "Haute" if self.priority == "Critique" else "Moyenne",
                "preconditions": "Données limite disponibles",
                "test_data": "Valeur minimale acceptée (ex: montant=0.01)",
                "steps": [
                    "1. Utiliser la valeur limite inférieure",
                    "2. Exécuter le scénario",
                    "3. Vérifier le traitement"
                ],
                "expected": "Système accepte et traite la valeur limite",
                "obtained": "À remplir lors de l'exécution",
                "status": "Non exécuté",
                "comments": "Teste la boundary condition basse",
                "author": "Test Generator",
                "created": datetime.now().strftime("%Y-%m-%d"),
                "version": "1.0"
            },
            {
                "id": f"TC_{self.slug}_LIM_002",
                "title": "Valeur limite supérieure",
                "module": "Cas limite",
                "user_story": self.user_story,
                "test_type": "Fonctionnel",
                "test_level": "Système",
                "priority": "Haute" if self.priority == "Critique" else "Moyenne",
                "preconditions": "Données limite disponibles",
                "test_data": "Valeur maximale acceptée",
                "steps": [
                    "1. Utiliser la valeur limite supérieure",
                    "2. Exécuter le scénario",
                    "3. Vérifier le traitement"
                ],
                "expected": "Système accepte et traite la valeur limite",
                "obtained": "À remplir lors de l'exécution",
                "status": "Non exécuté",
                "comments": "Teste la boundary condition haute",
                "author": "Test Generator",
                "created": datetime.now().strftime("%Y-%m-%d"),
                "version": "1.0"
            }
        ]
    
    def _generate_error_cases(self) -> List[Dict]:
        """Generate error handling test cases."""
        return [
            {
                "id": f"TC_{self.slug}_ERR_001",
                "title": "Gestion des données manquantes",
                "module": "Gestion des erreurs",
                "user_story": self.user_story,
                "test_type": "Fonctionnel",
                "test_level": "Système",
                "priority": "Critique",
                "preconditions": "Données incomplètes préparées",
                "test_data": "Champ(s) obligatoire(s) manquant(s)",
                "steps": [
                    "1. Soumettre les données incomplètes",
                    "2. Observer le comportement du système",
                    "3. Vérifier le message d'erreur"
                ],
                "expected": "Erreur explicite; pas de génération fictive de données",
                "obtained": "À remplir lors de l'exécution",
                "status": "Non exécuté",
                "comments": "Critique pour les systèmes IA (anti-hallucination)",
                "author": "Test Generator",
                "created": datetime.now().strftime("%Y-%m-%d"),
                "version": "1.0"
            },
            {
                "id": f"TC_{self.slug}_ERR_002",
                "title": "Gestion des valeurs invalides",
                "module": "Gestion des erreurs",
                "user_story": self.user_story,
                "test_type": "Fonctionnel",
                "test_level": "Système",
                "priority": "Haute",
                "preconditions": "Données invalides préparées",
                "test_data": "Valeur hors plage acceptable (ex: montant=-100)",
                "steps": [
                    "1. Soumettre la valeur invalide",
                    "2. Observer la réponse du système",
                    "3. Vérifier le rejet"
                ],
                "expected": "Rejet avec message d'erreur clair",
                "obtained": "À remplir lors de l'exécution",
                "status": "Non exécuté",
                "comments": "Validez la robustesse",
                "author": "Test Generator",
                "created": datetime.now().strftime("%Y-%m-%d"),
                "version": "1.0"
            }
        ]
    
    def export_to_excel(self, output_path: str = None) -> str:
        """Export test cases to Excel."""
        if not OPENPYXL_AVAILABLE:
            print("ERROR: openpyxl not available. Cannot generate Excel.")
            return None
        
        if output_path is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = f"test_cases_{self.slug}_{timestamp}.xlsx"
        
        wb = Workbook()
        ws_cases = wb.active
        ws_cases.title = "Cas de test"
        
        # Sheet 1: Test Cases
        self._write_test_cases_sheet(ws_cases)
        
        # Sheet 2: Coverage Summary
        ws_summary = wb.create_sheet("Couverture")
        self._write_coverage_sheet(ws_summary)
        
        # Sheet 3: Dashboard
        ws_dashboard = wb.create_sheet("Tableau de bord")
        self._write_dashboard_sheet(ws_dashboard)
        
        wb.save(output_path)
        return output_path
    
    def _write_test_cases_sheet(self, ws):
        """Write test cases to worksheet."""
        cases = self.generate_test_cases()
        
        # Header
        for col, header in enumerate(self.COLUMNS_ISO, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color=self.COLORS["header"], end_color=self.COLORS["header"], fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        # Data rows
        for row_idx, case in enumerate(cases, 2):
            ws.cell(row=row_idx, column=1, value=case["id"])
            ws.cell(row=row_idx, column=2, value=case["title"])
            ws.cell(row=row_idx, column=3, value=case["module"])
            ws.cell(row=row_idx, column=4, value=case["user_story"])
            ws.cell(row=row_idx, column=5, value=case["test_type"])
            ws.cell(row=row_idx, column=6, value=case["test_level"])
            ws.cell(row=row_idx, column=7, value=case["priority"])
            ws.cell(row=row_idx, column=8, value=case["preconditions"])
            ws.cell(row=row_idx, column=9, value=case["test_data"])
            
            # Steps (columns 10-12)
            for step_idx, step in enumerate(case["steps"][:3]):
                ws.cell(row=row_idx, column=10+step_idx, value=step)
            
            ws.cell(row=row_idx, column=13, value=case["expected"])
            ws.cell(row=row_idx, column=14, value=case["obtained"])
            ws.cell(row=row_idx, column=15, value=case["status"])
            ws.cell(row=row_idx, column=16, value=case["comments"])
            ws.cell(row=row_idx, column=17, value=case["author"])
            ws.cell(row=row_idx, column=18, value=case["created"])
            ws.cell(row=row_idx, column=19, value=case["version"])
            
            # Color based on priority
            fill_color = self.COLORS.get(case["priority"].lower(), self.COLORS["medium"])
            for col in range(1, len(self.COLUMNS_ISO) + 1):
                ws.cell(row=row_idx, column=col).fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
                ws.cell(row=row_idx, column=col).alignment = Alignment(wrap_text=True, vertical="top")
        
        # Auto-fit columns
        for col in range(1, len(self.COLUMNS_ISO) + 1):
            ws.column_dimensions[chr(64 + col)].width = 20
        
        # Freeze panes
        ws.freeze_panes = "A2"
        
        # Auto-filter
        ws.auto_filter.ref = f"A1:{chr(64+len(self.COLUMNS_ISO))}{len(cases)+1}"
    
    def _write_coverage_sheet(self, ws):
        """Write coverage summary."""
        cases = self.generate_test_cases()
        
        # Summary by type
        ws.cell(row=1, column=1, value="Type de cas").font = Font(bold=True)
        ws.cell(row=1, column=2, value="Nombre").font = Font(bold=True)
        ws.cell(row=1, column=1).fill = PatternFill(start_color=self.COLORS["header"], end_color=self.COLORS["header"], fill_type="solid")
        ws.cell(row=1, column=2).fill = PatternFill(start_color=self.COLORS["header"], end_color=self.COLORS["header"], fill_type="solid")
        
        types = {}
        for case in cases:
            module = case["module"]
            types[module] = types.get(module, 0) + 1
        
        for row_idx, (type_name, count) in enumerate(types.items(), 2):
            ws.cell(row=row_idx, column=1, value=type_name)
            ws.cell(row=row_idx, column=2, value=count)
        
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 15
    
    def _write_dashboard_sheet(self, ws):
        """Write dashboard with statistics."""
        cases = self.generate_test_cases()
        
        total = len(cases)
        by_priority = {}
        for case in cases:
            priority = case["priority"]
            by_priority[priority] = by_priority.get(priority, 0) + 1
        
        ws.cell(row=1, column=1, value="Tableau de bord").font = Font(bold=True, size=14)
        ws.cell(row=3, column=1, value="Total des cas de test").font = Font(bold=True)
        ws.cell(row=3, column=2, value=total)
        
        row = 5
        ws.cell(row=row, column=1, value="Par priorité").font = Font(bold=True)
        row += 1
        for priority, count in sorted(by_priority.items()):
            ws.cell(row=row, column=1, value=priority)
            ws.cell(row=row, column=2, value=count)
            row += 1


def main():
    parser = argparse.ArgumentParser(description="Generate test cases from config.json to Excel")
    parser.add_argument("config", help="Path to config.json file")
    parser.add_argument("-o", "--output", help="Output Excel path (optional)")
    
    args = parser.parse_args()
    
    if not Path(args.config).exists():
        print(f"ERROR: Config file not found: {args.config}")
        sys.exit(1)
    
    print(f"📋 Reading config: {args.config}")
    generator = TestCaseGenerator(args.config)
    
    print(f"🧪 Generating test cases for: {generator.project_name}")
    excel_path = generator.export_to_excel(args.output)
    
    if excel_path:
        print(f"✅ Test cases exported to: {excel_path}")
        return 0
    else:
        print("❌ Export failed")
        return 1


if __name__ == "__main__":
    sys.exit(main())
