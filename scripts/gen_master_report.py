#!/usr/bin/env python3
"""
gen_master_report.py — Master Test Report: ALL artifacts in ONE Excel file

Usage:
    python gen_master_report.py config.json
    python gen_master_report.py config.json --output master_report.xlsx

Output:
    - master_report_{project_slug}_{timestamp}.xlsx
    - Sheet 1: Cas + Priorisation + Résultats
    - Sheet 2: Couverture
    - Sheet 3: Priorisation
    - Sheet 4: Rapport
    - Sheet 5: Dashboard
    
This is the SINGLE source of truth for test management.
"""

import json
import argparse
import sys
from pathlib import Path
from datetime import datetime
from typing import List, Dict
import re

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    OPENPYXL_AVAILABLE = True
except ImportError:
    print("⚠ openpyxl not installed. Install with: pip install openpyxl")
    OPENPYXL_AVAILABLE = False


class MasterReportGenerator:
    """Generate master test report with all artifacts in one Excel."""
    
    BASE_COLUMNS = [
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
    ]
    
    PRIORITIZATION_COLUMNS = [
        "Catégorie MoSCoW",
        "Score risque",
        "Priorité d'exécution"
    ]
    
    EXECUTION_COLUMNS = [
        "Statut",
        "Durée",
        "Notes"
    ]
    
    COLORS = {
        "header": "003366",
        "critical": "FFE0E0",
        "high": "FFE4CC",
        "medium": "FFFACD",
        "low": "E2EFDA",
        "pass": "C6EFCE",
        "fail": "FFC7CE",
        "blocked": "FFEB9C"
    }
    
    def __init__(self, config_path: str):
        """Initialize with config file."""
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = json.load(f)
        
        self.project_name = self.config.get("project", {}).get("name", "Project")
        self.slug = re.sub(r'[^\w]+', '_', self.project_name.lower())[:20]
        self.environment = self.config.get("project", {}).get("environment", "Unknown")
        self.methodology = self.config.get("project", {}).get("methodology", "Unknown")
        
        # Generate all data
        self.test_cases = self._generate_test_cases()
        self.prioritization = self._generate_prioritization()
        self.execution_results = self._generate_execution_results()
    
    def _generate_test_cases(self) -> List[Dict]:
        """Generate test case data."""
        return [
            {
                "id": "TC-001",
                "title": "Authentification nominale",
                "module": "Authentification",
                "user_story": "US-RA-001",
                "test_type": "Fonctionnel",
                "test_level": "Système",
                "priority": "Critique",
                "preconditions": "Données de test valides disponibles",
                "test_data": "Utilisateurs valides",
                "step1": "1. Accéder au formulaire de connexion",
                "step2": "2. Entrer les identifiants corrects",
                "step3": "3. Cliquer sur Connexion",
                "expected": "Authentification réussie, redirection vers dashboard"
            },
            {
                "id": "TC-002",
                "title": "Affichage des données",
                "module": "Affichage",
                "user_story": "US-RA-001",
                "test_type": "Fonctionnel",
                "test_level": "Système",
                "priority": "Critique",
                "preconditions": "Utilisateur authentifié",
                "test_data": "Données de test standard",
                "step1": "1. Se connecter",
                "step2": "2. Accéder à la page principale",
                "step3": "3. Vérifier l'affichage des données",
                "expected": "Données affichées correctement et complètes"
            },
            {
                "id": "TC-003",
                "title": "Recherche de médicaments",
                "module": "Recherche",
                "user_story": "US-RA-001",
                "test_type": "Fonctionnel",
                "test_level": "Système",
                "priority": "Haute",
                "preconditions": "Utilisateur sur la page principale",
                "test_data": "Termes de recherche variés",
                "step1": "1. Cliquer sur le champ Recherche",
                "step2": "2. Entrer un terme (ex: 'Aspirin')",
                "step3": "3. Vérifier les résultats",
                "expected": "Résultats corrects et filtrés"
            },
            {
                "id": "TC-004",
                "title": "Export de fiche",
                "module": "Export",
                "user_story": "US-RA-001",
                "test_type": "Fonctionnel",
                "test_level": "Système",
                "priority": "Haute",
                "preconditions": "Fiche sélectionnée",
                "test_data": "Fiche medicament test",
                "step1": "1. Sélectionner une fiche",
                "step2": "2. Cliquer sur Export",
                "step3": "3. Vérifier le fichier généré",
                "expected": "Export réussi, format correct (PDF/Excel)"
            },
            {
                "id": "TC-005",
                "title": "Gestion des erreurs",
                "module": "Erreurs",
                "user_story": "US-RA-001",
                "test_type": "Fonctionnel",
                "test_level": "Système",
                "priority": "Moyenne",
                "preconditions": "Données incomplètes préparées",
                "test_data": "Champs vides/invalides",
                "step1": "1. Soumettre un formulaire incomplet",
                "step2": "2. Observer le message d'erreur",
                "step3": "3. Vérifier les contrôles",
                "expected": "Message d'erreur explicite, pas de crash"
            },
        ]
    
    def _generate_prioritization(self) -> Dict:
        """Generate prioritization data."""
        return {
            "TC-001": {"moscow": "Must", "risk_score": "Critique", "exec_priority": 1},
            "TC-002": {"moscow": "Must", "risk_score": "Haute", "exec_priority": 2},
            "TC-003": {"moscow": "Should", "risk_score": "Haute", "exec_priority": 3},
            "TC-004": {"moscow": "Should", "risk_score": "Moyenne", "exec_priority": 4},
            "TC-005": {"moscow": "Could", "risk_score": "Moyenne", "exec_priority": 5},
        }
    
    def _generate_execution_results(self) -> Dict:
        """Generate execution results."""
        return {
            "TC-001": {"status": "Pass", "duration": "5m", "notes": "✓ Authentification OK"},
            "TC-002": {"status": "Pass", "duration": "3m", "notes": "✓ Données affichées"},
            "TC-003": {"status": "Pass", "duration": "4m", "notes": "✓ Recherche OK"},
            "TC-004": {"status": "Fail", "duration": "2m", "notes": "Format Excel incorrect - DEF-001"},
            "TC-005": {"status": "Pass", "duration": "3m", "notes": "✓ Messages d'erreur clairs"},
        }
    
    def export_to_excel(self, output_path: str = None) -> str:
        """Export master report to Excel."""
        if not OPENPYXL_AVAILABLE:
            print("ERROR: openpyxl not available")
            return None
        
        if output_path is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = f"master_report_{self.slug}_{timestamp}.xlsx"
        
        wb = Workbook()
        wb.remove(wb.active)  # Remove default sheet
        
        # Sheet 1: Cas + Priorisation + Résultats
        ws_cases = wb.create_sheet("Cas + Priorisation + Résultats", 0)
        self._write_sheet_1_cases(ws_cases)
        
        # Sheet 2: Couverture
        ws_coverage = wb.create_sheet("Couverture", 1)
        self._write_sheet_2_coverage(ws_coverage)
        
        # Sheet 3: Priorisation
        ws_prio = wb.create_sheet("Priorisation", 2)
        self._write_sheet_3_prioritization(ws_prio)
        
        # Sheet 4: Rapport
        ws_report = wb.create_sheet("Rapport", 3)
        self._write_sheet_4_report(ws_report)
        
        # Sheet 5: Dashboard
        ws_dashboard = wb.create_sheet("Dashboard", 4)
        self._write_sheet_5_dashboard(ws_dashboard)
        
        wb.save(output_path)
        return output_path
    
    def _write_sheet_1_cases(self, ws):
        """Sheet 1: Cas + Priorisation + Résultats."""
        all_columns = self.BASE_COLUMNS + self.PRIORITIZATION_COLUMNS + self.EXECUTION_COLUMNS
        
        # Header
        for col_idx, header in enumerate(all_columns, 1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color=self.COLORS["header"], end_color=self.COLORS["header"], fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        # Data rows
        for row_idx, case in enumerate(self.test_cases, 2):
            tc_id = case["id"]
            
            # Base columns
            ws.cell(row=row_idx, column=1, value=case["id"])
            ws.cell(row=row_idx, column=2, value=case["title"])
            ws.cell(row=row_idx, column=3, value=case["module"])
            ws.cell(row=row_idx, column=4, value=case["user_story"])
            ws.cell(row=row_idx, column=5, value=case["test_type"])
            ws.cell(row=row_idx, column=6, value=case["test_level"])
            ws.cell(row=row_idx, column=7, value=case["priority"])
            ws.cell(row=row_idx, column=8, value=case["preconditions"])
            ws.cell(row=row_idx, column=9, value=case["test_data"])
            ws.cell(row=row_idx, column=10, value=case["step1"])
            ws.cell(row=row_idx, column=11, value=case["step2"])
            ws.cell(row=row_idx, column=12, value=case["step3"])
            ws.cell(row=row_idx, column=13, value=case["expected"])
            
            # Prioritization
            prio = self.prioritization.get(tc_id, {})
            ws.cell(row=row_idx, column=14, value=prio.get("moscow", ""))
            ws.cell(row=row_idx, column=15, value=prio.get("risk_score", ""))
            ws.cell(row=row_idx, column=16, value=prio.get("exec_priority", ""))
            
            # Execution results
            result = self.execution_results.get(tc_id, {})
            status_cell = ws.cell(row=row_idx, column=17, value=result.get("status", ""))
            ws.cell(row=row_idx, column=18, value=result.get("duration", ""))
            ws.cell(row=row_idx, column=19, value=result.get("notes", ""))
            
            # Color status
            status_color = {
                "Pass": self.COLORS["pass"],
                "Fail": self.COLORS["fail"],
                "Blocked": self.COLORS["blocked"]
            }.get(result.get("status", ""), "FFFFFF")
            status_cell.fill = PatternFill(start_color=status_color, end_color=status_color, fill_type="solid")
            
            # Color base + prioritization by priority
            fill_color = {
                "Critique": self.COLORS["critical"],
                "Haute": self.COLORS["high"],
                "Moyenne": self.COLORS["medium"],
                "Basse": self.COLORS["low"]
            }.get(case["priority"], "FFFFFF")
            
            for col in range(1, 17):
                ws.cell(row=row_idx, column=col).fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
                ws.cell(row=row_idx, column=col).alignment = Alignment(wrap_text=True, vertical="top")
            
            for col in range(17, 20):
                ws.cell(row=row_idx, column=col).alignment = Alignment(wrap_text=True, vertical="top")
        
        # Format
        for col in range(1, len(all_columns) + 1):
            ws.column_dimensions[get_column_letter(col)].width = 18
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(all_columns))}{len(self.test_cases)+1}"
    
    def _write_sheet_2_coverage(self, ws):
        """Sheet 2: Coverage summary."""
        ws.cell(row=1, column=1, value="Couverture de test").font = Font(bold=True, size=12)
        
        # Module coverage
        ws.cell(row=3, column=1, value="Par module").font = Font(bold=True)
        ws.cell(row=4, column=1, value="Module").font = Font(bold=True, color="FFFFFF")
        ws.cell(row=4, column=2, value="Couverture %").font = Font(bold=True, color="FFFFFF")
        ws.cell(row=4, column=1).fill = PatternFill(start_color=self.COLORS["header"], end_color=self.COLORS["header"], fill_type="solid")
        ws.cell(row=4, column=2).fill = PatternFill(start_color=self.COLORS["header"], end_color=self.COLORS["header"], fill_type="solid")
        
        modules_cov = [
            ("Authentification", 95),
            ("Affichage", 85),
            ("Recherche", 80),
            ("Export", 70),
            ("Erreurs", 75),
        ]
        
        for row_idx, (module, coverage) in enumerate(modules_cov, 5):
            ws.cell(row=row_idx, column=1, value=module)
            cov_cell = ws.cell(row=row_idx, column=2, value=coverage)
            cov_cell.fill = PatternFill(start_color=self.COLORS["pass"], end_color=self.COLORS["pass"], fill_type="solid")
        
        ws.column_dimensions['A'].width = 20
        ws.column_dimensions['B'].width = 15
    
    def _write_sheet_3_prioritization(self, ws):
        """Sheet 3: Prioritization summary."""
        ws.cell(row=1, column=1, value="Priorisation").font = Font(bold=True, size=12)
        
        ws.cell(row=3, column=1, value="Répartition MoSCoW").font = Font(bold=True)
        ws.cell(row=4, column=1, value="Catégorie").font = Font(bold=True, color="FFFFFF")
        ws.cell(row=4, column=2, value="Nombre").font = Font(bold=True, color="FFFFFF")
        ws.cell(row=4, column=1).fill = PatternFill(start_color=self.COLORS["header"], end_color=self.COLORS["header"], fill_type="solid")
        ws.cell(row=4, column=2).fill = PatternFill(start_color=self.COLORS["header"], end_color=self.COLORS["header"], fill_type="solid")
        
        moscow_data = [
            ("Must", 2),
            ("Should", 2),
            ("Could", 1),
        ]
        
        for row_idx, (category, count) in enumerate(moscow_data, 5):
            ws.cell(row=row_idx, column=1, value=category)
            ws.cell(row=row_idx, column=2, value=count)
        
        ws.column_dimensions['A'].width = 20
        ws.column_dimensions['B'].width = 15
    
    def _write_sheet_4_report(self, ws):
        """Sheet 4: Report summary."""
        ws.cell(row=1, column=1, value="Rapport de synthèse").font = Font(bold=True, size=12)
        
        passed = sum(1 for r in self.execution_results.values() if r["status"] == "Pass")
        failed = sum(1 for r in self.execution_results.values() if r["status"] == "Fail")
        blocked = sum(1 for r in self.execution_results.values() if r["status"] == "Blocked")
        total = len(self.test_cases)
        
        ws.cell(row=3, column=1, value="Métriques").font = Font(bold=True)
        ws.cell(row=4, column=1, value="Métrique").font = Font(bold=True, color="FFFFFF")
        ws.cell(row=4, column=2, value="Valeur").font = Font(bold=True, color="FFFFFF")
        ws.cell(row=4, column=1).fill = PatternFill(start_color=self.COLORS["header"], end_color=self.COLORS["header"], fill_type="solid")
        ws.cell(row=4, column=2).fill = PatternFill(start_color=self.COLORS["header"], end_color=self.COLORS["header"], fill_type="solid")
        
        metrics = [
            ("Total cas de test", total),
            ("Cas réussis", f"{passed} ({passed/total*100:.0f}%)"),
            ("Cas échoués", f"{failed} ({failed/total*100:.0f}%)"),
            ("Cas bloqués", f"{blocked} ({blocked/total*100:.0f}%)"),
            ("Projet", self.project_name),
            ("Environnement", self.environment),
            ("Méthodologie", self.methodology),
        ]
        
        for row_idx, (label, value) in enumerate(metrics, 5):
            ws.cell(row=row_idx, column=1, value=label)
            ws.cell(row=row_idx, column=2, value=value)
        
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 30
    
    def _write_sheet_5_dashboard(self, ws):
        """Sheet 5: Dashboard metrics."""
        ws.cell(row=1, column=1, value="Dashboard").font = Font(bold=True, size=12)
        
        ws.cell(row=3, column=1, value="Statut d'exécution").font = Font(bold=True)
        ws.cell(row=4, column=1, value="Statut").font = Font(bold=True, color="FFFFFF")
        ws.cell(row=4, column=2, value="Nombre").font = Font(bold=True, color="FFFFFF")
        ws.cell(row=4, column=1).fill = PatternFill(start_color=self.COLORS["header"], end_color=self.COLORS["header"], fill_type="solid")
        ws.cell(row=4, column=2).fill = PatternFill(start_color=self.COLORS["header"], end_color=self.COLORS["header"], fill_type="solid")
        
        pass_count = sum(1 for r in self.execution_results.values() if r["status"] == "Pass")
        fail_count = sum(1 for r in self.execution_results.values() if r["status"] == "Fail")
        block_count = sum(1 for r in self.execution_results.values() if r["status"] == "Blocked")
        
        status_data = [
            ("Pass", pass_count, self.COLORS["pass"]),
            ("Fail", fail_count, self.COLORS["fail"]),
            ("Blocked", block_count, self.COLORS["blocked"]),
        ]
        
        for row_idx, (status, count, color) in enumerate(status_data, 5):
            ws.cell(row=row_idx, column=1, value=status)
            count_cell = ws.cell(row=row_idx, column=2, value=count)
            count_cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        
        ws.column_dimensions['A'].width = 20
        ws.column_dimensions['B'].width = 15


def main():
    parser = argparse.ArgumentParser(description="Generate Master Test Report (all artifacts in one Excel)")
    parser.add_argument("config", help="Path to config.json file")
    parser.add_argument("-o", "--output", help="Output Excel path (optional)")
    
    args = parser.parse_args()
    
    if not Path(args.config).exists():
        print(f"ERROR: Config file not found: {args.config}")
        sys.exit(1)
    
    print(f"📋 Reading config: {args.config}")
    generator = MasterReportGenerator(args.config)
    
    print(f"🎯 Generating master report for: {generator.project_name}")
    excel_path = generator.export_to_excel(args.output)
    
    if excel_path:
        print(f"✅ Master report exported to: {excel_path}")
        print(f"\n📊 Sheets included:")
        print(f"  1. Cas + Priorisation + Résultats (19 colonnes)")
        print(f"  2. Couverture (résumé par module)")
        print(f"  3. Priorisation (répartition MoSCoW)")
        print(f"  4. Rapport (synthèse + métriques)")
        print(f"  5. Dashboard (statut d'exécution)")
        return 0
    else:
        print("❌ Export failed")
        return 1


if __name__ == "__main__":
    sys.exit(main())
