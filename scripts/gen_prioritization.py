#!/usr/bin/env python3
"""
gen_prioritization.py — Generate test prioritization matrix (MoSCoW + Risk scoring)

Usage:
    python gen_prioritization.py config.json
    python gen_prioritization.py config.json --output prioritization.xlsx

Output:
    - prioritization_{project_slug}_{timestamp}.xlsx
    - Sheets: Matrice MoSCoW, Scoring des risques, Plan d'exécution, Dashboard priorisation
"""

import json
import argparse
import sys
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Tuple
import re

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.chart import BarChart, Reference
    OPENPYXL_AVAILABLE = True
except ImportError:
    print("⚠ openpyxl not installed. Install with: pip install openpyxl")
    OPENPYXL_AVAILABLE = False


class PrioritizationGenerator:
    """Generate test prioritization matrix (MoSCoW + Risk)."""
    
    COLORS = {
        "header": "003366",
        "must": "FFE0E0",        # Red - Must have
        "should": "FFE4CC",      # Orange - Should have
        "could": "FFFACD",       # Yellow - Could have
        "wont": "E2EFDA",        # Green - Won't have
        "critical_risk": "FF0000",
        "high_risk": "FFA500",
        "medium_risk": "FFFF00",
        "low_risk": "90EE90"
    }
    
    def __init__(self, config_path: str):
        """Initialize with config file."""
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = json.load(f)
        
        self.project_name = self.config.get("project", {}).get("name", "Project")
        self.slug = re.sub(r'[^\w]+', '_', self.project_name.lower())[:20]
        
        # Extract data
        if self.config.get("deliverable") == "plan-test":
            self.scope = self.config.get("testPlan", {}).get("scope", "")
            self.risks_text = self.config.get("testPlan", {}).get("risks", "")
        else:
            self.scope = ""
            self.risks_text = ""
    
    def export_to_excel(self, output_path: str = None) -> str:
        """Export prioritization matrix to Excel."""
        if not OPENPYXL_AVAILABLE:
            print("ERROR: openpyxl not available")
            return None
        
        if output_path is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = f"prioritization_{self.slug}_{timestamp}.xlsx"
        
        wb = Workbook()
        ws_moscow = wb.active
        ws_moscow.title = "Matrice MoSCoW"
        
        # Sheet 1: MoSCoW Matrix
        self._write_moscow_matrix(ws_moscow)
        
        # Sheet 2: Risk scoring
        ws_risk = wb.create_sheet("Scoring des risques")
        self._write_risk_scoring(ws_risk)
        
        # Sheet 3: Execution plan
        ws_plan = wb.create_sheet("Plan d'exécution")
        self._write_execution_plan(ws_plan)
        
        # Sheet 4: Prioritization dashboard
        ws_dashboard = wb.create_sheet("Dashboard")
        self._write_dashboard(ws_dashboard)
        
        wb.save(output_path)
        return output_path
    
    def _write_moscow_matrix(self, ws):
        """Write MoSCoW prioritization matrix."""
        # Title
        ws.merge_cells('A1:E1')
        title = ws.cell(row=1, column=1, value="MATRICE DE PRIORISATION — MoSCoW")
        title.font = Font(bold=True, size=12, color="FFFFFF")
        title.fill = PatternFill(start_color=self.COLORS["header"], end_color=self.COLORS["header"], fill_type="solid")
        title.alignment = Alignment(horizontal="center")
        
        # Headers
        headers = ["Cas de test", "Fonctionnalité", "Catégorie MoSCoW", "Score risque", "Priorité d'exécution"]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color=self.COLORS["header"], end_color=self.COLORS["header"], fill_type="solid")
        
        # Test cases (simulated data)
        test_data = [
            ("TC-001", "Authentification", "Must", "Critique", "1"),
            ("TC-002", "Affichage données", "Must", "Haute", "2"),
            ("TC-003", "Recherche médicaments", "Should", "Haute", "3"),
            ("TC-004", "Export de fiche", "Should", "Moyenne", "4"),
            ("TC-005", "Validation limites", "Could", "Moyenne", "5"),
            ("TC-006", "Performance", "Could", "Basse", "6"),
            ("TC-007", "Sécurité avancée", "Won't", "Basse", "Phase 2"),
        ]
        
        for row_idx, (tc_id, func, moscow, risk, priority) in enumerate(test_data, 4):
            ws.cell(row=row_idx, column=1, value=tc_id)
            ws.cell(row=row_idx, column=2, value=func)
            
            moscow_cell = ws.cell(row=row_idx, column=3, value=moscow)
            risk_cell = ws.cell(row=row_idx, column=4, value=risk)
            priority_cell = ws.cell(row=row_idx, column=5, value=priority)
            
            # Color by MoSCoW
            color = {
                "Must": self.COLORS["must"],
                "Should": self.COLORS["should"],
                "Could": self.COLORS["could"],
                "Won't": self.COLORS["wont"]
            }.get(moscow, "FFFFFF")
            
            moscow_cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        
        for col in range(1, 6):
            ws.column_dimensions[get_column_letter(col)].width = 22
    
    def _write_risk_scoring(self, ws):
        """Write risk scoring table."""
        ws.cell(row=1, column=1, value="Score des risques").font = Font(bold=True, size=12)
        
        ws.cell(row=3, column=1, value="Fonctionnalité").font = Font(bold=True, color="FFFFFF")
        ws.cell(row=3, column=2, value="Impact").font = Font(bold=True, color="FFFFFF")
        ws.cell(row=3, column=3, value="Probabilité").font = Font(bold=True, color="FFFFFF")
        ws.cell(row=3, column=4, value="Score risque").font = Font(bold=True, color="FFFFFF")
        ws.cell(row=3, column=5, value="Catégorie").font = Font(bold=True, color="FFFFFF")
        
        for col in range(1, 6):
            ws.cell(row=3, column=col).fill = PatternFill(start_color=self.COLORS["header"], end_color=self.COLORS["header"], fill_type="solid")
        
        risk_data = [
            ("Authentification", "Critique", "Haute", 9, "Critique"),
            ("Affichage données", "Critique", "Moyenne", 6, "Haute"),
            ("Recherche", "Haute", "Moyenne", 6, "Haute"),
            ("Export", "Moyenne", "Basse", 2, "Basse"),
            ("Performance", "Haute", "Moyenne", 6, "Haute"),
        ]
        
        for row_idx, (func, impact, prob, score, category) in enumerate(risk_data, 4):
            ws.cell(row=row_idx, column=1, value=func)
            ws.cell(row=row_idx, column=2, value=impact)
            ws.cell(row=row_idx, column=3, value=prob)
            
            score_cell = ws.cell(row=row_idx, column=4, value=score)
            cat_cell = ws.cell(row=row_idx, column=5, value=category)
            
            # Color by risk score
            risk_color = {
                "Critique": self.COLORS["critical_risk"],
                "Haute": self.COLORS["high_risk"],
                "Moyenne": self.COLORS["medium_risk"],
                "Basse": self.COLORS["low_risk"]
            }.get(category, "FFFFFF")
            
            score_cell.fill = PatternFill(start_color=risk_color, end_color=risk_color, fill_type="solid")
            cat_cell.fill = PatternFill(start_color=risk_color, end_color=risk_color, fill_type="solid")
        
        for col in range(1, 6):
            ws.column_dimensions[get_column_letter(col)].width = 20
    
    def _write_execution_plan(self, ws):
        """Write test execution sequence plan."""
        ws.cell(row=1, column=1, value="PLAN D'EXÉCUTION DES TESTS").font = Font(bold=True, size=12)
        
        ws.cell(row=3, column=1, value="Phase").font = Font(bold=True, color="FFFFFF")
        ws.cell(row=3, column=2, value="Cas de test").font = Font(bold=True, color="FFFFFF")
        ws.cell(row=3, column=3, value="Durée est.").font = Font(bold=True, color="FFFFFF")
        ws.cell(row=3, column=4, value="Dépendances").font = Font(bold=True, color="FFFFFF")
        ws.cell(row=3, column=5, value="Assigné à").font = Font(bold=True, color="FFFFFF")
        
        for col in range(1, 6):
            ws.cell(row=3, column=col).fill = PatternFill(start_color=self.COLORS["header"], end_color=self.COLORS["header"], fill_type="solid")
        
        execution_data = [
            ("Phase 1 (J1)", "TC-001, TC-002", "4h", "Aucune", "QA Team"),
            ("Phase 1 (J1)", "TC-003, TC-004", "4h", "TC-001, TC-002", "QA Team"),
            ("Phase 2 (J2)", "TC-005, TC-006", "6h", "Phase 1", "QA + Perf"),
            ("Phase 2 (J2)", "TC-007", "TBD", "Phase 1", "Sécurité"),
        ]
        
        for row_idx, (phase, tests, duration, deps, assignee) in enumerate(execution_data, 4):
            ws.cell(row=row_idx, column=1, value=phase)
            ws.cell(row=row_idx, column=2, value=tests)
            ws.cell(row=row_idx, column=3, value=duration)
            ws.cell(row=row_idx, column=4, value=deps)
            ws.cell(row=row_idx, column=5, value=assignee)
        
        for col in range(1, 6):
            ws.column_dimensions[get_column_letter(col)].width = 20
    
    def _write_dashboard(self, ws):
        """Write prioritization dashboard."""
        ws.cell(row=1, column=1, value="DASHBOARD DE PRIORISATION").font = Font(bold=True, size=12)
        
        # MoSCoW breakdown
        ws.cell(row=3, column=1, value="Répartition MoSCoW").font = Font(bold=True)
        ws.cell(row=4, column=1, value="Must")
        ws.cell(row=4, column=2, value=2)
        ws.cell(row=4, column=2).fill = PatternFill(start_color=self.COLORS["must"], end_color=self.COLORS["must"], fill_type="solid")
        
        ws.cell(row=5, column=1, value="Should")
        ws.cell(row=5, column=2, value=2)
        ws.cell(row=5, column=2).fill = PatternFill(start_color=self.COLORS["should"], end_color=self.COLORS["should"], fill_type="solid")
        
        ws.cell(row=6, column=1, value="Could")
        ws.cell(row=6, column=2, value=2)
        ws.cell(row=6, column=2).fill = PatternFill(start_color=self.COLORS["could"], end_color=self.COLORS["could"], fill_type="solid")
        
        ws.cell(row=7, column=1, value="Won't")
        ws.cell(row=7, column=2, value=1)
        ws.cell(row=7, column=2).fill = PatternFill(start_color=self.COLORS["wont"], end_color=self.COLORS["wont"], fill_type="solid")
        
        # Risk breakdown
        ws.cell(row=9, column=1, value="Répartition par risque").font = Font(bold=True)
        ws.cell(row=10, column=1, value="Critique")
        ws.cell(row=10, column=2, value=1)
        ws.cell(row=10, column=2).fill = PatternFill(start_color=self.COLORS["critical_risk"], end_color=self.COLORS["critical_risk"], fill_type="solid")
        
        ws.cell(row=11, column=1, value="Haute")
        ws.cell(row=11, column=2, value=3)
        ws.cell(row=11, column=2).fill = PatternFill(start_color=self.COLORS["high_risk"], end_color=self.COLORS["high_risk"], fill_type="solid")
        
        ws.cell(row=12, column=1, value="Moyenne")
        ws.cell(row=12, column=2, value=2)
        ws.cell(row=12, column=2).fill = PatternFill(start_color=self.COLORS["medium_risk"], end_color=self.COLORS["medium_risk"], fill_type="solid")
        
        ws.cell(row=13, column=1, value="Basse")
        ws.cell(row=13, column=2, value=1)
        ws.cell(row=13, column=2).fill = PatternFill(start_color=self.COLORS["low_risk"], end_color=self.COLORS["low_risk"], fill_type="solid")
        
        # Summary metrics
        ws.cell(row=15, column=1, value="Métriques").font = Font(bold=True)
        ws.cell(row=16, column=1, value="Total cas de test")
        ws.cell(row=16, column=2, value=7)
        ws.cell(row=17, column=1, value="% Priorisation critique/haute")
        ws.cell(row=17, column=2, value="57%")
        ws.cell(row=18, column=1, value="Timeline estimée")
        ws.cell(row=18, column=2, value="2 jours")
        
        for col in range(1, 3):
            ws.column_dimensions[get_column_letter(col)].width = 25


def main():
    parser = argparse.ArgumentParser(description="Generate test prioritization matrix")
    parser.add_argument("config", help="Path to config.json file")
    parser.add_argument("-o", "--output", help="Output Excel path (optional)")
    
    args = parser.parse_args()
    
    if not Path(args.config).exists():
        print(f"ERROR: Config file not found: {args.config}")
        sys.exit(1)
    
    print(f"🎯 Reading config: {args.config}")
    generator = PrioritizationGenerator(args.config)
    
    print(f"📊 Generating prioritization matrix for: {generator.project_name}")
    excel_path = generator.export_to_excel(args.output)
    
    if excel_path:
        print(f"✅ Prioritization exported to: {excel_path}")
        return 0
    else:
        print("❌ Export failed")
        return 1


if __name__ == "__main__":
    sys.exit(main())
