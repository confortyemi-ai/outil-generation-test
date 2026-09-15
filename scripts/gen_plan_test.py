#!/usr/bin/env python3
"""
gen_plan_test.py — Generate test plans from config.json to Word/Excel (IEEE 829)

Usage:
    python gen_plan_test.py config_regulatory_agent.json
    python gen_plan_test.py config.json --output my_plan.docx

Output:
    - plan_test_{project_slug}_{timestamp}.docx (Word document)
    - plan_test_{project_slug}_{timestamp}.xlsx (Excel version)

Standard: IEEE 829 Test Plan documentation
"""

import json
import argparse
import sys
from pathlib import Path
from datetime import datetime
from typing import Dict
import re

# Try to import libraries
try:
    from docx import Document
    from docx.shared import Pt, Inches, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False
    print("⚠ python-docx not installed. Install with: pip install python-docx")

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False
    print("⚠ openpyxl not installed. Install with: pip install openpyxl")


class TestPlanGenerator:
    """Generate test plans from config JSON (IEEE 829 standard)."""
    
    IEEE829_SECTIONS = [
        "Identifiant du plan de test",
        "Introduction et objectifs",
        "Éléments à tester",
        "Fonctionnalités à tester / à ne pas tester",
        "Approche de test",
        "Critères d'entrée et de sortie",
        "Livrables de test",
        "Besoins en environnement",
        "Responsabilités",
        "Risques et plans de contingence",
        "Calendrier"
    ]
    
    def __init__(self, config_path: str):
        """Initialize generator with config file."""
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = json.load(f)
        
        self.project_name = self.config.get("project", {}).get("name", "Project")
        self.slug = re.sub(r'[^\w]+', '_', self.project_name.lower())[:20]
        self.environment = self.config.get("project", {}).get("environment", "Unknown")
        self.methodology = self.config.get("project", {}).get("methodology", "Unknown")
        self.language = self.config.get("project", {}).get("language", "en")
        self.test_levels = self.config.get("project", {}).get("testLevels", [])
        
        self.test_plan_config = self.config.get("testPlan", {})
        self.scope = self.test_plan_config.get("scope", "À définir")
        self.criteria = self.test_plan_config.get("entryExitCriteria", "À définir")
        self.stakeholders = self.test_plan_config.get("stakeholders", "À définir")
        self.risks = self.test_plan_config.get("risks", "À définir")
        self.timeline = self.test_plan_config.get("timeline", "À définir")
        self.tools = self.test_plan_config.get("tools", "À définir")
    
    def _format_text(self, text: str) -> str:
        """Clean and format text."""
        if not text:
            return "À définir"
        return text.strip()
    
    def export_to_docx(self, output_path: str = None) -> str:
        """Export test plan to Word document (IEEE 829)."""
        if not DOCX_AVAILABLE:
            print("ERROR: python-docx not available. Cannot generate Word document.")
            return None
        
        if output_path is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = f"plan_test_{self.slug}_{timestamp}.docx"
        
        doc = Document()
        
        # Title
        title = doc.add_heading(f"Plan de test — {self.project_name}", level=1)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Metadata
        meta_table = doc.add_table(rows=6, cols=2)
        meta_table.style = 'Light Grid Accent 1'
        meta_table.cell(0, 0).text = "Projet"
        meta_table.cell(0, 1).text = self.project_name
        meta_table.cell(1, 0).text = "Environnement"
        meta_table.cell(1, 1).text = self.environment
        meta_table.cell(2, 0).text = "Méthodologie"
        meta_table.cell(2, 1).text = self.methodology
        meta_table.cell(3, 0).text = "Niveaux de test"
        meta_table.cell(3, 1).text = ", ".join(self.test_levels) if self.test_levels else "Non spécifié"
        meta_table.cell(4, 0).text = "Date de création"
        meta_table.cell(4, 1).text = datetime.now().strftime("%Y-%m-%d %H:%M")
        meta_table.cell(5, 0).text = "Statut"
        meta_table.cell(5, 1).text = "Brouillon"
        
        doc.add_paragraph()
        
        # 1. Identifiant du plan de test
        doc.add_heading("1. Identifiant du plan de test", level=2)
        doc.add_paragraph(f"Plan-{self.slug}-{datetime.now().strftime('%Y%m%d')}")
        
        # 2. Introduction et objectifs
        doc.add_heading("2. Introduction et objectifs", level=2)
        doc.add_paragraph(
            f"Ce plan de test couvre les objectifs suivants :\n\n"
            f"• Tester les fonctionnalités critiques identifiées\n"
            f"• Assurer la couverture de test appropriée\n"
            f"• Documenter les cas de test et résultats\n"
            f"• Valider la conformité aux exigences"
        )
        
        # 3. Éléments à tester
        doc.add_heading("3. Éléments à tester", level=2)
        doc.add_paragraph("Environnement :", style='List Bullet')
        doc.add_paragraph(self.environment, style='List Bullet 2')
        doc.add_paragraph("Outils :", style='List Bullet')
        doc.add_paragraph(self.tools, style='List Bullet 2')
        
        # 4. Fonctionnalités à tester / à ne pas tester
        doc.add_heading("4. Fonctionnalités à tester", level=2)
        doc.add_paragraph("À inclure :")
        scope_items = [s.strip() for s in self.scope.split(',')]
        for item in scope_items:
            doc.add_paragraph(item, style='List Bullet')
        
        # 5. Approche de test
        doc.add_heading("5. Approche de test", level=2)
        doc.add_paragraph(
            f"Méthodologie : {self.methodology}\n\n"
            f"Stratégie :\n"
            f"• Tests manuels + automatisés (si applicable)\n"
            f"• Exécution par modules fonctionnels\n"
            f"• Validation des critères de sortie"
        )
        
        # 6. Critères d'entrée et de sortie
        doc.add_heading("6. Critères d'entrée et de sortie", level=2)
        doc.add_paragraph("Critères de sortie :")
        doc.add_paragraph(self._format_text(self.criteria), style='List Bullet')
        
        # 7. Livrables de test
        doc.add_heading("7. Livrables de test", level=2)
        doc.add_paragraph(
            "• Plan de test (ce document)\n"
            "• Cas de test (Excel)\n"
            "• Résultats d'exécution\n"
            "• Rapport de synthèse\n"
            "• Matrice de traçabilité"
        , style='List Bullet')
        
        # 8. Besoins en environnement
        doc.add_heading("8. Besoins en environnement", level=2)
        doc.add_paragraph(
            f"Environnement : {self.environment}\n"
            f"Outils de test : {self.tools}"
        )
        
        # 9. Responsabilités
        doc.add_heading("9. Responsabilités", level=2)
        resp_table = doc.add_table(rows=4, cols=2)
        resp_table.style = 'Light Grid Accent 1'
        resp_table.cell(0, 0).text = "Rôle"
        resp_table.cell(0, 1).text = "Responsabilité"
        resp_table.cell(1, 0).text = "Test Manager"
        resp_table.cell(1, 1).text = "Pilotage du plan de test"
        resp_table.cell(2, 0).text = "QA Engineer"
        resp_table.cell(2, 1).text = "Rédaction et exécution des cas de test"
        resp_table.cell(3, 0).text = "Stakeholder"
        resp_table.cell(3, 1).text = self._format_text(self.stakeholders)
        
        # 10. Risques et plans de contingence
        doc.add_heading("10. Risques et plans de contingence", level=2)
        doc.add_paragraph("Risques identifiés :")
        risk_items = [r.strip() for r in self.risks.split(',')]
        for risk in risk_items:
            doc.add_paragraph(f"{risk} → Plan de mitigation : À définir", style='List Bullet')
        
        # 11. Calendrier
        doc.add_heading("11. Calendrier", level=2)
        doc.add_paragraph(f"Timeline prévue : {self._format_text(self.timeline)}")
        doc.add_paragraph(
            f"Début : {datetime.now().strftime('%Y-%m-%d')}\n"
            f"Fin estimée : À ajuster selon les résultats"
        )
        
        # Footer
        doc.add_paragraph()
        footer = doc.add_paragraph("---")
        footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        footer_text = doc.add_paragraph(
            f"Généré automatiquement par Générateur de Plan de Test\n"
            f"Project: {self.project_name} | {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        )
        footer_text.alignment = WD_ALIGN_PARAGRAPH.CENTER
        footer_text.runs[0].font.size = Pt(10)
        footer_text.runs[0].font.italic = True
        
        doc.save(output_path)
        return output_path
    
    def export_to_excel(self, output_path: str = None) -> str:
        """Export test plan to Excel (summary + matrix)."""
        if not OPENPYXL_AVAILABLE:
            print("ERROR: openpyxl not available. Cannot generate Excel.")
            return None
        
        if output_path is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = f"plan_test_{self.slug}_{timestamp}.xlsx"
        
        wb = Workbook()
        ws_summary = wb.active
        ws_summary.title = "Résumé du plan"
        
        # Sheet 1: Summary
        self._write_summary_sheet(ws_summary)
        
        # Sheet 2: Scope Matrix
        ws_scope = wb.create_sheet("Portée")
        self._write_scope_sheet(ws_scope)
        
        # Sheet 3: Risk Register
        ws_risks = wb.create_sheet("Registre des risques")
        self._write_risk_sheet(ws_risks)
        
        wb.save(output_path)
        return output_path
    
    def _write_summary_sheet(self, ws):
        """Write summary information."""
        ws.cell(row=1, column=1, value="RÉSUMÉ DU PLAN DE TEST").font = Font(bold=True, size=14)
        ws.merge_cells('A1:B1')
        
        row = 3
        headers = [
            ("Projet", self.project_name),
            ("Environnement", self.environment),
            ("Méthodologie", self.methodology),
            ("Niveaux de test", ", ".join(self.test_levels)),
            ("Timeline", self._format_text(self.timeline)),
            ("Outils", self.tools),
            ("Stakeholders", self._format_text(self.stakeholders)),
            ("Critères de sortie", self._format_text(self.criteria)),
        ]
        
        for label, value in headers:
            cell_label = ws.cell(row=row, column=1, value=label)
            cell_label.font = Font(bold=True)
            ws.cell(row=row, column=2, value=value)
            row += 1
        
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 50
    
    def _write_scope_sheet(self, ws):
        """Write scope/functionality matrix."""
        ws.cell(row=1, column=1, value="Fonctionnalité").font = Font(bold=True)
        ws.cell(row=1, column=2, value="À tester ?").font = Font(bold=True)
        ws.cell(row=1, column=3, value="Priorité").font = Font(bold=True)
        ws.cell(row=1, column=4, value="Notes").font = Font(bold=True)
        
        scope_items = [s.strip() for s in self.scope.split(',')]
        for idx, item in enumerate(scope_items, 2):
            ws.cell(row=idx, column=1, value=item)
            ws.cell(row=idx, column=2, value="Oui")
            ws.cell(row=idx, column=3, value="Moyenne")
            ws.cell(row=idx, column=4, value="")
        
        for col in range(1, 5):
            ws.column_dimensions[chr(64+col)].width = 30
    
    def _write_risk_sheet(self, ws):
        """Write risk register."""
        ws.cell(row=1, column=1, value="Risque").font = Font(bold=True)
        ws.cell(row=1, column=2, value="Sévérité").font = Font(bold=True)
        ws.cell(row=1, column=3, value="Probabilité").font = Font(bold=True)
        ws.cell(row=1, column=4, value="Plan de mitigation").font = Font(bold=True)
        
        risk_items = [r.strip() for r in self.risks.split(',')]
        for idx, risk in enumerate(risk_items, 2):
            ws.cell(row=idx, column=1, value=risk)
            ws.cell(row=idx, column=2, value="Haute")
            ws.cell(row=idx, column=3, value="Moyenne")
            ws.cell(row=idx, column=4, value="À définir")
        
        for col in range(1, 5):
            ws.column_dimensions[chr(64+col)].width = 30


def main():
    parser = argparse.ArgumentParser(description="Generate test plan from config.json (IEEE 829)")
    parser.add_argument("config", help="Path to config.json file")
    parser.add_argument("-o", "--output", help="Output path (Word or Excel)")
    parser.add_argument("--format", choices=["docx", "xlsx", "both"], default="docx", help="Output format")
    
    args = parser.parse_args()
    
    if not Path(args.config).exists():
        print(f"ERROR: Config file not found: {args.config}")
        sys.exit(1)
    
    print(f"📋 Reading config: {args.config}")
    generator = TestPlanGenerator(args.config)
    
    print(f"📄 Generating test plan for: {generator.project_name}")
    
    outputs = []
    
    if args.format in ["docx", "both"]:
        word_path = generator.export_to_docx(args.output if args.format == "docx" else None)
        if word_path:
            print(f"✅ Test plan exported to Word: {word_path}")
            outputs.append(word_path)
    
    if args.format in ["xlsx", "both"]:
        excel_path = generator.export_to_excel(args.output if args.format == "xlsx" else None)
        if excel_path:
            print(f"✅ Test plan exported to Excel: {excel_path}")
            outputs.append(excel_path)
    
    if outputs:
        return 0
    else:
        print("❌ Export failed")
        return 1


if __name__ == "__main__":
    sys.exit(main())
