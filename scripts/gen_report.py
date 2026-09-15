#!/usr/bin/env python3
"""
gen_report.py — Generate test summary report (HTML + Excel with synthetic results)

Usage:
    python gen_report.py config.json
    python gen_report.py config.json --output report.html

Output:
    - report_{project_slug}_{timestamp}.html (HTML report)
    - report_{project_slug}_{timestamp}.xlsx (Excel with metrics)
"""

import json
import argparse
import sys
from pathlib import Path
from datetime import datetime
from typing import Dict, List
import re

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False


class ReportGenerator:
    """Generate test execution report (synthetic + template for real data)."""
    
    def __init__(self, config_path: str):
        """Initialize with config file."""
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = json.load(f)
        
        self.project_name = self.config.get("project", {}).get("name", "Project")
        self.slug = re.sub(r'[^\w]+', '_', self.project_name.lower())[:20]
        self.environment = self.config.get("project", {}).get("environment", "Unknown")
        self.methodology = self.config.get("project", {}).get("methodology", "Unknown")
        self.test_levels = self.config.get("project", {}).get("testLevels", [])
        
        # Synthetic test results
        self.test_results = [
            {"id": "TC-001", "title": "Authentification nominale", "status": "Pass", "duration": "5m", "notes": "✓"},
            {"id": "TC-002", "title": "Affichage des données", "status": "Pass", "duration": "3m", "notes": "✓"},
            {"id": "TC-003", "title": "Recherche exacte", "status": "Pass", "duration": "4m", "notes": "✓"},
            {"id": "TC-004", "title": "Export de fiche", "status": "Fail", "duration": "2m", "notes": "Format Excel incorrect"},
            {"id": "TC-005", "title": "Données limites (vide)", "status": "Pass", "duration": "3m", "notes": "✓"},
            {"id": "TC-006", "title": "Performance : timeout", "status": "Fail", "duration": "15m", "notes": "Délai > 5s"},
            {"id": "TC-007", "title": "Erreur : entrée invalide", "status": "Blocked", "duration": "0m", "notes": "Env. non prêt"},
        ]
    
    def export_to_html(self, output_path: str = None) -> str:
        """Export report to HTML."""
        if output_path is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = f"report_{self.slug}_{timestamp}.html"
        
        html_content = self._generate_html()
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        return output_path
    
    def _generate_html(self) -> str:
        """Generate HTML report content."""
        passed = sum(1 for r in self.test_results if r["status"] == "Pass")
        failed = sum(1 for r in self.test_results if r["status"] == "Fail")
        blocked = sum(1 for r in self.test_results if r["status"] == "Blocked")
        total = len(self.test_results)
        pass_rate = (passed / total * 100) if total > 0 else 0
        
        total_duration = sum(
            int(r["duration"].replace("m", ""))
            for r in self.test_results
            if r["duration"] != "0m"
        )
        
        rows_html = ""
        for result in self.test_results:
            status_class = "pass" if result["status"] == "Pass" else ("fail" if result["status"] == "Fail" else "blocked")
            rows_html += f"""
        <tr class="status-{status_class}">
            <td>{result["id"]}</td>
            <td>{result["title"]}</td>
            <td><span class="status-badge status-{status_class}">{result["status"]}</span></td>
            <td>{result["duration"]}</td>
            <td>{result["notes"]}</td>
        </tr>
            """
        
        html = f"""<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Rapport de test — {self.project_name}</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f5f5f5; padding: 20px; }}
        .container {{ max-width: 1200px; margin: 0 auto; background: white; padding: 40px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }}
        h1 {{ color: #003366; margin-bottom: 10px; border-bottom: 3px solid #003366; padding-bottom: 10px; }}
        h2 {{ color: #003366; margin-top: 30px; margin-bottom: 15px; font-size: 18px; }}
        .metadata {{ display: grid; grid-template-columns: repeat(2, 1fr); gap: 15px; margin: 20px 0; font-size: 14px; }}
        .metadata-item {{ background: #f9f9f9; padding: 12px; border-left: 4px solid #003366; }}
        .metadata-item strong {{ display: block; color: #003366; margin-bottom: 5px; }}
        
        .metrics {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px; margin: 20px 0; }}
        .metric-card {{ background: white; border: 1px solid #ddd; border-radius: 8px; padding: 20px; text-align: center; }}
        .metric-card .number {{ font-size: 32px; font-weight: bold; margin: 10px 0; }}
        .metric-card.pass {{ border-left: 4px solid #4CAF50; }}
        .metric-card.pass .number {{ color: #4CAF50; }}
        .metric-card.fail {{ border-left: 4px solid #f44336; }}
        .metric-card.fail .number {{ color: #f44336; }}
        .metric-card.blocked {{ border-left: 4px solid #ff9800; }}
        .metric-card.blocked .number {{ color: #ff9800; }}
        .metric-card.total {{ border-left: 4px solid #2196F3; }}
        .metric-card.total .number {{ color: #2196F3; }}
        
        table {{ width: 100%; border-collapse: collapse; margin: 20px 0; }}
        th {{ background: #003366; color: white; padding: 12px; text-align: left; font-weight: 600; }}
        td {{ padding: 10px 12px; border-bottom: 1px solid #e0e0e0; }}
        tr:hover {{ background: #f9f9f9; }}
        
        .status-badge {{ display: inline-block; padding: 4px 12px; border-radius: 4px; font-weight: 600; font-size: 12px; }}
        .status-pass {{ background: #c8e6c9; color: #2e7d32; }}
        .status-fail {{ background: #ffcdd2; color: #c62828; }}
        .status-blocked {{ background: #ffe0b2; color: #e65100; }}
        
        tr.status-pass {{ background: #f1f8f4; }}
        tr.status-fail {{ background: #fde8e8; }}
        tr.status-blocked {{ background: #fff3e0; }}
        
        .summary {{ background: #f0f4f8; padding: 20px; border-radius: 8px; margin: 20px 0; }}
        .summary p {{ margin: 8px 0; line-height: 1.6; }}
        
        .footer {{ margin-top: 40px; padding-top: 20px; border-top: 1px solid #ddd; font-size: 12px; color: #666; text-align: center; }}
        .progress-bar {{ width: 100%; height: 8px; background: #e0e0e0; border-radius: 4px; overflow: hidden; margin: 10px 0; }}
        .progress-fill {{ height: 100%; background: #4CAF50; }}
    </style>
</head>
<body>
    <div class="container">
        <h1>Rapport de test — {self.project_name}</h1>
        
        <div class="metadata">
            <div class="metadata-item">
                <strong>Projet</strong>
                {self.project_name}
            </div>
            <div class="metadata-item">
                <strong>Environnement</strong>
                {self.environment}
            </div>
            <div class="metadata-item">
                <strong>Méthodologie</strong>
                {self.methodology}
            </div>
            <div class="metadata-item">
                <strong>Date du rapport</strong>
                {datetime.now().strftime('%d/%m/%Y %H:%M')}
            </div>
        </div>
        
        <h2>Résumé des résultats</h2>
        <div class="metrics">
            <div class="metric-card total">
                <div>Total</div>
                <div class="number">{total}</div>
                <div>cas de test</div>
            </div>
            <div class="metric-card pass">
                <div>Réussis</div>
                <div class="number">{passed}</div>
                <div>{pass_rate:.1f}%</div>
            </div>
            <div class="metric-card fail">
                <div>Échoués</div>
                <div class="number">{failed}</div>
                <div>{failed/total*100:.1f}%</div>
            </div>
            <div class="metric-card blocked">
                <div>Bloqués</div>
                <div class="number">{blocked}</div>
                <div>{blocked/total*100:.1f}%</div>
            </div>
        </div>
        
        <div class="summary">
            <p><strong>Taux de réussite :</strong></p>
            <div class="progress-bar">
                <div class="progress-fill" style="width: {pass_rate}%;"></div>
            </div>
            <p><strong>Analyse :</strong> {passed} cas réussis sur {total}. {failed} défaut(s) identifié(s). {blocked} cas bloqué(s) par dépendance(s).</p>
            <p><strong>Durée totale d'exécution :</strong> {total_duration} minutes</p>
        </div>
        
        <h2>Détails des cas de test</h2>
        <table>
            <thead>
                <tr>
                    <th>ID</th>
                    <th>Titre</th>
                    <th>Statut</th>
                    <th>Durée</th>
                    <th>Notes</th>
                </tr>
            </thead>
            <tbody>
                {rows_html}
            </tbody>
        </table>
        
        <h2>Recommandations</h2>
        <div class="summary">
            <p>✓ <strong>Cas nominaux :</strong> Tous les scénarios principaux passent.</p>
            <p>⚠ <strong>Défauts détectés :</strong></p>
            <ul style="margin-left: 20px;">
                <li>Défaut #1 : Format Excel export incorrect (priorité Haute)</li>
                <li>Défaut #2 : Performance insuffisante (>5s, priorité Haute)</li>
            </ul>
            <p>→ <strong>Action requise :</strong> Corriger les défauts avant release en production.</p>
            <p>→ <strong>Prochaines étapes :</strong> Régression testing après corrections.</p>
        </div>
        
        <div class="footer">
            <p>Rapport généré automatiquement par Générateur de Rapport de Test</p>
            <p>{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
        </div>
    </div>
</body>
</html>
        """
        
        return html
    
    def export_to_excel(self, output_path: str = None) -> str:
        """Export report to Excel with metrics."""
        if not OPENPYXL_AVAILABLE:
            print("WARNING: openpyxl not available. Skipping Excel export.")
            return None
        
        if output_path is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = f"report_{self.slug}_{timestamp}.xlsx"
        
        wb = Workbook()
        ws_summary = wb.active
        ws_summary.title = "Résumé"
        
        # Summary metrics
        self._write_summary_sheet(ws_summary)
        
        # Detailed results
        ws_details = wb.create_sheet("Résultats détaillés")
        self._write_details_sheet(ws_details)
        
        wb.save(output_path)
        return output_path
    
    def _write_summary_sheet(self, ws):
        """Write summary metrics."""
        ws.cell(row=1, column=1, value="RÉSUMÉ DU RAPPORT DE TEST").font = Font(bold=True, size=12)
        
        passed = sum(1 for r in self.test_results if r["status"] == "Pass")
        failed = sum(1 for r in self.test_results if r["status"] == "Fail")
        blocked = sum(1 for r in self.test_results if r["status"] == "Blocked")
        total = len(self.test_results)
        
        metrics = [
            ("Projet", self.project_name),
            ("Date du rapport", datetime.now().strftime("%Y-%m-%d %H:%M")),
            ("", ""),
            ("Total cas de test", total),
            ("Cas réussis", f"{passed} ({passed/total*100:.1f}%)"),
            ("Cas échoués", f"{failed} ({failed/total*100:.1f}%)"),
            ("Cas bloqués", f"{blocked} ({blocked/total*100:.1f}%)"),
        ]
        
        for row_idx, (label, value) in enumerate(metrics, 3):
            if label:
                cell_label = ws.cell(row=row_idx, column=1, value=label)
                cell_label.font = Font(bold=True)
                ws.cell(row=row_idx, column=2, value=value)
        
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 30
    
    def _write_details_sheet(self, ws):
        """Write detailed test results."""
        headers = ["ID", "Titre", "Statut", "Durée", "Notes"]
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="003366", end_color="003366", fill_type="solid")
        
        for row_idx, result in enumerate(self.test_results, 2):
            ws.cell(row=row_idx, column=1, value=result["id"])
            ws.cell(row=row_idx, column=2, value=result["title"])
            
            status_cell = ws.cell(row=row_idx, column=3, value=result["status"])
            color = {"Pass": "C6EFCE", "Fail": "FFC7CE", "Blocked": "FFEB9C"}.get(result["status"], "FFFFFF")
            status_cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
            
            ws.cell(row=row_idx, column=4, value=result["duration"])
            ws.cell(row=row_idx, column=5, value=result["notes"])
        
        for col in range(1, 6):
            ws.column_dimensions[get_column_letter(col)].width = 25


def main():
    parser = argparse.ArgumentParser(description="Generate test summary report")
    parser.add_argument("config", help="Path to config.json file")
    parser.add_argument("-o", "--output", help="Output path (optional)")
    parser.add_argument("--format", choices=["html", "xlsx", "both"], default="html")
    
    args = parser.parse_args()
    
    if not Path(args.config).exists():
        print(f"ERROR: Config file not found: {args.config}")
        sys.exit(1)
    
    print(f"📊 Reading config: {args.config}")
    generator = ReportGenerator(args.config)
    
    print(f"📝 Generating report for: {generator.project_name}")
    
    outputs = []
    
    if args.format in ["html", "both"]:
        html_path = generator.export_to_html(args.output if args.format == "html" else None)
        if html_path:
            print(f"✅ Report exported to HTML: {html_path}")
            outputs.append(html_path)
    
    if args.format in ["xlsx", "both"]:
        excel_path = generator.export_to_excel(args.output if args.format == "xlsx" else None)
        if excel_path:
            print(f"✅ Report exported to Excel: {excel_path}")
            outputs.append(excel_path)
    
    if outputs:
        return 0
    else:
        print("❌ Export failed")
        return 1


if __name__ == "__main__":
    sys.exit(main())
