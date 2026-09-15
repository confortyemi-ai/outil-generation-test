#!/usr/bin/env python3
"""
gen_dashboard.py — Generate test dashboard (HTML + metrics overview)

Usage:
    python gen_dashboard.py config.json
    python gen_dashboard.py config.json --output dashboard.html

Output:
    - dashboard_{project_slug}_{timestamp}.html (Interactive HTML dashboard)
    
Note: Phase 2 will integrate real-time Xray data. This version shows template structure.
"""

import json
import argparse
import sys
from pathlib import Path
from datetime import datetime
from typing import Dict
import re


class DashboardGenerator:
    """Generate test dashboard in HTML."""
    
    def __init__(self, config_path: str):
        """Initialize with config file."""
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = json.load(f)
        
        self.project_name = self.config.get("project", {}).get("name", "Project")
        self.slug = re.sub(r'[^\w]+', '_', self.project_name.lower())[:20]
        self.environment = self.config.get("project", {}).get("environment", "Unknown")
        self.methodology = self.config.get("project", {}).get("methodology", "Unknown")
    
    def export_to_html(self, output_path: str = None) -> str:
        """Export dashboard to HTML."""
        if output_path is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = f"dashboard_{self.slug}_{timestamp}.html"
        
        html_content = self._generate_html()
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        return output_path
    
    def _generate_html(self) -> str:
        """Generate dashboard HTML with charts and metrics."""
        html = f"""<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Tableau de bord — {self.project_name}</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/3.9.1/chart.min.js"></script>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); min-height: 100vh; padding: 20px; }}
        
        .container {{ max-width: 1400px; margin: 0 auto; }}
        
        .header {{ background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); margin-bottom: 20px; }}
        .header h1 {{ color: #003366; margin-bottom: 10px; }}
        .header p {{ color: #666; font-size: 14px; }}
        
        .grid-2 {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 20px; margin-bottom: 20px; }}
        .grid-3 {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(350px, 1fr)); gap: 20px; margin-bottom: 20px; }}
        
        .card {{ background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }}
        .card h3 {{ color: #003366; margin-bottom: 15px; font-size: 16px; border-bottom: 2px solid #003366; padding-bottom: 10px; }}
        
        .metric {{ display: flex; align-items: center; justify-content: space-between; margin: 10px 0; padding: 10px; background: #f9f9f9; border-radius: 4px; }}
        .metric-label {{ font-size: 14px; color: #666; }}
        .metric-value {{ font-size: 24px; font-weight: bold; color: #003366; }}
        
        .status-indicator {{ display: inline-block; width: 12px; height: 12px; border-radius: 50%; margin-right: 8px; }}
        .status-pass {{ background: #4CAF50; }}
        .status-fail {{ background: #f44336; }}
        .status-blocked {{ background: #ff9800; }}
        .status-pending {{ background: #2196F3; }}
        
        .chart-container {{ position: relative; height: 300px; margin: 20px 0; }}
        
        .progress-bar {{ width: 100%; height: 6px; background: #e0e0e0; border-radius: 3px; overflow: hidden; margin: 10px 0; }}
        .progress-fill {{ height: 100%; transition: width 0.3s ease; }}
        .progress-pass {{ background: #4CAF50; }}
        .progress-fail {{ background: #f44336; }}
        .progress-blocked {{ background: #ff9800; }}
        .progress-pending {{ background: #2196F3; }}
        
        .alert {{ padding: 15px; border-radius: 4px; margin: 10px 0; }}
        .alert-info {{ background: #e3f2fd; color: #0d47a1; border-left: 4px solid #2196F3; }}
        .alert-warning {{ background: #fff3e0; color: #e65100; border-left: 4px solid #ff9800; }}
        .alert-error {{ background: #ffebee; color: #c62828; border-left: 4px solid #f44336; }}
        .alert-success {{ background: #e8f5e9; color: #2e7d32; border-left: 4px solid #4CAF50; }}
        
        .footer {{ text-align: center; color: white; margin-top: 40px; font-size: 12px; }}
        
        @media (max-width: 768px) {{
            .grid-2, .grid-3 {{ grid-template-columns: 1fr; }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 Tableau de bord de test — {self.project_name}</h1>
            <p>Environnement : {self.environment} | Méthodologie : {self.methodology}</p>
            <p style="margin-top: 10px; font-size: 12px; color: #999;">Mise à jour : {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}</p>
        </div>
        
        <div class="grid-2">
            <div class="card">
                <h3>📈 Métriques globales</h3>
                <div class="metric">
                    <span class="metric-label">Total cas de test</span>
                    <span class="metric-value">42</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Cas exécutés</span>
                    <span class="metric-value">38</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Taux d'exécution</span>
                    <span class="metric-value">90%</span>
                </div>
                <div class="alert alert-info">
                    <strong>Info :</strong> 4 cas en cours d'exécution
                </div>
            </div>
            
            <div class="card">
                <h3>✅ Résultats d'exécution</h3>
                <div class="metric">
                    <span class="metric-label"><span class="status-indicator status-pass"></span>Réussis</span>
                    <span class="metric-value" style="color: #4CAF50;">32</span>
                </div>
                <div class="metric">
                    <span class="metric-label"><span class="status-indicator status-fail"></span>Échoués</span>
                    <span class="metric-value" style="color: #f44336;">4</span>
                </div>
                <div class="metric">
                    <span class="metric-label"><span class="status-indicator status-blocked"></span>Bloqués</span>
                    <span class="metric-value" style="color: #ff9800;">2</span>
                </div>
                <div class="progress-bar">
                    <div class="progress-fill progress-pass" style="width: 84%;"></div>
                </div>
                <p style="font-size: 12px; color: #666; margin-top: 10px;">Taux de réussite : <strong>84%</strong></p>
            </div>
        </div>
        
        <div class="grid-3">
            <div class="card">
                <h3>🎯 Couverture par module</h3>
                <div style="height: 250px;">
                    <canvas id="moduleChart"></canvas>
                </div>
            </div>
            
            <div class="card">
                <h3>🔴 Distribution des résultats</h3>
                <div style="height: 250px;">
                    <canvas id="statusChart"></canvas>
                </div>
            </div>
            
            <div class="card">
                <h3>📊 Priorité des défauts</h3>
                <div class="metric">
                    <span class="metric-label"><span class="status-indicator status-fail"></span>Critique</span>
                    <span class="metric-value" style="color: #f44336;">1</span>
                </div>
                <div class="metric">
                    <span class="metric-label"><span class="status-indicator" style="background: #ff6f00;"></span>Haute</span>
                    <span class="metric-value" style="color: #ff6f00;">2</span>
                </div>
                <div class="metric">
                    <span class="metric-label"><span class="status-indicator" style="background: #fbc02d;"></span>Moyenne</span>
                    <span class="metric-value" style="color: #fbc02d;">1</span>
                </div>
                <div class="alert alert-warning">
                    <strong>Alerte :</strong> 1 défaut critique à corriger avant release
                </div>
            </div>
        </div>
        
        <div class="grid-2">
            <div class="card">
                <h3>⏱️ Performance</h3>
                <div class="metric">
                    <span class="metric-label">Temps total d'exécution</span>
                    <span class="metric-value">4h 32m</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Temps moyen par cas</span>
                    <span class="metric-value">7m 10s</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Cas les plus rapides</span>
                    <span class="metric-value">UI tests (2m)</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Cas les plus lents</span>
                    <span class="metric-value">Perf tests (15m)</span>
                </div>
            </div>
            
            <div class="card">
                <h3>🚀 Status du cycle de test</h3>
                <div class="alert alert-warning">
                    <strong>⚠️ En cours :</strong> Phase 1 (Tests unitaires + intégration)
                </div>
                <div class="metric">
                    <span class="metric-label">Progression</span>
                    <span class="metric-value">60%</span>
                </div>
                <div class="progress-bar">
                    <div class="progress-fill progress-pending" style="width: 60%;"></div>
                </div>
                <p style="font-size: 12px; color: #666; margin-top: 15px;">
                    <strong>Prochaines étapes :</strong>
                </p>
                <ul style="margin-left: 20px; font-size: 12px; color: #666;">
                    <li>Correction des défauts identifiés (2j)</li>
                    <li>Régression testing (1j)</li>
                    <li>Phase 2 (Sécurité + Performance)</li>
                </ul>
            </div>
        </div>
        
        <div class="card">
            <h3>📋 Défauts ouverts</h3>
            <table style="width: 100%; border-collapse: collapse; font-size: 14px;">
                <thead>
                    <tr style="background: #f5f5f5;">
                        <th style="padding: 10px; text-align: left; border-bottom: 2px solid #ddd;">ID</th>
                        <th style="padding: 10px; text-align: left; border-bottom: 2px solid #ddd;">Titre</th>
                        <th style="padding: 10px; text-align: left; border-bottom: 2px solid #ddd;">Priorité</th>
                        <th style="padding: 10px; text-align: left; border-bottom: 2px solid #ddd;">Assigné</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td style="padding: 10px; border-bottom: 1px solid #ddd;">DEF-001</td>
                        <td style="padding: 10px; border-bottom: 1px solid #ddd;">Export PDF ne fonctionne pas</td>
                        <td style="padding: 10px; border-bottom: 1px solid #ddd;"><span style="color: #f44336; font-weight: bold;">Critique</span></td>
                        <td style="padding: 10px; border-bottom: 1px solid #ddd;">Dev Team</td>
                    </tr>
                    <tr>
                        <td style="padding: 10px; border-bottom: 1px solid #ddd;">DEF-002</td>
                        <td style="padding: 10px; border-bottom: 1px solid #ddd;">Performance timeout (> 5s)</td>
                        <td style="padding: 10px; border-bottom: 1px solid #ddd;"><span style="color: #ff6f00; font-weight: bold;">Haute</span></td>
                        <td style="padding: 10px; border-bottom: 1px solid #ddd;">Infra Team</td>
                    </tr>
                    <tr>
                        <td style="padding: 10px; border-bottom: 1px solid #ddd;">DEF-003</td>
                        <td style="padding: 10px; border-bottom: 1px solid #ddd;">Message d'erreur en anglais</td>
                        <td style="padding: 10px; border-bottom: 1px solid #ddd;"><span style="color: #fbc02d; font-weight: bold;">Moyenne</span></td>
                        <td style="padding: 10px; border-bottom: 1px solid #ddd;">Frontend Team</td>
                    </tr>
                    <tr>
                        <td style="padding: 10px; border-bottom: 1px solid #ddd;">DEF-004</td>
                        <td style="padding: 10px; border-bottom: 1px solid #ddd;">Typo dans la doc</td>
                        <td style="padding: 10px; border-bottom: 1px solid #ddd;"><span style="color: #2196F3; font-weight: bold;">Basse</span></td>
                        <td style="padding: 10px; border-bottom: 1px solid #ddd;">Tech Writer</td>
                    </tr>
                </tbody>
            </table>
        </div>
        
        <div class="footer">
            <p>Tableau de bord généré par Générateur de Test Management</p>
            <p style="margin-top: 10px;">Phase 2 : Intégration Xray Cloud pour données en temps réel</p>
        </div>
    </div>
    
    <script>
        // Module coverage chart
        const moduleCtx = document.getElementById('moduleChart').getContext('2d');
        new Chart(moduleCtx, {{
            type: 'bar',
            data: {{
                labels: ['Auth', 'Search', 'Export', 'Performance', 'Security'],
                datasets: [{{
                    label: 'Couverture %',
                    data: [95, 85, 70, 60, 40],
                    backgroundColor: ['#4CAF50', '#4CAF50', '#FFC107', '#FF9800', '#F44336']
                }}]
            }},
            options: {{
                indexAxis: 'y',
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{ legend: {{ display: false }} }},
                scales: {{ x: {{ beginAtZero: true, max: 100 }} }}
            }}
        }});
        
        // Status distribution chart
        const statusCtx = document.getElementById('statusChart').getContext('2d');
        new Chart(statusCtx, {{
            type: 'doughnut',
            data: {{
                labels: ['Pass', 'Fail', 'Blocked'],
                datasets: [{{
                    data: [32, 4, 2],
                    backgroundColor: ['#4CAF50', '#F44336', '#FF9800']
                }}]
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{ legend: {{ position: 'bottom' }} }}
            }}
        }});
    </script>
</body>
</html>
        """
        
        return html


def main():
    parser = argparse.ArgumentParser(description="Generate test dashboard")
    parser.add_argument("config", help="Path to config.json file")
    parser.add_argument("-o", "--output", help="Output HTML path (optional)")
    
    args = parser.parse_args()
    
    if not Path(args.config).exists():
        print(f"ERROR: Config file not found: {args.config}")
        sys.exit(1)
    
    print(f"🎨 Reading config: {args.config}")
    generator = DashboardGenerator(args.config)
    
    print(f"📊 Generating dashboard for: {generator.project_name}")
    html_path = generator.export_to_html(args.output)
    
    if html_path:
        print(f"✅ Dashboard exported to: {html_path}")
        return 0
    else:
        print("❌ Export failed")
        return 1


if __name__ == "__main__":
    sys.exit(main())
