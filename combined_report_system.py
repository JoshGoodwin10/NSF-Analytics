"""
Requirements:
    pip install pandas openpyxl matplotlib seaborn jinja2 tkinterdnd2 Pillow
"""

import sys
import warnings
from datetime import datetime
from pathlib import Path

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from jinja2 import Template

warnings.filterwarnings('ignore')
plt.style.use('seaborn-v0_8-whitegrid')
sns.set_palette("husl")

# Shared helpers

def build_pdf(pdf_path: Path, title: str, stats: dict, charts_dir: Path) -> bool:
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib import colors
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import cm
        from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                        Table, TableStyle, Image, PageBreak,
                                        HRFlowable)
        from reportlab.lib.enums import TA_CENTER, TA_LEFT

        W, H = A4
        doc  = SimpleDocTemplate(
            str(pdf_path), pagesize=A4,
            leftMargin=2*cm, rightMargin=2*cm,
            topMargin=2*cm,  bottomMargin=2*cm,
        )

        styles   = getSampleStyleSheet()
        BLUE     = colors.HexColor("#2E75B6")
        LT_BLUE  = colors.HexColor("#EBF3FB")
        GREEN    = colors.HexColor("#2ecc71")
        GREY     = colors.HexColor("#555555")

        h1_style = ParagraphStyle("H1", parent=styles["Heading1"],
                                  textColor=BLUE, fontSize=18, spaceAfter=6)
        h2_style = ParagraphStyle("H2", parent=styles["Heading2"],
                                  textColor=BLUE, fontSize=13, spaceAfter=4,
                                  spaceBefore=12)
        body_style = ParagraphStyle("Body", parent=styles["Normal"],
                                    textColor=GREY, fontSize=9, spaceAfter=3)
        small_style = ParagraphStyle("Small", parent=styles["Normal"],
                                     textColor=GREY, fontSize=8)

        def kv_table(rows: list[tuple]) -> Table:
            """Two-column key/value table."""
            data   = [[Paragraph(f"<b>{k}</b>", small_style),
                       Paragraph(str(v), small_style)] for k, v in rows]
            tbl    = Table(data, colWidths=[7*cm, 9*cm])
            tbl.setStyle(TableStyle([
                ("BACKGROUND", (0,0), (-1,0), LT_BLUE),
                ("ROWBACKGROUNDS", (0,0), (-1,-1), [colors.white, LT_BLUE]),
                ("GRID", (0,0), (-1,-1), 0.3, colors.HexColor("#cccccc")),
                ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                ("TOPPADDING", (0,0), (-1,-1), 4),
                ("BOTTOMPADDING", (0,0), (-1,-1), 4),
                ("LEFTPADDING", (0,0), (-1,-1), 6),
            ]))
            return tbl

        def ranked_table(d: dict, col1="Item", col2="Count", pct=False) -> Table:
            total = sum(d.values()) or 1
            header = [Paragraph(f"<b>{col1}</b>", small_style),
                      Paragraph(f"<b>{col2}</b>", small_style)]
            if pct:
                header.append(Paragraph("<b>%</b>", small_style))
            data = [header]
            for i, (k, v) in enumerate(d.items(), 1):
                row = [Paragraph(f"{i}. {k}", small_style),
                       Paragraph(str(v), small_style)]
                if pct:
                    row.append(Paragraph(f"{v/total*100:.1f}%", small_style))
                data.append(row)
            widths = ([1*cm, 7*cm, 7*cm, 2*cm] if pct
                      else [1*cm, 10*cm, 5.5*cm])
            tbl = Table(data, colWidths=widths[:len(header)+1] if not pct else widths)
            tbl = Table(data)
            tbl.setStyle(TableStyle([
                ("BACKGROUND", (0,0), (-1,0), BLUE),
                ("TEXTCOLOR",  (0,0), (-1,0), colors.white),
                ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, LT_BLUE]),
                ("GRID", (0,0), (-1,-1), 0.3, colors.HexColor("#cccccc")),
                ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                ("TOPPADDING", (0,0), (-1,-1), 3),
                ("BOTTOMPADDING", (0,0), (-1,-1), 3),
                ("LEFTPADDING", (0,0), (-1,-1), 5),
            ]))
            return tbl

        def insert_chart(name: str) -> Image | None:
            p = charts_dir / f"{name}.png"
            if p.exists():
                img = Image(str(p), width=W - 5*cm, height=7*cm, kind="proportional")
                return img
            return None

        story = []

        # Cover
        story.append(Spacer(1, 2*cm))
        story.append(Paragraph(title, h1_style))
        story.append(HRFlowable(width="100%", thickness=1, color=BLUE))
        story.append(Spacer(1, 0.3*cm))
        story.append(Paragraph(f"Generated: {stats.get('report_generated','')}", body_style))
        story.append(Spacer(1, 0.6*cm))

        # Key metrics
        kv_rows = []
        for key in ("total_dropouts","total_learners","with_departure_info",
                    "without_departure_info","recent_dropouts_12mo","high_demand_2024"):
            if key in stats:
                label = key.replace("_", " ").title()
                kv_rows.append((label, stats[key]))
        if "age_stats" in stats:
            kv_rows.append(("Average Age at Start", f"{stats['age_stats']['mean']} yrs"))
        if "duration_stats" in stats:
            kv_rows.append(("Avg Programme Duration", f"{stats['duration_stats']['mean']} months"))
        if kv_rows:
            story.append(Paragraph("Key Metrics", h2_style))
            story.append(kv_table(kv_rows))
            story.append(Spacer(1, 0.5*cm))

        sections = [
            ("learners_per_site",         None,                        "Host Employers",                   "Site",     False),
            ("top_programmes_wil",        "top_programmes_wil",        "Top WIL Learning Programmes",      "Programme",True),
            ("top_programmes_graduates",  "top_programmes_graduates",  "Top Graduate Learning Programmes", "Programme",True),
            ("top_sectors",             "top_sectors",               "Top Economic Sectors",             "Sector",   True),
            ("learner_status",          "learner_status",            "Learner Status",                 "Status",   True),
            ("gender",                  "gender_distribution",       "Gender Distribution",            "Gender",   True),
            ("population_group",        "population_group",          "Population Group",               "Group",    True),
            ("household_language",      None,                        "Household Language",             "Language", True),
            ("reasons",                 "reasons",               "Reasons for Leaving",            "Reason",   True),
            ("residential_province",    "residential_province",      "Residential Province",           "Province", True),
            ("dropouts_by_year",        "dropouts_by_year",          "Dropouts by Year",               "Year",     False),
            ("dropouts_by_month",       "dropouts_by_month",         "Dropouts by Month",              "Month",    False),
        ]

        for stat_key, chart_name, section_title, col1, pct in sections:
            if stat_key not in stats or not stats[stat_key]:
                continue
            story.append(Paragraph(section_title, h2_style))
            if chart_name:
                img = insert_chart(chart_name)
                if img:
                    story.append(img)
                    story.append(Spacer(1, 0.2*cm))
            story.append(ranked_table(stats[stat_key], col1=col1, col2="Count", pct=pct))
            story.append(Spacer(1, 0.4*cm))

        if stats.get("dropouts_by_year_month"):
            story.append(Paragraph("Dropouts Over Time", h2_style))
            img = insert_chart("dropouts_timeline")
            if img:
                story.append(img)
            story.append(Spacer(1, 0.4*cm))

        if stats.get("disability"):
            active = {k: v for k, v in stats["disability"].items() if v > 0}
            if active:
                story.append(Paragraph("Disability Types", h2_style))
                img = insert_chart("disability_types")
                if img:
                    story.append(img)
                story.append(ranked_table(
                    {k.replace("_"," ").title(): v for k, v in active.items()},
                    col1="Type", pct=False))
                story.append(Spacer(1, 0.4*cm))

        if stats.get("metro_rural_crosstab"):
            story.append(Paragraph("Metro and Rural Classification", h2_style))
            img = insert_chart("metro_rural")
            if img:
                story.append(img)
            story.append(Spacer(1, 0.4*cm))

        for chart, label in [("age_distribution","Age Distribution"),
                              ("duration_distribution","Programme Duration Distribution")]:
            img = insert_chart(chart)
            if img:
                story.append(Paragraph(label, h2_style))
                story.append(img)
                story.append(Spacer(1, 0.4*cm))

        hd = {}
        if "high_demand_2020" in stats: hd["2020 List"] = stats["high_demand_2020"]
        if "high_demand_2024" in stats: hd["2024 List"] = stats["high_demand_2024"]
        if hd:
            story.append(Paragraph("High-Demand Occupations", h2_style))
            img = insert_chart("high_demand")
            if img:
                story.append(img)
            story.append(kv_table(list(hd.items())))
            story.append(Spacer(1, 0.4*cm))

        doc.build(story)
        print(f"  PDF saved: {pdf_path}")
        return True

    except ImportError as e:
        print(f"\n  ⚠  reportlab not installed ({e}). Run:  pip install reportlab")
        return False
    except Exception as e:
        print(f"\n  ⚠  PDF build failed: {e}")
        return False


# Dropouts Report

class DropoutsReportSystem:
    """Reporting system for Dropouts Master List analytics."""

    def __init__(self, input_path: Path, output_dir: Path):
        self.input_path = input_path
        self.output_dir = output_dir
        self.df = None
        self.summary_stats = {}
        self.charts = []

        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.charts_dir = self.output_dir / "dropout_charts"
        self.charts_dir.mkdir(parents=True, exist_ok=True)

    def load_data(self):
        print(f"  Loading data from: {self.input_path}")
        try:
            raw_df = pd.read_excel(self.input_path, sheet_name="Dropouts Master List",
                                   header=1, dtype=str)
        except ValueError:
            raw_df = pd.read_excel(self.input_path, sheet_name=0, header=1, dtype=str)

        self.df = raw_df.copy()
        print(f"  Loaded {len(self.df)} records with {len(self.df.columns)} columns")
        return self

    def clean_column_names(self):
        cleaned = []
        for col in self.df.columns:
            c = str(col)
            for old, new in [(' ','_'),('/','_'),(chr(92),'_'),('?',''),('<<',''),
                             ('>>',''),('(',''),(')',''),(('-','_'))]:
                c = c.replace(old, new)
            c = c.replace('__', '_').strip('_')
            cleaned.append(c)
        self.df.columns = cleaned
        return self

    def create_derived_fields(self):
        def find_col(partial):
            m = [c for c in self.df.columns if partial.lower() in c.lower()]
            return m[0] if m else None

        date_col = find_col('Date_of_Termination')
        if date_col:
            self.df['termination_date_parsed'] = pd.to_datetime(self.df[date_col], errors='coerce')
            self.df['termination_year']       = self.df['termination_date_parsed'].dt.year
            self.df['termination_month']      = self.df['termination_date_parsed'].dt.month_name()
            self.df['termination_year_month'] = self.df['termination_date_parsed'].dt.strftime('%Y-%m')

        info_col = find_col('Information_related_to_departure')
        if info_col:
            self.df['has_departure_info'] = (
                self.df[info_col].notna() &
                (self.df[info_col].astype(str).str.strip() != '')
            )
        return self

    def generate_summary_statistics(self):
        stats = {}

        def find_col(partial):
            m = [c for c in self.df.columns if partial.lower() in c.lower()]
            return m[0] if m else None

        surname_col = find_col('Student_Surname')
        name_col    = find_col('Student_Name')
        valid_rows  = 0
        for idx in range(len(self.df)):
            for col in [surname_col, name_col]:
                if col and pd.notna(self.df.at[idx, col]) and str(self.df.at[idx, col]).strip():
                    valid_rows += 1
                    break
        if valid_rows == 0:
            valid_rows = self.df.dropna(how='all').shape[0]
        stats['total_dropouts'] = valid_rows

        info_col = find_col('Information_related_to_departure')
        if info_col:
            has_info = self.df[info_col].notna() & (self.df[info_col].astype(str).str.strip() != '')
            stats['with_departure_info']    = int(has_info.sum())
            stats['without_departure_info'] = valid_rows - int(has_info.sum())

        reason_col = (find_col('Reason_for_Laving_the_Programme') or
                      find_col('Reason_for_Leaving_the_Programme') or
                      find_col('Reason_for_Leaving'))
        if reason_col:
            stats['reasons'] = self.df[reason_col].value_counts().to_dict()

        if 'termination_year' in self.df.columns:
            stats['dropouts_by_year'] = self.df['termination_year'].value_counts().sort_index().to_dict()

        if 'termination_month' in self.df.columns:
            month_order = ['January','February','March','April','May','June',
                           'July','August','September','October','November','December']
            mc = self.df['termination_month'].value_counts().to_dict()
            stats['dropouts_by_month'] = {m: mc.get(m, 0) for m in month_order}

        if 'termination_year_month' in self.df.columns:
            stats['dropouts_by_year_month'] = (
                self.df['termination_year_month'].value_counts().sort_index().to_dict()
            )

        if 'termination_date_parsed' in self.df.columns:
            cutoff = pd.Timestamp.now() - pd.DateOffset(months=12)
            stats['recent_dropouts_12mo'] = int((self.df['termination_date_parsed'] >= cutoff).sum())

        stats['report_generated'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        self.summary_stats = stats
        return self

    def create_charts(self):
        charts = []

        def save_chart(name, fig):
            path = self.charts_dir / f"{name}.png"
            fig.savefig(path, dpi=150, bbox_inches='tight', facecolor='white')
            plt.close(fig)
            charts.append(str(path))

        if 'with_departure_info' in self.summary_stats:
            fig, ax = plt.subplots(figsize=(6, 6))
            data = pd.Series({'Has Departure Info': self.summary_stats['with_departure_info'],
                               'No Departure Info':  self.summary_stats['without_departure_info']})
            ax.pie(data.values, labels=data.index, autopct='%1.1f%%',
                   colors=['#67c7f1','#B4C7E7'], startangle=90)
            ax.set_title('Departure Information Coverage', fontsize=14, fontweight='bold')
            save_chart('departure_info_coverage', fig)

        if 'reasons' in self.summary_stats:
            fig, ax = plt.subplots(figsize=(12, 7))
            data   = pd.Series(self.summary_stats['reasons']).sort_values(ascending=True)
            colors = sns.color_palette("viridis", len(data))
            data.plot(kind='barh', ax=ax, color=colors)
            ax.set_title('Reasons for Leaving', fontsize=14, fontweight='bold')
            ax.set_xlabel('Number of Dropouts')
            plt.tight_layout()
            save_chart('reasons', fig)

        if self.summary_stats.get('dropouts_by_year'):
            data = pd.to_numeric(pd.Series(self.summary_stats['dropouts_by_year']).sort_index(), errors='coerce').dropna()
            if data.sum() > 0:
                fig, ax = plt.subplots(figsize=(10, 5))
                data.plot(kind='bar', ax=ax, color=sns.color_palette("coolwarm", len(data)))
                ax.set_title('Dropouts by Year', fontsize=14, fontweight='bold')
                ax.set_xlabel('Year'); ax.set_ylabel('Number of Dropouts')
                plt.xticks(rotation=45, ha='right'); plt.tight_layout()
                save_chart('dropouts_by_year', fig)

        if self.summary_stats.get('dropouts_by_month'):
            data = pd.to_numeric(pd.Series(self.summary_stats['dropouts_by_month']), errors='coerce').dropna()
            if data.sum() > 0:
                fig, ax = plt.subplots(figsize=(12, 5))
                data.plot(kind='bar', ax=ax, color=sns.color_palette("husl", len(data)))
                ax.set_title('Dropouts by Month (All Years)', fontsize=14, fontweight='bold')
                ax.set_xlabel('Month'); ax.set_ylabel('Number of Dropouts')
                plt.xticks(rotation=45, ha='right'); plt.tight_layout()
                save_chart('dropouts_by_month', fig)

        if self.summary_stats.get('dropouts_by_year_month'):
            data = pd.to_numeric(pd.Series(self.summary_stats['dropouts_by_year_month']).sort_index(), errors='coerce').dropna()
            if data.sum() > 0:
                fig, ax = plt.subplots(figsize=(14, 5))
                ax.plot(range(len(data)), data.values, marker='o', color='#67c7f1', linewidth=2, markersize=6)
                ax.fill_between(range(len(data)), data.values, alpha=0.3, color='#67c7f1')
                ax.set_title('Dropouts Over Time (Year-Month)', fontsize=14, fontweight='bold')
                step = max(1, len(data) // 12)
                ax.set_xticks(range(0, len(data), step))
                ax.set_xticklabels([data.index[i] for i in range(0, len(data), step)], rotation=45, ha='right')
                plt.tight_layout()
                save_chart('dropouts_timeline', fig)

        self.charts = charts
        print(f"  Created {len(charts)} charts")
        return self

    def generate_html_report(self) -> Path:
        html_template = """
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Dropouts Master List Analytics Report</title>
<style>
  * { margin:0; padding:0; box-sizing:border-box; }
  body { font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif; background:#f5f7fa; color:#333; line-height:1.6; }
  .container { max-width:1100px; margin:0 auto; padding:20px; }
  header { background:linear-gradient(135deg,#67c7f1 0%,#5bbcec 100%); color:white; padding:40px 20px;
           text-align:center; border-radius:10px; margin-bottom:30px; }
  header h1 { font-size:2.2em; margin-bottom:8px; }
  header p  { font-size:1.05em; opacity:.9; }
  .stats-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(220px,1fr)); gap:20px; margin-bottom:30px; }
  .stat-card { background:white; padding:22px; border-radius:10px; box-shadow:0 2px 4px rgba(0,0,0,.1);
               border-left:4px solid #67c7f1; }
  .stat-card h3 { color:#67c7f1; font-size:.85em; text-transform:uppercase; letter-spacing:1px; margin-bottom:8px; }
  .stat-card .value { font-size:1.9em; font-weight:bold; }
  .section { background:white; padding:28px; border-radius:10px; margin-bottom:28px;
             box-shadow:0 2px 4px rgba(0,0,0,.1); }
  .section h2 { color:#67c7f1; font-size:1.4em; margin-bottom:18px; padding-bottom:8px;
                border-bottom:2px solid #f0f0f0; }
  table { width:100%; border-collapse:collapse; margin-top:14px; }
  th,td { padding:10px 12px; text-align:left; border-bottom:1px solid #eee; }
  th { background:#f8f9fa; font-weight:600; color:#67c7f1; }
  tr:hover { background:#f8f9fa; }
  .chart-container { text-align:center; margin:18px 0; }
  .chart-container img { max-width:100%; height:auto; border-radius:8px;
                          box-shadow:0 2px 8px rgba(0,0,0,.1); }
  footer { text-align:center; padding:20px; color:#999; font-size:.85em; }
</style>
</head>
<body>
<div class="container">
  <header>
    <h1>Dropouts Master List Analytics Report</h1>
    <p>Comprehensive analysis of learner dropout data</p>
    <p style="margin-top:8px;font-size:.88em;">Generated: {{ stats.report_generated }}</p>
  </header>

  <div class="stats-grid">
    <div class="stat-card"><h3>Total Dropouts</h3><div class="value">{{ stats.total_dropouts }}</div></div>
    {% if stats.recent_dropouts_12mo is defined %}
    <div class="stat-card"><h3>Last 12 Months</h3><div class="value">{{ stats.recent_dropouts_12mo }}</div></div>
    {% endif %}
  </div>

  {% if stats.reasons %}
  <div class="section">
    <h2>Reasons for Leaving the Programme</h2>
    <div class="chart-container"><img src="{{ chart_dir }}/reasons.png" alt="Reasons"></div>
    <table>
      <tr><th>Rank</th><th>Reason</th><th>Count</th><th>Percentage</th></tr>
      {% set reason_total = stats.reasons.values() | sum %}
      {% for reason, count in stats.reasons.items() %}
      <tr><td>{{ loop.index }}</td><td>{{ reason }}</td><td>{{ count }}</td>
          <td>{{ "%.1f"|format((count/reason_total)*100) }}%</td></tr>
      {% endfor %}
    </table>
  </div>
  {% endif %}

  {% if stats.dropouts_by_year %}
  <div class="section">
    <h2>Dropouts by Year</h2>
    <div class="chart-container"><img src="{{ chart_dir }}/dropouts_by_year.png" alt="By Year"></div>
    <table>
      <tr><th>Year</th><th>Count</th></tr>
      {% for year, count in stats.dropouts_by_year.items() %}
      <tr><td>{{ year }}</td><td>{{ count }}</td></tr>
      {% endfor %}
    </table>
  </div>
  {% endif %}

  {% if stats.dropouts_by_month %}
  <div class="section">
    <h2>Dropouts by Month</h2>
    <div class="chart-container"><img src="{{ chart_dir }}/dropouts_by_month.png" alt="By Month"></div>
    <table>
      <tr><th>Month</th><th>Count</th></tr>
      {% for month, count in stats.dropouts_by_month.items() %}{% if count > 0 %}
      <tr><td>{{ month }}</td><td>{{ count }}</td></tr>
      {% endif %}{% endfor %}
    </table>
  </div>
  {% endif %}

  {% if stats.dropouts_by_year_month %}
  <div class="section">
    <h2>Dropouts Over Time</h2>
    <div class="chart-container"><img src="{{ chart_dir }}/dropouts_timeline.png" alt="Timeline"></div>
  </div>
  {% endif %}

  {% if stats.with_departure_info is defined %}
  <div class="section">
    <h2>Departure Information Coverage</h2>
    <div class="chart-container"><img src="{{ chart_dir }}/departure_info_coverage.png" alt="Coverage"></div>
  </div>
  {% endif %}

  <footer>Lead HR Consulting | Dropouts Master List Analytics System | Generated on {{ stats.report_generated }}</footer>
</div>
</body>
</html>
"""
        template     = Template(html_template)
        html_content = template.render(stats=self.summary_stats,
                                       chart_dir=str(self.charts_dir))
        report_path  = self.output_dir / "dropouts_report.html"
        report_path.write_text(html_content, encoding='utf-8')
        print(f"  HTML report saved: {report_path}")
        return report_path

    def run(self) -> dict:
        print("\n" + "="*60)
        print("  DROPOUTS MASTER LIST ANALYTICS")
        print("="*60)
        self.load_data()
        self.clean_column_names()
        self.create_derived_fields()
        self.generate_summary_statistics()
        self.create_charts()
        html_path = self.generate_html_report()

        pdf_path = self.output_dir / "dropouts_report.pdf"
        pdf_ok   = build_pdf(pdf_path,
                             "Dropouts Master List Analytics Report",
                             self.summary_stats,
                             self.charts_dir)

        print("\n  ✔  Dropouts report complete!")
        return {"html": html_path, "pdf": pdf_path if pdf_ok else None}


# Learner Report

class LearnerReportSystem:
    """Reporting system for DHET/SETA Learner Data analytics."""

    def __init__(self, input_path: Path, output_dir: Path):
        self.input_path = input_path
        self.output_dir = output_dir
        self.df         = None
        self.summary_stats = {}
        self.charts        = []

        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.charts_dir = self.output_dir / "learner_charts"
        self.charts_dir.mkdir(parents=True, exist_ok=True)

    def load_data(self):
        print(f"  Loading data from: {self.input_path}")
        raw_df = pd.read_excel(self.input_path, sheet_name="(1) Learner Data",
                               header=11, dtype=str)
        raw_df = raw_df.iloc[1:].reset_index(drop=True)
        if raw_df.columns[0] == 'Unnamed: 0' or raw_df.iloc[:, 0].isna().all():
            raw_df = raw_df.iloc[:, 1:]
        self.df = raw_df.copy()
        print(f"  Loaded {len(self.df)} records with {len(self.df.columns)} columns")
        return self

    def clean_column_names(self):
        cleaned = []
        for col in self.df.columns:
            c = str(col)
            for old, new in [(' ','_'),('/','_'),(chr(92),'_'),('?',''),('<<',''),
                             ('>>',''),('(',''),(')',''),(('-','_'))]:
                c = c.replace(old, new)
            c = c.replace('__', '_').strip('_')
            cleaned.append(c)
        self.df.columns = cleaned
        return self

    def create_derived_fields(self):
        def find_col(partial):
            m = [c for c in self.df.columns if partial.lower() in c.lower()]
            return m[0] if m else None

        def parse_date(col_name):
            if col_name and col_name in self.df.columns:
                return pd.to_datetime(self.df[col_name], errors='coerce', format='%Y/%m/%d')
            return None

        dob_col   = find_col('Date_of_Birth')
        start_col = find_col('Learning_Programme_Start_Date')
        end_col   = find_col('Learning_Programme_End_Date')

        if dob_col and start_col:
            dob   = parse_date(dob_col)
            start = parse_date(start_col)
            if dob is not None and start is not None:
                self.df['age_at_start_years'] = ((start - dob).dt.days / 365.25).round(1)

        if start_col and end_col:
            start = parse_date(start_col)
            end   = parse_date(end_col)
            if start is not None and end is not None:
                self.df['programme_duration_months'] = ((end - start).dt.days / 30.44).round(1)

        def to_bool_flag(val):
            if pd.isna(val): return None
            v = str(val).strip().upper()
            if v in ['Y','YES','1','TRUE']: return True
            if v in ['N','NO','0','FALSE']: return False
            return None

        metro_col = find_col('Located_within_one_of_the_Metros')
        rural_col = find_col('Rural')
        if metro_col: self.df['is_metro'] = self.df[metro_col].apply(to_bool_flag)
        if rural_col: self.df['is_rural'] = self.df[rural_col].apply(to_bool_flag)
        return self

    def generate_summary_statistics(self):
        stats = {}

        def find_col(partial):
            m = [c for c in self.df.columns if partial.lower() in c.lower()]
            return m[0] if m else None

        for key, partial in [('learner_status','New_Learner_Existing_Learner_Re_instated'),
                              ('gender','Gender'),
                              ('population_group','Equity_Population_Group'),
                              ('household_language','Household_Language'),
                              ('residential_province','Residential_Address_Province'),
                              ('site_province','Learning_Site_Province')]:
            col = find_col(partial)
            if col: stats[key] = self.df[col].value_counts().to_dict()

        prog_col = find_col('Learning_Programme_Name')
        if prog_col:
            all_progs = self.df[prog_col].value_counts()
            wil_progs = all_progs[all_progs.index.str.startswith('WIL:', na=False)]
            grad_progs = all_progs[all_progs.index.str.startswith('Graduates:', na=False)]
            if len(wil_progs) > 0:
                stats['top_programmes_wil'] = wil_progs.to_dict()
            if len(grad_progs) > 0:
                stats['top_programmes_graduates'] = grad_progs.to_dict()

        sector_col = find_col('ECONOMIC_SECTOR')
        if sector_col:
            stats['top_sectors'] = self.df[sector_col].value_counts().to_dict()

        site_col = find_col('Learning_Site_Name')
        if site_col:
            stats['learners_per_site'] = self.df[site_col].value_counts().to_dict()

        disability_cols = ['Blind','Partially_Sighted','Deaf','Hearing_Impaired',
                           'Deaf_Blind','Neurodevelopmental','Psychosocial_Disability','Physical_Disability']
        disability_stats = {}
        for d in disability_cols:
            col = find_col(d)
            if col:
                series = self.df[col].astype(str).str.strip()
                has_disability = (
                    series.notna() &
                    (series != '') &
                    (series.str.upper() != 'NOT APPLICABLE') &
                    (series.str.upper() != 'N/A') &
                    (series.str.upper() != 'NA')
                )
                disability_stats[d] = int(has_disability.sum())
        stats['disability'] = disability_stats

        for key, partial in [('high_demand_2020','Occupations_in_High_Demand_2020'),
                              ('high_demand_2024','Occupations_in_High_Demand_2024')]:
            col = find_col(partial)
            if col:
                stats[key] = int(self.df[col].str.strip().str.upper().isin(['Y','YES','1']).sum())

        if 'is_metro' in self.df.columns and 'is_rural' in self.df.columns:
            crosstab = pd.crosstab(self.df['is_metro'], self.df['is_rural'], margins=True)
            raw = crosstab.to_dict(orient='index')
            clean = {}
            for k, v in raw.items():
                ck = bool(k) if isinstance(k, (bool,)) or 'bool_' in str(type(k)) else k
                cv = {}
                for vk, vv in v.items():
                    cvk = bool(vk) if isinstance(vk, (bool,)) or 'bool_' in str(type(vk)) else vk
                    cv[cvk] = int(vv)
                clean[ck] = cv
            stats['metro_rural_crosstab'] = clean

        for attr, key in [('age_at_start_years','age_stats'),
                           ('programme_duration_months','duration_stats')]:
            if attr in self.df.columns:
                d = self.df[attr].describe()
                stats[key] = {k: round(float(d[k]),2) if k!='count' else int(d[k])
                               for k in ['count','mean','std','min','25%','50%','75%','max']}

        id_col   = find_col('ID_Number') or find_col('Passport_Number')
        name_col = find_col('Full_Names') or find_col('Surname')
        valid = 0
        for idx in range(len(self.df)):
            for col in [id_col, name_col]:
                if col and pd.notna(self.df.at[idx,col]) and str(self.df.at[idx,col]).strip():
                    valid += 1; break
        stats['total_learners']    = valid
        stats['report_generated']  = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        self.summary_stats = stats
        return self

    def create_charts(self):
        charts = []

        def save_chart(name, fig):
            path = self.charts_dir / f"{name}.png"
            fig.savefig(path, dpi=150, bbox_inches='tight', facecolor='white')
            plt.close(fig)
            charts.append(str(path))

        if 'learner_status' in self.summary_stats:
            fig, ax = plt.subplots(figsize=(8,5))
            data = pd.Series(self.summary_stats['learner_status']).sort_values(ascending=False)
            data.plot(kind='bar', ax=ax, color=sns.color_palette("husl",len(data)))
            ax.set_title('Learner Status Distribution', fontsize=14, fontweight='bold')
            ax.set_xlabel('Status'); ax.set_ylabel('Count')
            plt.xticks(rotation=45,ha='right'); plt.tight_layout()
            save_chart('learner_status', fig)

        if 'gender' in self.summary_stats:
            fig, ax = plt.subplots(figsize=(6,6))
            data = pd.Series(self.summary_stats['gender'])
            ax.pie(data.values, labels=data.index, autopct='%1.1f%%',
                   colors=sns.color_palette("pastel",len(data)), startangle=90)
            ax.set_title('Gender Distribution', fontsize=14, fontweight='bold')
            save_chart('gender_distribution', fig)

        if 'population_group' in self.summary_stats:
            fig, ax = plt.subplots(figsize=(10,5))
            data = pd.Series(self.summary_stats['population_group']).sort_values(ascending=True)
            data.plot(kind='barh', ax=ax, color=sns.color_palette("muted",len(data)))
            ax.set_title('Population Group Distribution', fontsize=14, fontweight='bold')
            ax.set_xlabel('Count'); plt.tight_layout()
            save_chart('population_group', fig)

        if 'top_programmes_wil' in self.summary_stats:
            fig, ax = plt.subplots(figsize=(12,7))
            data = pd.Series(self.summary_stats['top_programmes_wil']).head(15).sort_values(ascending=True)
            data.plot(kind='barh', ax=ax, color=sns.color_palette("viridis",len(data)))
            ax.set_title('WIL Learning Programmes', fontsize=14, fontweight='bold')
            ax.set_xlabel('Number of Learners'); plt.tight_layout()
            save_chart('top_programmes_wil', fig)

        if 'top_programmes_graduates' in self.summary_stats:
            fig, ax = plt.subplots(figsize=(12,7))
            data = pd.Series(self.summary_stats['top_programmes_graduates']).head(15).sort_values(ascending=True)
            data.plot(kind='barh', ax=ax, color=sns.color_palette("plasma",len(data)))
            ax.set_title('Graduate Learning Programmes', fontsize=14, fontweight='bold')
            ax.set_xlabel('Number of Learners'); plt.tight_layout()
            save_chart('top_programmes_graduates', fig)

        if 'top_sectors' in self.summary_stats:
            fig, ax = plt.subplots(figsize=(10,6))
            data = pd.Series(self.summary_stats['top_sectors']).sort_values(ascending=True)
            data.plot(kind='barh', ax=ax, color=sns.color_palette("coolwarm",len(data)))
            ax.set_title('Economic Sectors', fontsize=14, fontweight='bold')
            ax.set_xlabel('Number of Learners'); ax.set_ylabel('Sector')
            plt.tight_layout()
            save_chart('top_sectors', fig)

        if 'residential_province' in self.summary_stats:
            fig, ax = plt.subplots(figsize=(10,6))
            data = pd.Series(self.summary_stats['residential_province']).sort_values(ascending=False)
            data.plot(kind='bar', ax=ax, color=sns.color_palette("Set2",len(data)))
            ax.set_title('Residential Province Distribution', fontsize=14, fontweight='bold')
            ax.set_xlabel('Province'); ax.set_ylabel('Count')
            plt.xticks(rotation=45,ha='right'); plt.tight_layout()
            save_chart('residential_province', fig)

        if 'age_at_start_years' in self.df.columns:
            fig, ax = plt.subplots(figsize=(10,5))
            col = self.df['age_at_start_years'].dropna()
            col.plot(kind='hist', bins=30, ax=ax, color='steelblue', edgecolor='black', alpha=0.7)
            ax.axvline(col.mean(), color='red', linestyle='--', linewidth=2,
                       label=f"Mean: {col.mean():.1f}")
            ax.set_title('Age Distribution at Programme Start', fontsize=14, fontweight='bold')
            ax.set_xlabel('Age (Years)'); ax.set_ylabel('Frequency'); ax.legend()
            plt.tight_layout()
            save_chart('age_distribution', fig)

        if 'programme_duration_months' in self.df.columns:
            fig, ax = plt.subplots(figsize=(10,5))
            col = self.df['programme_duration_months'].dropna()
            col.plot(kind='hist', bins=30, ax=ax, color='seagreen', edgecolor='black', alpha=0.7)
            ax.axvline(col.mean(), color='red', linestyle='--', linewidth=2,
                       label=f"Mean: {col.mean():.1f}")
            ax.set_title('Programme Duration Distribution', fontsize=14, fontweight='bold')
            ax.set_xlabel('Duration (Months)'); ax.set_ylabel('Frequency'); ax.legend()
            plt.tight_layout()
            save_chart('duration_distribution', fig)

        if self.summary_stats.get('disability'):
            data = pd.Series(self.summary_stats['disability'])
            data = data[data > 0].sort_values(ascending=True)
            if len(data) > 0:
                data.index = [idx.replace('_', ' ') for idx in data.index]
                fig, ax = plt.subplots(figsize=(10,6))
                data.plot(kind='barh', ax=ax, color=sns.color_palette("rocket",len(data)))
                ax.set_title('Disability Types (Learners Affected)', fontsize=14, fontweight='bold')
                ax.set_xlabel('Count'); plt.tight_layout()
                save_chart('disability_types', fig)

        if 'is_metro' in self.df.columns and 'is_rural' in self.df.columns:
            ct = pd.crosstab(self.df['is_metro'], self.df['is_rural'])
            if not ct.empty:
                ct.index   = ['Not in Metro' if x is False else 'In Metro'   for x in ct.index]
                ct.columns = ['Not Rural'    if x is False else 'Is Rural'   for x in ct.columns]
                fig, ax = plt.subplots(figsize=(8,5))
                ct.plot(kind='bar', ax=ax, color=['coral','lightblue'])
                ax.set_title('Metro and Rural Classification', fontsize=14, fontweight='bold')
                ax.set_xlabel('Metro Classification'); ax.set_ylabel('Count')
                ax.legend(title='Rural Classification'); plt.xticks(rotation=0); plt.tight_layout()
                save_chart('metro_rural', fig)

        hd = {}
        if 'high_demand_2020' in self.summary_stats: hd['2020 List'] = self.summary_stats['high_demand_2020']
        if 'high_demand_2024' in self.summary_stats: hd['2024 List'] = self.summary_stats['high_demand_2024']
        if hd:
            fig, ax = plt.subplots(figsize=(6,5))
            pd.Series(hd).plot(kind='bar', ax=ax, color=['#2ecc71','#3498db'])
            ax.set_title('High-Demand Occupations', fontsize=14, fontweight='bold')
            ax.set_ylabel('Number of Learners'); plt.xticks(rotation=0); plt.tight_layout()
            save_chart('high_demand', fig)

        self.charts = charts
        print(f"  Created {len(charts)} charts")
        return self

    def generate_html_report(self) -> Path:
        html_template = """
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Learner Data Analytics Report</title>
<style>
  * { margin:0; padding:0; box-sizing:border-box; }
  body { font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif; background:#f5f7fa; color:#333; line-height:1.6; }
  .container { max-width:1100px; margin:0 auto; padding:20px; }
  header { background:linear-gradient(135deg,#67c7f1 0%,#2E75B6 100%); color:white; padding:40px 20px;
           text-align:center; border-radius:10px; margin-bottom:30px; }
  header h1 { font-size:2.2em; margin-bottom:8px; }
  header p  { font-size:1.05em; opacity:.9; }
  .stats-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(220px,1fr)); gap:20px; margin-bottom:30px; }
  .stat-card { background:white; padding:22px; border-radius:10px; box-shadow:0 2px 4px rgba(0,0,0,.1);
               border-left:4px solid #67c7f1; }
  .stat-card h3 { color:#67c7f1; font-size:.85em; text-transform:uppercase; letter-spacing:1px; margin-bottom:8px; }
  .stat-card .value { font-size:1.9em; font-weight:bold; }
  .section { background:white; padding:28px; border-radius:10px; margin-bottom:28px;
             box-shadow:0 2px 4px rgba(0,0,0,.1); }
  .section h2 { color:#67c7f1; font-size:1.4em; margin-bottom:18px; padding-bottom:8px;
                border-bottom:2px solid #f0f0f0; }
  table { width:100%; border-collapse:collapse; margin-top:14px; }
  th,td { padding:10px 12px; text-align:left; border-bottom:1px solid #eee; }
  th { background:#f8f9fa; font-weight:600; color:#67c7f1; }
  tr:hover { background:#f8f9fa; }
  .chart-container { text-align:center; margin:18px 0; }
  .chart-container img { max-width:100%; height:auto; border-radius:8px;
                          box-shadow:0 2px 8px rgba(0,0,0,.1); }
  .two-col { display:grid; grid-template-columns:1fr 1fr; gap:20px; }
  footer { text-align:center; padding:20px; color:#999; font-size:.85em; }
</style>
</head>
<body>
<div class="container">
  <header>
    <h1>Learner Data Analytics Report</h1>
    <p>Comprehensive analysis of learner enrolment data</p>
    <p style="margin-top:8px;font-size:.88em;">Generated: {{ stats.report_generated }}</p>
  </header>

  <div class="stats-grid">
    <div class="stat-card"><h3>Total Learners</h3><div class="value">{{ stats.total_learners }}</div></div>
    {% if stats.age_stats %}<div class="stat-card"><h3>Avg Age at Start</h3><div class="value">{{ stats.age_stats.mean }} yrs</div></div>{% endif %}
    {% if stats.duration_stats %}<div class="stat-card"><h3>Avg Programme Duration</h3><div class="value">{{ stats.duration_stats.mean }} mo</div></div>{% endif %}
    {% if stats.high_demand_2024 is defined %}<div class="stat-card"><h3>High Demand 2024</h3><div class="value">{{ stats.high_demand_2024 }}</div></div>{% endif %}
  </div>

  {% if stats.learners_per_site %}
  <div class="section">
    <h2>Host Employers</h2>
    <table>
      <tr><th>Rank</th><th>Site Name</th><th>Learners</th></tr>
      {% for site, count in stats.learners_per_site.items() %}
      <tr><td>{{ loop.index }}</td><td>{{ site }}</td><td>{{ count }}</td></tr>
      {% endfor %}
    </table>
  </div>
  {% endif %}

  <div class="two-col">
    {% if stats.top_programmes %}<div class="section">
      <h2>Top Learning Programmes</h2>
      <div class="chart-container"><img src="{{ chart_dir }}/top_programmes.png" alt="Top Programmes"></div>
      <table><tr><th>Rank</th><th>Programme</th><th>Learners</th></tr>
      {% for p,c in stats.top_programmes.items() %}<tr><td>{{ loop.index }}</td><td>{{ p }}</td><td>{{ c }}</td></tr>{% endfor %}
      </table>
    </div>{% endif %}
    {% if stats.top_sectors %}<div class="section">
      <h2>Top Economic Sectors</h2>
      <div class="chart-container"><img src="{{ chart_dir }}/top_sectors.png" alt="Top Sectors"></div>
      <table><tr><th>Rank</th><th>Sector</th><th>Learners</th></tr>
      {% for s,c in stats.top_sectors.items() %}<tr><td>{{ loop.index }}</td><td>{{ s }}</td><td>{{ c }}</td></tr>{% endfor %}
      </table>
    </div>{% endif %}
  </div>

  {% if stats.learner_status %}<div class="section">
    <h2>Learner Status Distribution</h2>
    <div class="chart-container"><img src="{{ chart_dir }}/learner_status.png" alt="Learner Status"></div>
    <table><tr><th>Status</th><th>Count</th><th>Percentage</th></tr>
    {% set t = stats.learner_status.values()|sum %}
    {% for s,c in stats.learner_status.items() %}<tr><td>{{ s }}</td><td>{{ c }}</td><td>{{ "%.1f"|format((c/t)*100) }}%</td></tr>{% endfor %}
    </table>
  </div>{% endif %}

  {% if stats.gender %}<div class="section">
    <h2>Gender Distribution</h2>
    <div class="chart-container"><img src="{{ chart_dir }}/gender_distribution.png" alt="Gender"></div>
    <table><tr><th>Gender</th><th>Count</th><th>%</th></tr>
    {% set t = stats.gender.values()|sum %}
    {% for g,c in stats.gender.items() %}<tr><td>{{ g }}</td><td>{{ c }}</td><td>{{ "%.1f"|format((c/t)*100) }}%</td></tr>{% endfor %}
    </table>
  </div>{% endif %}

  {% if stats.population_group %}<div class="section">
    <h2>Population Group</h2>
    <div class="chart-container"><img src="{{ chart_dir }}/population_group.png" alt="Population Group"></div>
    <table><tr><th>Group</th><th>Count</th><th>%</th></tr>
    {% set t = stats.population_group.values()|sum %}
    {% for g,c in stats.population_group.items() %}<tr><td>{{ g }}</td><td>{{ c }}</td><td>{{ "%.1f"|format((c/t)*100) }}%</td></tr>{% endfor %}
    </table>
  </div>{% endif %}

  {% if stats.household_language %}<div class="section">
    <h2>Household Language Distribution</h2>
    <table><tr><th>Language</th><th>Count</th><th>Percentage</th></tr>
    {% set t = stats.household_language.values()|sum %}
    {% for l,c in stats.household_language.items() %}<tr><td>{{ l }}</td><td>{{ c }}</td><td>{{ "%.1f"|format((c/t)*100) }}%</td></tr>{% endfor %}
    </table>
  </div>{% endif %}

  {% if stats.disability %}<div class="section">
    <h2>Disability Types</h2>
    {% set active = stats.disability.values()|select('gt',0)|list %}
    {% if active|length > 0 %}
    <div class="chart-container"><img src="{{ chart_dir }}/disability_types.png" alt="Disabilities"></div>
    <table><tr><th>Disability Type</th><th>Affected Learners</th></tr>
    {% for d,c in stats.disability.items() %}{% if c > 0 %}<tr><td>{{ d.replace('_',' ').title() }}</td><td>{{ c }}</td></tr>{% endif %}{% endfor %}
    </table>
    {% else %}<p style="color:#888;font-style:italic;padding:20px;">No disability data recorded.</p>{% endif %}
  </div>{% endif %}

  {% if stats.high_demand_2020 is defined or stats.high_demand_2024 is defined %}<div class="section">
    <h2>High-Demand Occupations</h2>
    <div class="chart-container"><img src="{{ chart_dir }}/high_demand.png" alt="High Demand"></div>
    <div class="stats-grid">
      {% if stats.high_demand_2020 is defined %}<div class="stat-card"><h3>2020 High-Demand List</h3><div class="value">{{ stats.high_demand_2020 }}</div></div>{% endif %}
      {% if stats.high_demand_2024 is defined %}<div class="stat-card"><h3>2024 High-Demand List</h3><div class="value">{{ stats.high_demand_2024 }}</div></div>{% endif %}
    </div>
  </div>{% endif %}

  <div class="two-col">
    {% if stats.residential_province %}<div class="section">
      <h2>Residential Province</h2>
      <div class="chart-container"><img src="{{ chart_dir }}/residential_province.png" alt="Residential Province"></div>
      <table><tr><th>Province</th><th>Count</th><th>%</th></tr>
      {% set t = stats.residential_province.values()|sum %}
      {% for p,c in stats.residential_province.items() %}<tr><td>{{ p }}</td><td>{{ c }}</td><td>{{ "%.1f"|format((c/t)*100) }}%</td></tr>{% endfor %}
      </table>
    </div>{% endif %}
    {% if stats.site_province %}<div class="section">
      <h2>Learning Site Province</h2>
      <table><tr><th>Province</th><th>Count</th><th>%</th></tr>
      {% set t = stats.site_province.values()|sum %}
      {% for p,c in stats.site_province.items() %}<tr><td>{{ p }}</td><td>{{ c }}</td><td>{{ "%.1f"|format((c/t)*100) }}%</td></tr>{% endfor %}
      </table>
    </div>{% endif %}
  </div>

  {% if stats.metro_rural_crosstab %}<div class="section">
    <h2>Metro and Rural Classification</h2>
    <div class="chart-container"><img src="{{ chart_dir }}/metro_rural.png" alt="Metro Rural"></div>
    <table><tr><th>Metro / Rural</th><th>Not Rural</th><th>Is Rural</th></tr>
    {% for key,vals in stats.metro_rural_crosstab.items() %}{% if key != 'All' %}
    <tr><td>{% if key == True %}In Metro{% elif key == False %}Not in Metro{% else %}{{ key }}{% endif %}</td>
        <td>{{ vals.get(False,0) }}</td><td>{{ vals.get(True,0) }}</td></tr>
    {% endif %}{% endfor %}
    </table>
  </div>{% endif %}

  {% if stats.age_stats %}<div class="section">
    <h2>Age Statistics (at Programme Start)</h2>
    <div class="chart-container"><img src="{{ chart_dir }}/age_distribution.png" alt="Age Distribution"></div>
    <div class="stats-grid">
      <div class="stat-card"><h3>Count</h3><div class="value">{{ stats.age_stats.count }}</div></div>
      <div class="stat-card"><h3>Mean</h3><div class="value">{{ stats.age_stats.mean }}</div></div>
      <div class="stat-card"><h3>Std Dev</h3><div class="value">{{ stats.age_stats.std }}</div></div>
      <div class="stat-card"><h3>Median</h3><div class="value">{{ stats.age_stats['50%'] }}</div></div>
    </div>
  </div>{% endif %}

  {% if stats.duration_stats %}<div class="section">
    <h2>Programme Duration Statistics</h2>
    <div class="chart-container"><img src="{{ chart_dir }}/duration_distribution.png" alt="Duration Distribution"></div>
    <div class="stats-grid">
      <div class="stat-card"><h3>Count</h3><div class="value">{{ stats.duration_stats.count }}</div></div>
      <div class="stat-card"><h3>Mean</h3><div class="value">{{ stats.duration_stats.mean }}</div></div>
      <div class="stat-card"><h3>Min</h3><div class="value">{{ stats.duration_stats.min }}</div></div>
      <div class="stat-card"><h3>Max</h3><div class="value">{{ stats.duration_stats.max }}</div></div>
    </div>
  </div>{% endif %}

  <footer>Learner Data Analytics System | Generated on {{ stats.report_generated }}</footer>
</div>
</body>
</html>
"""
        template     = Template(html_template)
        html_content = template.render(stats=self.summary_stats,
                                       chart_dir=str(self.charts_dir))
        report_path  = self.output_dir / "learner_report.html"
        report_path.write_text(html_content, encoding='utf-8')
        print(f"  HTML report saved: {report_path}")
        return report_path

    def run(self) -> dict:
        print("\n" + "="*60)
        print("  LEARNER DATA ANALYTICS")
        print("="*60)
        self.load_data()
        self.clean_column_names()
        self.create_derived_fields()
        self.generate_summary_statistics()
        self.create_charts()
        html_path = self.generate_html_report()

        pdf_path = self.output_dir / "learner_report.pdf"
        pdf_ok   = build_pdf(pdf_path,
                             "Learner Data Analytics Report",
                             self.summary_stats,
                             self.charts_dir)

        print("\n  ✔  Learner report complete!")
        return {"html": html_path, "pdf": pdf_path if pdf_ok else None}


# Combined Report System

class CombinedReportSystem:
    """Orchestrates both Learner and Dropouts reports into a single output."""

    def __init__(self, learner_path: Path, dropout_path: Path, output_dir: Path):
        self.learner_path = learner_path
        self.dropout_path = dropout_path
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.charts_dir = self.output_dir / "charts"
        self.charts_dir.mkdir(parents=True, exist_ok=True)
        self.combined_stats = {}

    def _embed_images_as_base64(self, html_content: str) -> str:
        """Replace local image src paths with base64 data URIs."""
        import re, base64, mimetypes

        def replace_src(match):
            src = match.group(1)
            # Skip already-embedded data URIs and remote URLs
            if src.startswith('data:') or src.startswith('http'):
                return match.group(0)
            # Try path relative to output_dir first, then cwd
            for base in (self.output_dir, Path('.')):
                img_path = base / src
                if img_path.exists():
                    mime, _ = mimetypes.guess_type(str(img_path))
                    mime = mime or 'image/png'
                    data = base64.b64encode(img_path.read_bytes()).decode('ascii')
                    return f'src="data:{mime};base64,{data}"'
            # If not found, leave as-is
            return match.group(0)

        return re.sub(r'src="([^"]+)"', replace_src, html_content)

    def run(self) -> dict:
        print("" + "="*60)
        print("  COMBINED ANALYTICS REPORT")
        print("="*60)

        import shutil

        # Learner Data
        print("  [1/2] Processing Learner Data...")
        learner_temp = self.output_dir / "_temp_learner"
        learner_temp.mkdir(parents=True, exist_ok=True)
        learner = LearnerReportSystem(self.learner_path, learner_temp)
        learner.load_data()
        learner.clean_column_names()
        learner.create_derived_fields()
        learner.generate_summary_statistics()
        learner.create_charts()

        for src in learner.charts_dir.glob("*.png"):
            shutil.copy2(src, self.charts_dir / src.name)

        # Dropouts
        print("  [2/2] Processing Dropouts Master List...")
        dropout_temp = self.output_dir / "_temp_dropouts"
        dropout_temp.mkdir(parents=True, exist_ok=True)
        dropout = DropoutsReportSystem(self.dropout_path, dropout_temp)
        dropout.load_data()
        dropout.clean_column_names()
        dropout.create_derived_fields()
        dropout.generate_summary_statistics()
        dropout.create_charts()

        for src in dropout.charts_dir.glob("*.png"):
            shutil.copy2(src, self.charts_dir / src.name)

        # Merge stats
        self.combined_stats = {}
        self.combined_stats.update(learner.summary_stats)
        self.combined_stats.update(dropout.summary_stats)
        self.combined_stats['report_generated'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        # Rotate logo and copy to output dir
        logo_src = Path("LeadHR Logo.png")
        logo_dest = self.output_dir / "LeadHR Logo.png"
        if logo_src.exists():
            try:
                from PIL import Image as PilImage
                img = PilImage.open(logo_src)
                img = img.rotate(90, expand=True)
                img.save(logo_dest)
            except ImportError:
                shutil.copy2(logo_src, logo_dest)
        elif logo_dest != logo_src and logo_src.exists():
            shutil.copy2(logo_src, logo_dest)

        # Copy story images to output dir so they resolve correctly
        story_images = [
            "Youth Empowerment Conference 1.png",
            "Youth Empowerment Conference 2.png",
            "Youth Empowerment Conference 3.png",
            "Youth Empowerment Conference 4.png",
            "Zanele 1.png",
            "Zanele 2.png",
        ]
        for img_name in story_images:
            src = Path(img_name)
            if src.exists():
                shutil.copy2(src, self.output_dir / img_name)

        # Generate HTML (normal version for browser viewing)
        html_path, embedded_path = self._generate_html_report()

        # Generate PDF using Brave/Chrome headless
        pdf_path = self.output_dir / "combined_report.pdf"
        pdf_ok = False
        try:
            import subprocess

            browser_paths = [
                r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe",
                r"C:\Program Files (x86)\BraveSoftware\Brave-Browser\Application\brave.exe",
                r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
                r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
            ]

            browser_exe = next((p for p in browser_paths if Path(p).exists()), None)

            if browser_exe:
                print(f"  Generating PDF with: {Path(browser_exe).name}")
                subprocess.run([
                    browser_exe,
                    "--headless=new",
                    "--disable-gpu",
                    "--no-sandbox",
                    "--run-all-compositor-stages-before-draw",
                    "--disable-extensions",
                    f"--print-to-pdf={pdf_path}",
                    "--print-to-pdf-no-header",
                    "--no-pdf-header-footer",
                    str(embedded_path)
                ], check=True, capture_output=True, timeout=60)
                print(f"  PDF saved: {pdf_path}")
                pdf_ok = True
            else:
                print("  ⚠  No supported browser found for PDF generation")

        except Exception as e:
            print(f"  ⚠  PDF generation failed: {e}")

        # Clean up the temporary embedded HTML
        embedded_path.unlink(missing_ok=True)

        # Clean up temp dirs
        shutil.rmtree(learner_temp, ignore_errors=True)
        shutil.rmtree(dropout_temp, ignore_errors=True)

        print("  ✔  Combined report complete!")
        print(f"     HTML: {html_path}")
        print(f"     PDF:  {pdf_path if pdf_ok else 'PDF generation skipped'}")
        return {"html": html_path, "pdf": pdf_path if pdf_ok else None}

    def _generate_html_report(self):
        html_template = """
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>NSF Analytics Report</title>
<style>
  * { margin:0; padding:0; box-sizing:border-box; }
  body { font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif; background:#f5f7fa; color:#333; line-height:1.6; }
  .container { max-width:1200px; margin:0 auto; padding:20px; }
  header { background:linear-gradient(135deg,#67c7f1 0%,#4bbaea 100%); color:white; padding:40px 20px;
           text-align:center; border-radius:10px; margin-bottom:30px; }
  header h1 { font-size:2.4em; margin-bottom:8px; }
  header p  { font-size:1.1em; opacity:.9; }
  .stats-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(220px,1fr)); gap:20px; margin-bottom:30px; }
  .stat-card { background:white; padding:22px; border-radius:10px; box-shadow:0 2px 4px rgba(0,0,0,.1);
               border-left:4px solid #67c7f1; }
  .stat-card h3 { color:#67c7f1; font-size:.85em; text-transform:uppercase; letter-spacing:1px; margin-bottom:8px; }
  .stat-card .value { font-size:1.9em; font-weight:bold; }
  .section { background:white; padding:28px; border-radius:10px; margin-bottom:28px;
             box-shadow:0 2px 4px rgba(0,0,0,.1); }
  .section h2 { color:#67c7f1; font-size:1.4em; margin-bottom:18px; padding-bottom:8px;
                border-bottom:2px solid #f0f0f0; }
  table { width:100%; border-collapse:collapse; margin-top:14px; }
  th,td { padding:10px 12px; text-align:left; border-bottom:1px solid #eee; }
  th { background:#f8f9fa; font-weight:600; color:#4bbaea; }
  tr:hover { background:#f8f9fa; }
  .chart-container { text-align:center; margin:18px 0; }
  .chart-container img { max-width:100%; height:auto; border-radius:8px;
                          box-shadow:0 2px 8px rgba(0,0,0,.1); }
  .two-col { display:grid; grid-template-columns:1fr 1fr; gap:20px; }
  @media (max-width:768px) { .two-col { grid-template-columns:1fr; } }
  .story-block { display:grid; grid-template-columns:2fr 1fr; gap:30px; align-items:stretch; }
  @media (max-width:768px) { .story-block { grid-template-columns:1fr; } }
  .story-text h3 { color:#67c7f1; font-size:1.15em; margin-bottom:12px; }
  .story-text p { margin-bottom:12px; text-align:justify; }
  .story-text ul { margin-left:20px; margin-bottom:12px; }
  .story-text li { margin-bottom:6px; }
  .story-images { display:flex; flex-direction:column; height:100%; }
  .img-stack { display:flex; flex-direction:column; gap:8px; width:100%; height:100%; }
  .img-stack img { width:100%; flex:1 1 0; min-height:0; object-fit:cover; border-radius:8px; }
  footer { text-align:center; padding:20px; color:#999; font-size:.85em; }
  .none-msg { color:#888; font-style:italic; padding:20px; }
</style>
</head>
<body>
<div class="container">
  <header style="position:relative; padding:40px 20px; text-align:center;">
    <div style="position:absolute; left:0; top:0; bottom:0; background:#ffe100; display:flex; align-items:center; justify-content:center; padding:0 20px; border-radius:10px 0 0 10px;">
      <img src="LeadHR Logo.png" alt="Lead HR Logo"
        style="height:110px; width:auto; display:block; object-fit:contain;">
    </div>
    <h1>NSF Analytics Report</h1>
    <p>Combined Learner Data and Dropouts Analysis</p>
    <p style="margin-top:8px;font-size:.88em;">Generated: {{ stats.report_generated }}</p>
  </header>

  <!-- Key Metrics -->
  <div class="stats-grid">
    <div class="stat-card"><h3>Total Learners</h3><div class="value">{{ stats.total_learners }}</div></div>
    <div class="stat-card"><h3>Total Dropouts</h3><div class="value">{{ stats.total_dropouts }}</div></div>
    {% if stats.recent_dropouts_12mo is defined %}
    <div class="stat-card"><h3>Dropouts (Last 12 Mo)</h3><div class="value">{{ stats.recent_dropouts_12mo }}</div></div>
    {% endif %}
    {% if stats.age_stats %}<div class="stat-card"><h3>Avg Age at Start</h3><div class="value">{{ stats.age_stats.mean }} yrs</div></div>{% endif %}
    {% if stats.duration_stats %}<div class="stat-card"><h3>Avg Programme Duration</h3><div class="value">{{ stats.duration_stats.mean }} mo</div></div>{% endif %}
    {% if stats.high_demand_2024 is defined %}<div class="stat-card"><h3>High Demand 2024</h3><div class="value">{{ stats.high_demand_2024 }}</div></div>{% endif %}
  </div>

  <!-- Success Stories -->
  <div class="section">
    <h2>Success Stories</h2>

    <div class="story-block">
      <div class="story-text">
        <h3>Youth Empowerment Conference</h3>
        <p>The Wentworth Youth Centre buzzed with energy as it hosted the impactful Youth Empowerment Conference. This vital initiative was born from the passionate vision of Nonelwa, one of our dedicated interns, whose desire to effect positive change in Ward 66 truly inspired us. Working hand-in-hand with Councillor Zoe and a committed team of individuals, Nonelwa's dream transformed into a tangible reality, ensuring the conference was a resounding success. Over two dynamic days, the conference provided a crucial platform for the unemployed youth of Ward 66. Various organizations stepped forward, offering invaluable guidance on a multitude of pathways to self-improvement and opportunity.</p>
        <p>Attendees gained insights into:</p>
        <ul>
          <li><b>Acquiring Employment:</b> Practical advice and strategies for job searching.</li>
          <li><b>Internships &amp; Learnerships:</b> Information on gaining vital work experience.</li>
          <li><b>Education &amp; Skills Development:</b> Opportunities for further learning and acquiring new, in-demand skills.</li>
        </ul>
        <p>The overwhelming participation and positive feedback from the youth underscored the critical need for such initiatives. It was a powerful demonstration of community collaboration, igniting hope and providing concrete steps for a brighter future for the young people of Ward 66. We're incredibly proud of Nonelwa's leadership and the collective effort that made this event possible!</p>
      </div>
      <div class="story-images">
        <div class="img-stack">
          <img src="Youth Empowerment Conference 1.png" alt="Youth Empowerment Conference 1">
          <img src="Youth Empowerment Conference 2.png" alt="Youth Empowerment Conference 2">
          <img src="Youth Empowerment Conference 3.png" alt="Youth Empowerment Conference 3">
          <img src="Youth Empowerment Conference 4.png" alt="Youth Empowerment Conference 4">
        </div>
      </div>
    </div>

    <div class="story-block" style="margin-top:40px;">
      <div class="story-text">
        <h3>Spotlight on Zanele Mhlongo</h3>
        <p>We're incredibly proud to shine a spotlight on Zanele Mhlongo, an outstanding intern from our NSF WIL placement programme. Zanele has consistently developed and added immense value to those around her. Her insightful articles, which have graced the pages of various newspapers, stand as a true testament to her exceptional skill and dedication. She has fully integrated into her environment, pushing the boundaries of what it means to be an 'intern'.</p>
        <p style="font-style:italic; border-left:3px solid #67c7f1; padding-left:15px; margin:15px 0; color:#555;">"I feel like I'm not expressing my gratitude fully. I would like to take this opportunity to thank you for the chance you have given me. I'm absolutely loving the experience I'm gaining here. One of the highlights is seeing my work featured in prominent publications like Metro Ezasegagasini, Newsflashes and Workplace magazine. These contributions are a testament to the skills I'm developing and the value I'm adding. KEEP UP THE GOOD WORK"</p>
        <p><b>Zanele.</b></p>
      </div>
      <div class="story-images">
        <div class="img-stack">
          <img src="Zanele 1.png" alt="Zanele 1">
          <img src="Zanele 2.png" alt="Zanele 2">
        </div>
      </div>
    </div>
  </div>

  <!-- Host Employers -->
  {% if stats.learners_per_site %}<div class="section">
    <h2>Host Employers</h2>
    <table>
      <tr><th>Rank</th><th>Site Name</th><th>Learners</th></tr>
      {% for site, count in stats.learners_per_site.items() %}
      <tr><td>{{ loop.index }}</td><td>{{ site }}</td><td>{{ count }}</td></tr>
      {% endfor %}
    </table>
  </div>{% endif %}

  <!-- Top WIL Programmes -->
  {% if stats.top_programmes_wil %}<div class="section">
    <h2>Top WIL Learning Programmes</h2>
    <div class="chart-container"><img src="charts/top_programmes_wil.png" alt="Top WIL Programmes"></div>
    <table>
      <tr><th>Rank</th><th>Programme</th><th>Learners</th></tr>
      {% for p,c in stats.top_programmes_wil.items() %}<tr><td>{{ loop.index }}</td><td>{{ p }}</td><td>{{ c }}</td></tr>{% endfor %}
    </table>
  </div>{% endif %}

  <!-- Top Graduate Programmes -->
  {% if stats.top_programmes_graduates %}<div class="section">
    <h2>Top Graduate Learning Programmes</h2>
    <div class="chart-container"><img src="charts/top_programmes_graduates.png" alt="Top Graduate Programmes"></div>
    <table>
      <tr><th>Rank</th><th>Programme</th><th>Learners</th></tr>
      {% for p,c in stats.top_programmes_graduates.items() %}<tr><td>{{ loop.index }}</td><td>{{ p }}</td><td>{{ c }}</td></tr>{% endfor %}
    </table>
  </div>{% endif %}

  <!-- Top Economic Sectors -->
  {% if stats.top_sectors %}<div class="section">
    <h2>Top Economic Sectors</h2>
    <div class="chart-container"><img src="charts/top_sectors.png" alt="Top Sectors"></div>
    <table>
      <tr><th>Rank</th><th>Sector</th><th>Learners</th></tr>
      {% for s,c in stats.top_sectors.items() %}<tr><td>{{ loop.index }}</td><td>{{ s }}</td><td>{{ c }}</td></tr>{% endfor %}
    </table>
  </div>{% endif %}

  <!-- Learner Status -->
  {% if stats.learner_status %}<div class="section">
    <h2>Learner Status Distribution</h2>
    <div class="chart-container"><img src="charts/learner_status.png" alt="Learner Status"></div>
    <table>
      <tr><th>Status</th><th>Count</th><th>Percentage</th></tr>
      {% set t = stats.learner_status.values()|sum %}
      {% for s,c in stats.learner_status.items() %}<tr><td>{{ s }}</td><td>{{ c }}</td><td>{{ "%.1f"|format((c/t)*100) }}%</td></tr>{% endfor %}
    </table>
  </div>{% endif %}

  <!-- Gender Distribution -->
  {% if stats.gender %}<div class="section">
    <h2>Gender Distribution</h2>
    <div class="chart-container"><img src="charts/gender_distribution.png" alt="Gender"></div>
    <table>
      <tr><th>Gender</th><th>Count</th><th>%</th></tr>
      {% set t = stats.gender.values()|sum %}
      {% for g,c in stats.gender.items() %}<tr><td>{{ g }}</td><td>{{ c }}</td><td>{{ "%.1f"|format((c/t)*100) }}%</td></tr>{% endfor %}
    </table>
  </div>{% endif %}

  <!-- Population Group -->
  {% if stats.population_group %}<div class="section">
    <h2>Population Group</h2>
    <div class="chart-container"><img src="charts/population_group.png" alt="Population Group"></div>
    <table>
      <tr><th>Group</th><th>Count</th><th>%</th></tr>
      {% set t = stats.population_group.values()|sum %}
      {% for g,c in stats.population_group.items() %}<tr><td>{{ g }}</td><td>{{ c }}</td><td>{{ "%.1f"|format((c/t)*100) }}%</td></tr>{% endfor %}
    </table>
  </div>{% endif %}

  <!-- Household Language -->
  {% if stats.household_language %}<div class="section">
    <h2>Household Language Distribution</h2>
    <table>
      <tr><th>Language</th><th>Count</th><th>Percentage</th></tr>
      {% set t = stats.household_language.values()|sum %}
      {% for l,c in stats.household_language.items() %}<tr><td>{{ l }}</td><td>{{ c }}</td><td>{{ "%.1f"|format((c/t)*100) }}%</td></tr>{% endfor %}
    </table>
  </div>{% endif %}

  <!-- Disability Types -->
  {% if stats.disability %}<div class="section">
    <h2>Disability Types</h2>
    {% set active = stats.disability.values()|select('gt',0)|list %}
    {% if active|length > 0 %}
    <div class="chart-container"><img src="charts/disability_types.png" alt="Disabilities"></div>
    <table>
      <tr><th>Disability Type</th><th>Affected Learners</th></tr>
      {% for d,c in stats.disability.items() %}{% if c > 0 %}<tr><td>{{ d.replace('_',' ').title() }}</td><td>{{ c }}</td></tr>{% endif %}{% endfor %}
    </table>
    {% else %}<p class="none-msg">No disability data recorded.</p>{% endif %}
  </div>{% endif %}

  <!-- High-Demand Occupations -->
  {% if stats.high_demand_2020 is defined or stats.high_demand_2024 is defined %}<div class="section">
    <h2>High-Demand Occupations</h2>
    <div class="chart-container"><img src="charts/high_demand.png" alt="High Demand"></div>
    <div class="stats-grid">
      {% if stats.high_demand_2020 is defined %}<div class="stat-card"><h3>2020 High-Demand List</h3><div class="value">{{ stats.high_demand_2020 }}</div></div>{% endif %}
      {% if stats.high_demand_2024 is defined %}<div class="stat-card"><h3>2024 High-Demand List</h3><div class="value">{{ stats.high_demand_2024 }}</div></div>{% endif %}
    </div>
  </div>{% endif %}

  <!-- Provinces -->
  <div class="two-col">
    {% if stats.residential_province %}<div class="section">
      <h2>Residential Province</h2>
      <div class="chart-container"><img src="charts/residential_province.png" alt="Residential Province"></div>
      <table>
        <tr><th>Province</th><th>Count</th><th>%</th></tr>
        {% set t = stats.residential_province.values()|sum %}
        {% for p,c in stats.residential_province.items() %}<tr><td>{{ p }}</td><td>{{ c }}</td><td>{{ "%.1f"|format((c/t)*100) }}%</td></tr>{% endfor %}
      </table>
    </div>{% endif %}
    {% if stats.site_province %}<div class="section">
      <h2>Learning Site Province</h2>
      <table>
        <tr><th>Province</th><th>Count</th><th>%</th></tr>
        {% set t = stats.site_province.values()|sum %}
        {% for p,c in stats.site_province.items() %}<tr><td>{{ p }}</td><td>{{ c }}</td><td>{{ "%.1f"|format((c/t)*100) }}%</td></tr>{% endfor %}
      </table>
    </div>{% endif %}
  </div>

  <!-- Metro and Rural -->
  {% if stats.metro_rural_crosstab %}<div class="section">
    <h2>Metro and Rural Classification</h2>
    <div class="chart-container"><img src="charts/metro_rural.png" alt="Metro Rural"></div>
    <table>
      <tr><th>Metro / Rural</th><th>Not Rural</th><th>Is Rural</th></tr>
      {% for key,vals in stats.metro_rural_crosstab.items() %}{% if key != 'All' %}
      <tr><td>{% if key == True %}In Metro{% elif key == False %}Not in Metro{% else %}{{ key }}{% endif %}</td>
          <td>{{ vals.get(False,0) }}</td><td>{{ vals.get(True,0) }}</td></tr>
      {% endif %}{% endfor %}
    </table>
  </div>{% endif %}

  <!-- Age Statistics -->
  {% if stats.age_stats %}<div class="section">
    <h2>Age Statistics (at Programme Start)</h2>
    <div class="chart-container"><img src="charts/age_distribution.png" alt="Age Distribution"></div>
    <div class="stats-grid">
      <div class="stat-card"><h3>Count</h3><div class="value">{{ stats.age_stats.count }}</div></div>
      <div class="stat-card"><h3>Mean</h3><div class="value">{{ stats.age_stats.mean }}</div></div>
      <div class="stat-card"><h3>Std Dev</h3><div class="value">{{ stats.age_stats.std }}</div></div>
      <div class="stat-card"><h3>Median</h3><div class="value">{{ stats.age_stats['50%'] }}</div></div>
    </div>
  </div>{% endif %}

  <!-- Duration Statistics -->
  {% if stats.duration_stats %}<div class="section">
    <h2>Programme Duration Statistics</h2>
    <div class="chart-container"><img src="charts/duration_distribution.png" alt="Duration Distribution"></div>
    <div class="stats-grid">
      <div class="stat-card"><h3>Count</h3><div class="value">{{ stats.duration_stats.count }}</div></div>
      <div class="stat-card"><h3>Mean</h3><div class="value">{{ stats.duration_stats.mean }}</div></div>
      <div class="stat-card"><h3>Min</h3><div class="value">{{ stats.duration_stats.min }}</div></div>
      <div class="stat-card"><h3>Max</h3><div class="value">{{ stats.duration_stats.max }}</div></div>
    </div>
  </div>{% endif %}

  <!-- Reasons for Leaving -->
  {% if stats.reasons %}<div class="section">
    <h2>Reasons for Leaving the Programme</h2>
    <div class="chart-container"><img src="charts/reasons.png" alt="Reasons"></div>
    <table>
      <tr><th>Rank</th><th>Reason</th><th>Count</th><th>Percentage</th></tr>
      {% set t = stats.reasons.values()|sum %}
      {% for r,c in stats.reasons.items() %}<tr><td>{{ loop.index }}</td><td>{{ r }}</td><td>{{ c }}</td><td>{{ "%.1f"|format((c/t)*100) }}%</td></tr>{% endfor %}
    </table>
  </div>{% endif %}

  <!-- Dropouts by Year -->
  {% if stats.dropouts_by_year %}<div class="section">
    <h2>Dropouts by Year</h2>
    <div class="chart-container"><img src="charts/dropouts_by_year.png" alt="Dropouts by Year"></div>
    <table>
      <tr><th>Year</th><th>Count</th></tr>
      {% for y,c in stats.dropouts_by_year.items() %}<tr><td>{{ y }}</td><td>{{ c }}</td></tr>{% endfor %}
    </table>
  </div>{% endif %}

  <!-- Dropouts by Month -->
  {% if stats.dropouts_by_month %}<div class="section">
    <h2>Dropouts by Month</h2>
    <div class="chart-container"><img src="charts/dropouts_by_month.png" alt="Dropouts by Month"></div>
    <table>
      <tr><th>Month</th><th>Count</th></tr>
      {% for m,c in stats.dropouts_by_month.items() %}{% if c > 0 %}<tr><td>{{ m }}</td><td>{{ c }}</td></tr>{% endif %}{% endfor %}
    </table>
  </div>{% endif %}

  <!-- Dropouts Timeline -->
  {% if stats.dropouts_by_year_month %}<div class="section">
    <h2>Dropouts Over Time</h2>
    <div class="chart-container"><img src="charts/dropouts_timeline.png" alt="Dropouts Timeline"></div>
  </div>{% endif %}

  <!-- Departure Info Coverage -->
  {% if stats.with_departure_info is defined %}<div class="section">
    <h2>Departure Information Coverage</h2>
    <div class="chart-container"><img src="charts/departure_info_coverage.png" alt="Coverage"></div>
    <div class="stats-grid">
      <div class="stat-card"><h3>Has Departure Info</h3><div class="value">{{ stats.with_departure_info }}</div></div>
      <div class="stat-card"><h3>No Departure Info</h3><div class="value">{{ stats.without_departure_info }}</div></div>
    </div>
  </div>{% endif %}

  <footer>Lead HR Consulting | NSF Analytics | Generated on {{ stats.report_generated }}</footer>
</div>
</body>
</html>
"""
        template = Template(html_template)
        html_content = template.render(stats=self.combined_stats)

        # Write the normal HTML (relative paths, for browser viewing)
        report_path = self.output_dir / "combined_report.html"
        report_path.write_text(html_content, encoding='utf-8')
        print(f"  Combined HTML saved: {report_path}")

        # Write a fully-embedded version for headless PDF generation
        embedded_path = self.output_dir / "combined_report_print.html"
        embedded_html = self._embed_images_as_base64(html_content)
        embedded_path.write_text(embedded_html, encoding='utf-8')

        return report_path, embedded_path


# GUI

import threading
import subprocess
import os

# Colours / fonts
BG        = "#1e1e2e"
PANEL     = "#2a2a3e"
ACCENT    = "#67c7f1"
ACCENT2   = "#5bbcec"
SUCCESS   = "#2ecc71"
ERROR     = "#e74c3c"
WARNING   = "#f39c12"
TEXT      = "#e0e0f0"
SUBTEXT   = "#9090b0"
DROP_IDLE = "#252538"
DROP_HOV  = "#2e2e50"
FONT      = ("Segoe UI", 10)
FONT_SM   = ("Segoe UI", 9)
FONT_LG   = ("Segoe UI", 13, "bold")
FONT_H    = ("Segoe UI", 11, "bold")


def _clean_drop_path(raw: str) -> str:
    """tkinterdnd2 wraps paths in braces on Windows; strip them."""
    raw = raw.strip()
    if raw.startswith("{") and raw.endswith("}"):
        raw = raw[1:-1]
    return raw


class DropZone:
    def __init__(self, parent, label: str, allowed_ext: tuple = (".xlsx",)):
        self.path: Path | None = None
        self.allowed_ext       = allowed_ext
        self._on_change        = []

        outer = parent

        self.frame = tk.Frame(outer, bg=PANEL, bd=0, highlightthickness=2, # type: ignore
                              highlightbackground=ACCENT, highlightcolor=ACCENT)

        tk.Label(self.frame, text=label, font=FONT_H, bg=PANEL, fg=ACCENT # type: ignore
                 ).pack(pady=(14, 0))

        self.drop_area = tk.Frame(self.frame, bg=DROP_IDLE, cursor="hand2") # type: ignore
        self.drop_area.pack(fill="both", expand=True, padx=16, pady=10)

        self.icon_lbl = tk.Label(self.drop_area, text="📂", font=("Segoe UI", 28), # type: ignore
                                 bg=DROP_IDLE, fg=SUBTEXT)
        self.icon_lbl.pack(pady=(18, 4))

        self.status_lbl = tk.Label(self.drop_area, # type: ignore
                                   text="Drag & drop your .xlsx file here\nor click to browse",
                                   font=FONT_SM, bg=DROP_IDLE, fg=SUBTEXT,
                                   justify="center", wraplength=230)
        self.status_lbl.pack(pady=(0, 6))

        self.file_lbl = tk.Label(self.drop_area, text="", font=FONT_SM, # type: ignore
                                 bg=DROP_IDLE, fg=SUCCESS, wraplength=230,
                                 justify="center")
        self.file_lbl.pack(pady=(0, 14))

        self.clear_btn = tk.Button(self.frame, text="✕  Clear", font=FONT_SM, # type: ignore
                                   bg=PANEL, fg=ERROR, bd=0, activebackground=PANEL,
                                   activeforeground=ERROR, cursor="hand2",
                                   command=self.clear)

        for w in (self.drop_area, self.icon_lbl, self.status_lbl, self.file_lbl):
            w.bind("<Button-1>", self._browse)
            w.bind("<Enter>",    self._hover_in)
            w.bind("<Leave>",    self._hover_out)

        self.drop_area.drop_target_register(DND_FILES) # type: ignore
        self.drop_area.dnd_bind("<<Drop>>",       self._on_drop)
        self.drop_area.dnd_bind("<<DragEnter>>",  self._drag_enter)
        self.drop_area.dnd_bind("<<DragLeave>>",  self._drag_leave)

    def on_change(self, fn):
        self._on_change.append(fn)

    def _notify(self):
        for fn in self._on_change:
            fn()

    def _set_path(self, p: Path):
        if p.suffix.lower() not in self.allowed_ext:
            self._set_error(f"Wrong file type: {p.suffix}\nExpected: {', '.join(self.allowed_ext)}")
            return
        self.path = p
        self.icon_lbl.config(text="✅", fg=SUCCESS)
        self.status_lbl.config(text="File loaded:", fg=SUBTEXT)
        self.file_lbl.config(text=p.name, fg=SUCCESS)
        self.drop_area.config(bg=DROP_IDLE)
        self.frame.config(highlightbackground=SUCCESS)
        self.clear_btn.pack(pady=(0, 10))
        self._notify()

    def _set_error(self, msg: str):
        self.icon_lbl.config(text="⚠️", fg=WARNING)
        self.status_lbl.config(text=msg, fg=WARNING)
        self.file_lbl.config(text="")
        self.drop_area.config(bg=DROP_IDLE)
        self.frame.config(highlightbackground=WARNING)

    def clear(self):
        self.path = None
        self.icon_lbl.config(text="📂", fg=SUBTEXT)
        self.status_lbl.config(text="Drag & drop your .xlsx file here\nor click to browse",
                               fg=SUBTEXT)
        self.file_lbl.config(text="")
        self.drop_area.config(bg=DROP_IDLE)
        self.frame.config(highlightbackground=ACCENT)
        self.clear_btn.pack_forget()
        self._notify()

    def _browse(self, _event=None):
        from tkinter import filedialog
        path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if path:
            self._set_path(Path(path))

    def _on_drop(self, event):
        self.drop_area.config(bg=DROP_IDLE)
        self._set_path(Path(_clean_drop_path(event.data)))

    def _drag_enter(self, _event):
        self.drop_area.config(bg=DROP_HOV)

    def _drag_leave(self, _event):
        self.drop_area.config(bg=DROP_IDLE)

    def _hover_in(self, _event):
        if not self.path:
            self.drop_area.config(bg=DROP_HOV)

    def _hover_out(self, _event):
        self.drop_area.config(bg=DROP_IDLE)


class ReportApp:
    """Main application window."""

    def __init__(self, root):
        self.root = root
        root.title("Lead HR NSF Analytics Reporting System")
        root.configure(bg=BG)
        root.resizable(False, False)

        self._build_ui()
        self._center_window()

    def _build_ui(self):
        root = self.root

        hdr = tk.Frame(root, bg=ACCENT2, pady=16) # type: ignore
        hdr.pack(fill="x")
        tk.Label(hdr, text="Lead HR NSF Analytics Reporting System", # type: ignore
                 font=("Segoe UI", 16, "bold"), bg=ACCENT2, fg="white").pack()
        tk.Label(hdr, text="Drop both Excel files below, choose an output folder, then Generate", # type: ignore
                 font=FONT_SM, bg=ACCENT2, fg="#cce0f5").pack(pady=(2, 0))

        zones_frame = tk.Frame(root, bg=BG) # type: ignore
        zones_frame.pack(fill="both", expand=True, padx=24, pady=20)

        self.learner_zone  = DropZone(zones_frame, "Learner Data (.xlsx)")
        self.dropouts_zone = DropZone(zones_frame, "Dropouts Master List (.xlsx)")

        self.learner_zone.frame.grid( row=0, column=0, sticky="nsew", padx=(0, 10), ipady=4)
        self.dropouts_zone.frame.grid(row=0, column=1, sticky="nsew", padx=(10, 0), ipady=4)

        zones_frame.columnconfigure(0, weight=1, minsize=270)
        zones_frame.columnconfigure(1, weight=1, minsize=270)
        zones_frame.rowconfigure(0, weight=1)

        for zone in (self.learner_zone, self.dropouts_zone):
            zone.on_change(self._refresh_button)

        out_frame = tk.Frame(root, bg=BG) # type: ignore
        out_frame.pack(fill="x", padx=24, pady=(0, 16))

        tk.Label(out_frame, text="Output folder:", font=FONT_H, # type: ignore
                 bg=BG, fg=TEXT).pack(side="left")

        self.out_var = tk.StringVar(value=str(Path.home() / "reports")) # type: ignore
        out_entry = tk.Entry(out_frame, textvariable=self.out_var, font=FONT, # type: ignore
                             bg=PANEL, fg=TEXT, insertbackground=TEXT,
                             relief="flat", bd=6, width=38)
        out_entry.pack(side="left", padx=10)

        tk.Button(out_frame, text="Browse…", font=FONT_SM, bg=ACCENT2, fg="white", # type: ignore
                  activebackground=ACCENT, activeforeground="white",
                  relief="flat", padx=10, cursor="hand2",
                  command=self._browse_output).pack(side="left")

        self.gen_btn = tk.Button(root, text="⚙  Generate Reports", # type: ignore
                                 font=("Segoe UI", 12, "bold"),
                                 bg=ACCENT2, fg="white",
                                 activebackground=ACCENT, activeforeground="white",
                                 relief="flat", padx=24, pady=12,
                                 cursor="hand2", state="disabled",
                                 command=self._start_generation)
        self.gen_btn.pack(pady=(0, 16))

        log_frame = tk.Frame(root, bg=PANEL, bd=0) # type: ignore
        log_frame.pack(fill="both", expand=True, padx=24, pady=(0, 20))

        tk.Label(log_frame, text="Log", font=FONT_H, bg=PANEL, fg=SUBTEXT, # type: ignore
                 anchor="w").pack(fill="x", padx=12, pady=(8, 0))

        self.log = tk.Text(log_frame, height=10, font=("Courier New", 9), # type: ignore
                           bg="#14141f", fg=TEXT, insertbackground=TEXT,
                           relief="flat", bd=8, state="disabled", wrap="word")
        self.log.pack(fill="both", expand=True, padx=8, pady=(4, 8))

        self.log.tag_config("ok",   foreground=SUCCESS)
        self.log.tag_config("err",  foreground=ERROR)
        self.log.tag_config("warn", foreground=WARNING)
        self.log.tag_config("head", foreground=ACCENT, font=("Courier New", 9, "bold"))

        bar = tk.Frame(root, bg=PANEL, pady=8) # type: ignore
        bar.pack(fill="x", side="bottom")

        self.status_lbl = tk.Label(bar, text="Ready — drop your files above to begin", # type: ignore
                                   font=FONT_SM, bg=PANEL, fg=SUBTEXT)
        self.status_lbl.pack()

        self._last_output_dir = None
        self.open_folder_btn = tk.Button( # type: ignore
            bar, text="Open Output Folder",
            font=FONT_SM, bg=PANEL, fg=ACCENT,
            activebackground=BG, activeforeground=ACCENT,
            relief="flat", padx=12, pady=4, cursor="hand2",
            command=self._open_output_folder)

    def _center_window(self):
        self.root.update_idletasks()
        w, h = 680, 700
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        self.root.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    def _refresh_button(self):
        both_ready = self.learner_zone.path and self.dropouts_zone.path
        self.gen_btn.config(state="normal" if both_ready else "disabled",
                            bg=SUCCESS if both_ready else ACCENT2)

    def _browse_output(self):
        from tkinter import filedialog
        d = filedialog.askdirectory(title="Select output folder")
        if d:
            self.out_var.set(d)

    def _log(self, msg: str, tag: str = ""):
        self.log.config(state="normal")
        if tag:
            self.log.insert("end", msg + "\n", tag)
        else:
            self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.config(state="disabled")

    def _set_status(self, msg: str, color: str = SUBTEXT):
        self.status_lbl.config(text=msg, fg=color)

    def _start_generation(self):
        """Run combined report in a background thread so the GUI stays responsive."""
        learner_path  = self.learner_zone.path
        dropouts_path = self.dropouts_zone.path
        output_dir    = Path(self.out_var.get().strip() or "./reports")

        self.gen_btn.config(state="disabled", text="Generating…", bg=ACCENT2)
        self._set_status("Generating reports, please wait…", ACCENT)
        self.log.config(state="normal"); self.log.delete("1.0", "end"); self.log.config(state="disabled")

        def run():
            result = None
            error  = None

            self._log("━━━  COMBINED NSF ANALYTICS REPORT  ━━━", "head")
            try:
                import builtins
                orig_print = builtins.print

                def gui_print(*args, **kwargs):
                    msg = " ".join(str(a) for a in args)
                    self.root.after(0, self._log, msg)

                builtins.print = gui_print
                try:
                    sys_obj = CombinedReportSystem(learner_path, dropouts_path, output_dir)
                    result = sys_obj.run()
                finally:
                    builtins.print = orig_print

                self.root.after(0, self._log, "✔  Combined report generated successfully", "ok")
            except Exception as e:
                error = str(e)
                self.root.after(0, self._log, f"✗  Report generation failed: {e}", "err")
                import traceback
                self.root.after(0, self._log, traceback.format_exc(), "err")

            self.root.after(0, self._finish, result, error, output_dir)

        threading.Thread(target=run, daemon=True).start()

    def _finish(self, result, error: str, output_dir: Path):
        self._log("")
        self._log("━━━  SUMMARY  ━━━", "head")

        if result and result.get("html"):
            self._log(f"  HTML → {result['html']}", "ok")
        if result and result.get("pdf"):
            self._log(f"  PDF  → {result['pdf']}", "ok")

        if error:
            self._log(f"  ⚠  {error}", "warn")
            self._set_status("Completed with errors — see log above", WARNING)
        else:
            self._set_status("✔  Combined report generated successfully!", SUCCESS)

        self.gen_btn.config(state="normal", text="⚙  Generate Reports", bg=ACCENT2)
        self._refresh_button()

        self._last_output_dir = output_dir
        self.open_folder_btn.pack(pady=(4, 0))

    def _open_output_folder(self):
        path = self._last_output_dir
        if not path:
            return
        path.mkdir(parents=True, exist_ok=True)
        import platform
        system = platform.system()
        if system == "Windows":
            os.startfile(str(path))
        elif system == "Darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])


def main():
    try:
        from tkinterdnd2 import TkinterDnD, DND_FILES
    except ImportError:
        print("ERROR: tkinterdnd2 is not installed.")
        print("Run:  pip install tkinterdnd2")
        sys.exit(1)

    import builtins
    builtins.DND_FILES = DND_FILES

    import tkinter as tk
    builtins.tk = tk

    root = TkinterDnD.Tk()
    app  = ReportApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()