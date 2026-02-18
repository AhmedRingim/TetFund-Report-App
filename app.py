import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
import base64
import json
import os
import traceback

# ===================== PDF GENERATION WITH WEASYPRINT =====================
try:
    from weasyprint import HTML
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

from PIL import Image

# ===================== CONFIG =====================
APP_NAME = "TETFUND Monitoring Report System"
DATA_DIR = "drafts"
IMAGES_DIR = "images"  # Directory for logo images
os.makedirs(DATA_DIR, exist_ok=True)

# Default bank charges amount (can be changed by user)
DEFAULT_BANK_CHARGES_AMOUNT = 215_013.00

PROJECT_COLUMNS = [
    "s_no", "project", "approved_cost", "contract_sum", "disbursed", "balance",
    "quality", "compliance", "other_obs", "completion", "docs", "recommendation"
]

DEFAULT_PROJECT = {
    "s_no": 1,
    "project": "New Project",
    "approved_cost": 0.0,
    "contract_sum": 0.0,
    "disbursed": 0.0,
    "balance": 100.0,
    "quality": "Good",
    "compliance": "Compliant",
    "other_obs": "",
    "completion": 0.0,
    "docs": "Pending",
    "recommendation": "Pending Review",
}

DEFAULT_MONITORING_TEAM = [
    {"name": "Arch. A.A.", "designation": "Team Lead"},
    {"name": "Engr. A.B.", "designation": "Monitor"},
    {"name": "Mr. A.C.", "designation": "Engineer"},
]

# ===================== LOAD LOGO =====================
def load_logo():
    """Load logo from images folder and convert to base64"""
    logo_path = os.path.join(IMAGES_DIR, "tetfund_logo.png")
    if os.path.exists(logo_path):
        try:
            img = Image.open(logo_path)
            # Resize image to reasonable size for logo
            max_size = (150, 150)
            img.thumbnail(max_size, Image.Resampling.LANCZOS)
            buf = BytesIO()
            img.save(buf, format="PNG")
            return base64.b64encode(buf.getvalue()).decode()
        except Exception as e:
            print(f"Error loading logo: {e}")
            return None
    return None

# Load logo at startup
LOGO_BASE64 = load_logo()

# ===================== PAGE =====================
st.set_page_config(
    page_title=APP_NAME,
    page_icon="images/tetfund_logo.png",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ===================== STYLING =====================
st.markdown("""
<style>
.main > div {
    padding: 0.5rem 1rem;
    max-width: 1750px;
    margin: auto;
}
[data-testid="stDataEditor"] { font-size: 12px; }
.main-header {
    font-size: 24px; font-weight: bold; color: #1E3A8A;
    text-align: center; padding: 12px;
    background-color: #EFF6FF; border-radius: 8px;
    margin-bottom: 10px; border-bottom: 3px solid #1E3A8A;
}
.section-header {
    font-size: 17px; font-weight: bold; color: white;
    background-color: #1E3A8A; padding: 8px 12px;
    border-radius: 5px; margin: 18px 0 10px 0;
    border-left: 5px solid #FCD34D;
}
.subsection-header {
    font-size: 15px; font-weight: bold; color: #2563EB;
    margin: 10px 0 6px 0; padding-bottom: 4px;
    border-bottom: 2px solid #93C5FD;
}
.metric-container {
    background-color: #F3F4F6; padding: 10px;
    border-radius: 6px; border-left: 4px solid #1E3A8A;
    margin-bottom: 8px;
}
.metric-value { font-size: 22px; font-weight: bold; color: #1E3A8A; }
.metric-label { font-size: 12px; color: #6B7280; text-transform: uppercase; }
.custom-divider {
    height: 3px;
    background: linear-gradient(90deg, #1E3A8A, #93C5FD, #EFF6FF);
    margin: 18px 0;
}
.info-box {
    background-color: #F0F9FF; padding: 10px;
    border-radius: 6px; border: 1px solid #BAE6FD;
    margin: 6px 0;
}
.footer {
    text-align: center; padding: 16px; color: #6B7280;
    font-size: 11px; margin-top: 30px; border-top: 1px solid #E5E7EB;
}
.stButton > button {
    width: 100%; background-color: #1E3A8A;
    color: white; font-weight: bold;
    border-radius: 6px; padding: 8px;
}
.stButton > button:hover { background-color: #2563EB; }
.stButton > button.danger {
    background-color: #DC2626;
}
.stButton > button.danger:hover { background-color: #B91C1C; }
.logo-container {
    text-align: center;
    margin-bottom: 10px;
    width: 100%;
}
.logo-img {
    max-height: 120px;
    max-width: 100%;
    object-fit: contain;
    display: block;
    margin: 0 auto;
}
.fund-title {
    font-size: 28px;
    font-weight: bold;
    color: #1E3A8A;
    text-align: center;
    margin: 10px 0 5px 0;
    padding: 0;
    width: 100%;
}

/* Print-friendly */
@media print {
    header, footer, .stButton, .stDownloadButton, .stFileUploader,
    .stCheckbox, .stTextInput, .stSelectbox, .stDateInput, .stTextArea {
        display: none !important;
    }
    .main > div { max-width: 100% !important; padding: 0 !important; }
}
</style>
""", unsafe_allow_html=True)

# ===================== STATE =====================
def init_state():
    defaults = {
        "projects": [DEFAULT_PROJECT.copy()],
        "institution_name": "",
        "location": "",
        "intervention_year": str(datetime.now().year),
        "inspection_date": datetime.now().date(),
        "bank_charges_added": False,
        "bank_charges_amount": DEFAULT_BANK_CHARGES_AMOUNT,
        "monitoring_team": [m.copy() for m in DEFAULT_MONITORING_TEAM],
        "approval_status": "Pending",
        "dme_officer": "",
        "approval_comments": "",
        "generated_reports": [],
        "draft_loaded": None,
        "pdf_orientation": "Landscape",
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()

# ===================== AUTOSAVE =====================
def autosave():
    draft = {
        "projects": st.session_state.projects,
        "institution_name": st.session_state.institution_name,
        "location": st.session_state.location,
        "intervention_year": st.session_state.intervention_year,
        "inspection_date": st.session_state.inspection_date.strftime("%Y-%m-%d"),
        "bank_charges_added": st.session_state.bank_charges_added,
        "bank_charges_amount": st.session_state.bank_charges_amount,
        "monitoring_team": st.session_state.monitoring_team,
        "approval_status": st.session_state.approval_status,
        "dme_officer": st.session_state.dme_officer,
        "approval_comments": st.session_state.approval_comments,
        "pdf_orientation": st.session_state.pdf_orientation,
        "saved_at": datetime.now().isoformat(),
    }
    path = os.path.join(DATA_DIR, "autosave.json")
    with open(path, "w") as f:
        json.dump(draft, f)

def load_autosave():
    path = os.path.join(DATA_DIR, "autosave.json")
    if os.path.exists(path):
        with open(path) as f:
            draft = json.load(f)
        st.session_state.projects = draft["projects"]
        st.session_state.institution_name = draft["institution_name"]
        st.session_state.location = draft["location"]
        st.session_state.intervention_year = draft["intervention_year"]
        st.session_state.inspection_date = datetime.fromisoformat(draft["inspection_date"]).date()
        st.session_state.bank_charges_added = draft["bank_charges_added"]
        st.session_state.bank_charges_amount = draft.get("bank_charges_amount", DEFAULT_BANK_CHARGES_AMOUNT)
        st.session_state.monitoring_team = draft["monitoring_team"]
        st.session_state.approval_status = draft["approval_status"]
        st.session_state.dme_officer = draft["dme_officer"]
        st.session_state.approval_comments = draft["approval_comments"]
        st.session_state.pdf_orientation = draft.get("pdf_orientation", "Landscape")
        st.session_state.draft_loaded = draft.get("saved_at")

if os.path.exists(os.path.join(DATA_DIR, "autosave.json")) and not st.session_state.get("draft_loaded"):
    load_autosave()

# ===================== HELPERS =====================
def recalc_projects(df):
    df["disbursed"] = pd.to_numeric(df["disbursed"], errors="coerce").fillna(0)
    df["completion"] = pd.to_numeric(df["completion"], errors="coerce").fillna(0)
    df["balance"] = 100 - df["disbursed"]
    df["s_no"] = range(1, len(df) + 1)
    return df

# ===================== EXCEL EXPORT =====================
def export_excel(projects, institution_info, summary):
    df = pd.DataFrame(projects)
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Projects")
        pd.DataFrame([institution_info]).to_excel(writer, index=False, sheet_name="Institution")
        pd.DataFrame([summary]).to_excel(writer, index=False, sheet_name="Summary")
    return buf.getvalue()

# ===================== PDF GENERATION WITH WEASYPRINT =====================
def generate_pdf(projects, institution_info, monitoring_team, approval_info, summary, orientation, logo_base64):
    try:
        # Create HTML content
        html_content = create_html_report(projects, institution_info, monitoring_team, 
                                         approval_info, summary, orientation, logo_base64)
        
        # Generate PDF using WeasyPrint
        pdf_file = BytesIO()
        HTML(string=html_content).write_pdf(pdf_file)
        pdf_file.seek(0)
        
        return pdf_file.getvalue()
        
    except Exception as e:
        st.error(f"PDF generation failed: {str(e)}")
        st.error(traceback.format_exc())
        return None

def create_html_report(projects, institution_info, monitoring_team, approval_info, summary, orientation, logo_base64):
    # Set orientation
    page_size = "A4 landscape" if orientation == "Landscape" else "A4"
    
    # Create project rows
    project_rows = ""
    for i, p in enumerate(projects, 1):
        approved_cost = float(p.get('approved_cost', 0))
        contract_sum = float(p.get('contract_sum', 0))
        disbursed = float(p.get('disbursed', 0))
        balance = float(p.get('balance', 0))
        completion = float(p.get('completion', 0))
        
        project_rows += f"""
        <tr>
            <td style="text-align: center;">{i}</td>
            <td>{p.get('project', '')}</td>
            <td style="text-align: right;">₦{approved_cost:,.0f}</td>
            <td style="text-align: right;">₦{contract_sum:,.0f}</td>
            <td style="text-align: center;">{disbursed:.1f}%</td>
            <td style="text-align: center;">{balance:.1f}%</td>
            <td style="text-align: center;">{p.get('quality', '')}</td>
            <td style="text-align: center;">{p.get('compliance', '')}</td>
            <td>{p.get('other_obs', '')}</td>
            <td style="text-align: center;">{completion:.1f}%</td>
            <td style="text-align: center;">{p.get('docs', '')}</td>
            <td>{p.get('recommendation', '')}</td>
        </tr>
        """
    
    # Create team rows
    team_rows = ""
    for i, m in enumerate(monitoring_team, 1):
        team_rows += f"""
        <tr>
            <td style="text-align: center;">{i}</td>
            <td>{m.get('name', '')}</td>
            <td>{m.get('designation', '')}</td>
            <td style="text-align: center;">________________</td>
        </tr>
        """
    
    # Logo HTML
    if logo_base64:
        logo_html = f'<div style="text-align: center;"><img src="data:image/png;base64,{logo_base64}" style="height: 80px; width: auto;"></div>'
    else:
        logo_html = '<h1 style="text-align: center; color: #1E3A8A;">TERTIARY EDUCATION TRUST FUND</h1>'
    
    # Complete HTML template
    return f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>TETFUND Monitoring Report</title>
    <style>
        @page {{
            size: {page_size};
            margin: 2cm;
        }}
        body {{
            font-family: Arial, Helvetica, sans-serif;
            font-size: 10pt;
            line-height: 1.4;
            color: #333;
        }}
        h1 {{
            color: #1E3A8A;
            font-size: 24pt;
            text-align: center;
            margin: 10px 0;
        }}
        h2 {{
            color: #1E3A8A;
            font-size: 16pt;
            text-align: center;
            margin: 5px 0;
        }}
        h3 {{
            color: #1E3A8A;
            font-size: 14pt;
            text-align: center;
            margin: 15px 0 10px 0;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin: 15px 0;
            font-size: 9pt;
        }}
        th {{
            background-color: #1E3A8A;
            color: white;
            padding: 8px 4px;
            font-weight: bold;
            text-align: center;
            border: 1px solid #1E3A8A;
        }}
        td {{
            border: 1px solid #999;
            padding: 6px 4px;
            vertical-align: top;
        }}
        .info-box {{
            background-color: #f5f5f5;
            padding: 10px;
            margin: 15px 0;
            border: 1px solid #ddd;
            border-radius: 3px;
        }}
        .signature {{
            margin-top: 30px;
            padding-top: 10px;
        }}
        .footer {{
            text-align: center;
            font-size: 8pt;
            color: #666;
            margin-top: 30px;
            padding-top: 10px;
            border-top: 1px solid #ccc;
        }}
        .text-center {{ text-align: center; }}
        .text-right {{ text-align: right; }}
    </style>
</head>
<body>
    {logo_html}
    
    <h1>TERTIARY EDUCATION TRUST FUND</h1>
    <h2>Monitoring & Evaluation Department</h2>
    <h3>Second / Final Tranche Monitoring Report</h3>

    <div class="info-box">
        <strong>Institution:</strong> {institution_info.get('name', '')} | 
        <strong>Location:</strong> {institution_info.get('location', '')} | 
        <strong>Inspection Date:</strong> {institution_info.get('date', '')} | 
        <strong>Intervention Year:</strong> {institution_info.get('year', '')}
    </div>

    <h3>Projects Monitoring Details</h3>
    <table>
        <thead>
            <tr>
                <th style="width: 3%;">S/N</th>
                <th style="width: 12%;">Project</th>
                <th style="width: 8%;">Approved (₦)</th>
                <th style="width: 8%;">Contract (₦)</th>
                <th style="width: 4%;">%Disb</th>
                <th style="width: 4%;">%Bal</th>
                <th style="width: 6%;">Quality</th>
                <th style="width: 6%;">Compl</th>
                <th style="width: 15%;">Observations</th>
                <th style="width: 4%;">%Comp</th>
                <th style="width: 6%;">Docs</th>
                <th style="width: 14%;">Recommendation</th>
            </tr>
        </thead>
        <tbody>
            {project_rows}
        </tbody>
    </table>

    <div class="info-box">
        <strong>Total Projects:</strong> {summary.get('total_projects', 0)} | 
        <strong>Completed:</strong> {summary.get('completed', 0)} | 
        <strong>In Progress:</strong> {summary.get('in_progress', 0)} | 
        <strong>Completion Rate:</strong> {summary.get('completion_rate', 0):.1f}%<br>
        <strong>Total Approved:</strong> ₦{summary.get('total_approved', 0):,.2f} | 
        <strong>Total Contract:</strong> ₦{summary.get('total_contract', 0):,.2f} | 
        <strong>Total Disbursed:</strong> ₦{summary.get('total_disbursed', 0):,.2f} | 
        <strong>Balance:</strong> ₦{summary.get('balance', 0):,.2f}
    </div>

    <h3>Monitoring Team</h3>
    <table>
        <thead>
            <tr>
                <th style="width: 10%;">S/N</th>
                <th style="width: 40%;">Name</th>
                <th style="width: 30%;">Designation</th>
                <th style="width: 20%;">Signature</th>
            </tr>
        </thead>
        <tbody>
            {team_rows}
        </tbody>
    </table>

    <h3>DM&E Approval</h3>
    <div class="info-box">
        <strong>Status:</strong> {approval_info.get('status', '')} | 
        <strong>Officer:</strong> {approval_info.get('officer', '')} | 
        <strong>Date:</strong> {approval_info.get('date', '')}<br>
        <strong>Comments:</strong> {approval_info.get('comments', '')}
    </div>

    <div class="signature">
        <strong>DM&E Officer Signature:</strong><br>
        _________________________________________
    </div>

    <div class="footer">
        Generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
    </div>
</body>
</html>"""

# ===================== RESET =====================
def reset_all():
    for k in list(st.session_state.keys()):
        del st.session_state[k]
    init_state()
    st.rerun()

# ===================== MAIN =====================
def main():
    
    # ---------------- LOGO/TITLE DISPLAY ----------------
    if LOGO_BASE64:
        # Display centered logo
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.markdown(f'<div style="text-align: center;"><img src="data:image/png;base64,{LOGO_BASE64}" style="max-height: 120px; max-width: 100%; object-fit: contain;"></div>', unsafe_allow_html=True)
            
            # "TERTIARY EDUCATION TRUST FUND" text under logo
            st.markdown('<p style="text-align:center;font-size:22px;font-weight:bold;color:#1E3A8A; margin: 5px 0 0 0;">TERTIARY EDUCATION TRUST FUND</p>', unsafe_allow_html=True)
    else:
        # Display centered title
        st.markdown('<h1 style="text-align: center; color: #1E3A8A; font-size: 28px; margin: 10px 0;">TERTIARY EDUCATION TRUST FUND</h1>', unsafe_allow_html=True)       
        
        
    
    st.markdown('<p style="text-align:center;font-size:18px;color:#2563EB; margin: 5px 0;">MONITORING AND EVALUATION DEPARTMENT</p>', unsafe_allow_html=True)
    st.markdown('<p style="text-align:center;font-size:16px;font-weight:bold; margin: 5px 0 20px 0;">SECOND / FINAL TRANCHE MONITORING REPORT</p>', unsafe_allow_html=True)

    # ---------------- INSTITUTION INFO ----------------
    st.markdown('<div class="section-header">🏢 INSTITUTION INFORMATION</div>', unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        st.session_state.institution_name = st.text_input("Institution Name *", st.session_state.institution_name)
        st.session_state.location = st.text_input("Location *", st.session_state.location)
    with c2:
        st.session_state.intervention_year = st.text_input(
            "Intervention Year *",
            st.session_state.intervention_year,
            help="Enter any year, e.g., 2019, 2025, etc."
        )
        st.session_state.inspection_date = st.date_input("Inspection Date *", st.session_state.inspection_date)
    with c3:
        st.markdown('<div class="metric-container">', unsafe_allow_html=True)
        st.markdown('<div class="metric-label">Total Projects</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-value">{len(st.session_state.projects)}</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Bank charges amount input
        bank_charges_amount = st.number_input(
            "Bank & Admin Charges (₦)",
            min_value=0.0,
            value=float(st.session_state.bank_charges_amount),
            step=1000.0,
            format="%.2f",
            key="bank_charges_input"
        )
        st.session_state.bank_charges_amount = bank_charges_amount
        st.session_state.bank_charges_added = st.checkbox(
            "Include Bank & Admin Charges",
            value=st.session_state.bank_charges_added,
        )

    st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)

    # ---------------- AUTOSAVE ----------------
    autosave()

    # ---------------- PROJECTS ----------------
    st.markdown('<div class="section-header">🏗️ PROJECTS MONITORING DETAILS</div>', unsafe_allow_html=True)

    df = pd.DataFrame(st.session_state.projects)
    edited_df = st.data_editor(
        df,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        column_config={
            "s_no": st.column_config.NumberColumn("S/N", disabled=True, width=45),
            "project": st.column_config.TextColumn("PROJECT", required=True, width=170),
            "approved_cost": st.column_config.NumberColumn("APPROVED (₦)", format="₦%.0f", min_value=0, width=95),
            "contract_sum": st.column_config.NumberColumn("CONTRACT (₦)", format="₦%.0f", min_value=0, width=95),
            "disbursed": st.column_config.NumberColumn("%DISB", min_value=0, max_value=100, width=60),
            "balance": st.column_config.NumberColumn("%BAL", disabled=True, width=60),
            "quality": st.column_config.SelectboxColumn("QUALITY", options=["Excellent", "Good", "Average", "Poor"], width=80),
            "compliance": st.column_config.SelectboxColumn("COMPL", options=["Compliant", "Partial", "Non-compliant"], width=90),
            "other_obs": st.column_config.TextColumn("OBSERVATIONS", width=180),
            "completion": st.column_config.NumberColumn("%COMP", min_value=0, max_value=100, width=60),
            "docs": st.column_config.SelectboxColumn("DOCS", options=["Submitted", "Pending", "Incomplete"], width=80),
            "recommendation": st.column_config.TextColumn("RECOMMENDATION", width=180),
        },
        key="projects_editor",
    )

    if edited_df is not None and not edited_df.empty:
        edited_df = recalc_projects(edited_df)
        st.session_state.projects = edited_df.to_dict("records")

    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("➕ Add Project"):
            new = DEFAULT_PROJECT.copy()
            new["s_no"] = len(st.session_state.projects) + 1
            st.session_state.projects.append(new)
            st.rerun()
    with col2:
        if st.button("🗑️ Remove Last Project"):
            if st.session_state.projects:
                st.session_state.projects.pop()
                st.rerun()
    with col3:
        # Reset All Data button with danger styling
        st.markdown('<style>.stButton > button[key="reset_button"] {background-color: #DC2626;}</style>', unsafe_allow_html=True)
        if st.button("🔄 RESET ALL DATA", key="reset_button"):
            reset_all()

    st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)

    # ---------------- MONITORING TEAM ----------------
    st.markdown('<div class="section-header">👥 MONITORING TEAM DETAILS</div>', unsafe_allow_html=True)

    for i in range(len(st.session_state.monitoring_team)):
        with st.expander(f"Team Member {i+1}", expanded=i == 0):
            st.session_state.monitoring_team[i]["name"] = st.text_input(
                "Full Name",
                st.session_state.monitoring_team[i]["name"],
                key=f"team_name_{i}",
            )
            st.session_state.monitoring_team[i]["designation"] = st.text_input(
                "Designation",
                st.session_state.monitoring_team[i]["designation"],
                key=f"team_desig_{i}",
            )

    tcol1, tcol2 = st.columns(2)
    with tcol1:
        if st.button("➕ Add Team Member"):
            st.session_state.monitoring_team.append({"name": "", "designation": ""})
            st.rerun()
    with tcol2:
        if st.button("🗑️ Remove Last Team Member"):
            if st.session_state.monitoring_team:
                st.session_state.monitoring_team.pop()
                st.rerun()

    st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)

    # ---------------- DM&E APPROVAL ----------------
    st.markdown('<div class="section-header">✅ DM&E APPROVAL</div>', unsafe_allow_html=True)

    left, right = st.columns(2)
    with left:
        st.session_state.approval_status = st.selectbox(
            "Approval Status",
            ["Pending", "Approved", "Returned for Correction", "Rejected"],
            index=["Pending", "Approved", "Returned for Correction", "Rejected"].index(st.session_state.approval_status),
        )
        st.session_state.dme_officer = st.text_input("DM&E Officer Name", st.session_state.dme_officer)
    with right:
        approval_date = st.date_input("Approval Date", datetime.now().date())
        st.session_state.approval_comments = st.text_area("Approval Comments", st.session_state.approval_comments, height=90)

    st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)

    # ---------------- SUMMARY STATISTICS (Hidden but calculated for exports) ----------------
    projects = st.session_state.projects
    total_projects = len(projects)
    completed = sum(1 for p in projects if p["completion"] >= 100)
    in_progress = total_projects - completed
    completion_rate = (completed / total_projects * 100) if total_projects else 0

    total_approved = sum(p["approved_cost"] for p in projects)
    total_contract = sum(p["contract_sum"] for p in projects)
    total_disbursed = sum((p["disbursed"] / 100) * p["contract_sum"] for p in projects)

    # Add bank charges if checked
    if st.session_state.bank_charges_added:
        bank_charges = st.session_state.bank_charges_amount
        total_approved += bank_charges
        total_contract += bank_charges
        total_disbursed += bank_charges

    # ---------------- EXPORT TO EXCEL ----------------
    st.markdown('<div class="section-header">📥 EXPORT TO EXCEL</div>', unsafe_allow_html=True)

    institution_info = {
        "name": st.session_state.institution_name,
        "location": st.session_state.location,
        "year": st.session_state.intervention_year,
        "date": st.session_state.inspection_date.strftime("%d-%b-%Y"),
    }

    summary = {
        "total_projects": total_projects,
        "completed": completed,
        "in_progress": in_progress,
        "completion_rate": completion_rate,
        "total_approved": total_approved,
        "total_contract": total_contract,
        "total_disbursed": total_disbursed,
        "balance": total_contract - total_disbursed,
    }

    excel_bytes = export_excel(st.session_state.projects, institution_info, summary)

    st.download_button(
        "📊 Download Excel Report",
        excel_bytes,
        file_name=f"TETFUND_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)

    # ---------------- PDF GENERATION ----------------
    st.markdown('<div class="section-header">📄 GENERATE OFFICIAL PDF</div>', unsafe_allow_html=True)

    if not PDF_AVAILABLE:
        st.error("PDF generation not installed. Run:")
        st.code("pip install weasyprint")
    else:
        # PDF Orientation selector
        st.session_state.pdf_orientation = st.radio(
            "PDF Orientation",
            ["Landscape", "Portrait"],
            horizontal=True,
            index=0 if st.session_state.pdf_orientation == "Landscape" else 1
        )
        
        valid = True
        if not st.session_state.institution_name:
            st.warning("❌ Institution name is required")
            valid = False
        if not st.session_state.location:
            st.warning("❌ Location is required")
            valid = False
        if not st.session_state.projects:
            st.warning("❌ At least one project is required")
            valid = False

        if st.button("📄 GENERATE OFFICIAL PDF REPORT", type="primary"):
            if not valid:
                st.error("Fix validation errors before generating PDF.")
            else:
                with st.spinner("Generating official PDF..."):
                    approval_info = {
                        "status": st.session_state.approval_status,
                        "officer": st.session_state.dme_officer,
                        "date": approval_date.strftime("%d-%b-%Y"),
                        "comments": st.session_state.approval_comments or "No comments provided",
                    }

                    pdf_projects = st.session_state.projects.copy()
                    
                    # Add bank charges as a separate project if checked
                    if st.session_state.bank_charges_added:
                        bank_charges = st.session_state.bank_charges_amount
                        pdf_projects.append({
                            "project": "Bank and Administrative Charges",
                            "approved_cost": bank_charges,
                            "contract_sum": bank_charges,
                            "disbursed": 100.0,
                            "balance": 0.0,
                            "quality": "N/A",
                            "compliance": "N/A",
                            "other_obs": "Administrative charges",
                            "completion": 100.0,
                            "docs": "Submitted",
                            "recommendation": "Processed",
                        })

                    pdf_bytes = generate_pdf(
                        pdf_projects,
                        institution_info,
                        st.session_state.monitoring_team,
                        approval_info,
                        summary,
                        st.session_state.pdf_orientation,
                        LOGO_BASE64,
                    )

                    if pdf_bytes:
                        st.success("✅ Official PDF generated successfully!")

                        filename = f"TETFUND_Report_{st.session_state.institution_name.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"

                        st.download_button(
                            "📥 Download Official PDF",
                            pdf_bytes,
                            file_name=filename,
                            mime="application/pdf",
                            use_container_width=True,
                        )

                        with st.expander("🖨️ Print Preview", expanded=True):
                            st.markdown(
                                f'<iframe src="data:application/pdf;base64,{base64.b64encode(pdf_bytes).decode()}" '
                                f'width="100%" height="650"></iframe>',
                                unsafe_allow_html=True,
                            )
                    else:
                        st.error("❌ PDF generation failed. Check the error messages above.")

    # ---------------- FOOTER ----------------
    st.markdown('<div class="footer">', unsafe_allow_html=True)
    st.markdown("© 2026 Tertiary Education Trust Fund — Monitoring & Evaluation Department")
    st.markdown(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    st.markdown('</div>', unsafe_allow_html=True)


if __name__ == "__main__":
    main()