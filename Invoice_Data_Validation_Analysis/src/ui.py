"""
ui.py - Tailwind UI Components & Streamlit Layout Helper
Project: PJ-IVD-2026 | Surbhi Sharma (2500320140152) | ABES Engineering College
"""

import streamlit as st

def inject_custom_css():
    """
    Injects Tailwind-compatible styling and clean modern UI CSS for cards, badges, and headers.
    """
    custom_css = """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
        
        html, body, [class*="css"] {
            font-family: 'Inter', sans-serif;
        }
        
        .main-header {
            background: linear-gradient(135deg, #1E3A8A 0%, #1E293B 100%);
            color: #FFFFFF;
            padding: 24px;
            border-radius: 12px;
            margin-bottom: 24px;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        }
        .main-header h1 {
            font-size: 26px;
            font-weight: 700;
            margin: 0;
            color: #FFFFFF;
        }
        .main-header p {
            font-size: 14px;
            color: #94A3B8;
            margin: 6px 0 0 0;
        }
        
        .kpi-card {
            background-color: #FFFFFF;
            border: 1px solid #E2E8F0;
            border-radius: 10px;
            padding: 18px;
            box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05);
            transition: all 0.2s ease;
        }
        .kpi-card:hover {
            box-shadow: 0 6px 12px -2px rgba(0, 0, 0, 0.1);
            transform: translateY(-2px);
        }
        .kpi-label {
            font-size: 12px;
            font-weight: 600;
            color: #64748B;
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }
        .kpi-val {
            font-size: 24px;
            font-weight: 700;
            color: #0F172A;
            margin-top: 4px;
        }
        .kpi-sub {
            font-size: 11.5px;
            margin-top: 4px;
            font-weight: 500;
        }
        .val-danger { color: #DC2626; }
        .val-success { color: #16A34A; }
        .val-warning { color: #D97706; }
        .val-primary { color: #2563EB; }
        
        .badge {
            display: inline-block;
            padding: 3px 8px;
            border-radius: 9999px;
            font-size: 11px;
            font-weight: 600;
        }
        .badge-red { background-color: #FEE2E2; color: #991B1B; }
        .badge-green { background-color: #DCFCE7; color: #166534; }
        .badge-yellow { background-color: #FEF3C7; color: #92400E; }
    </style>
    """
    st.markdown(custom_css, unsafe_allow_html=True)

def render_header(title, subtitle):
    inject_custom_css()
    html = f"""
    <div class="main-header">
        <h1>{title}</h1>
        <p>{subtitle} | Project ID: PJ-IVD-2026 | ABES Engineering College (MCA DA)</p>
    </div>
    """
    st.markdown(html, unsafe_allow_html=True)

def render_kpi(label, value, subtext="", trend_class="val-primary"):
    html = f"""
    <div class="kpi-card">
        <div class="kpi-label">{label}</div>
        <div class="kpi-val">{value}</div>
        <div class="kpi-sub {trend_class}">{subtext}</div>
    </div>
    """
    st.markdown(html, unsafe_allow_html=True)
