"""Sistem Pendataan & Interview Peserta Magang — Streamlit.

Run: streamlit run app.py
Storage: Supabase PostgreSQL (persistent) + optional Excel export/import.
"""
from __future__ import annotations

import csv
import html
import hashlib
from datetime import date, datetime
from io import StringIO
from pathlib import Path
import hmac

import altair as alt

import streamlit as st

from db_store import (
    DatabaseOperationError, configured, load_records, save_candidate, save_interview, import_backup,
)

from data_store import (
    ASPEK,
    DECISIONS,
    MAX_KANDIDAT,
    calculate,
    candidate_id_next,
    export_xlsx,
    get_rubric_and_questions,
    ranking_rows,
    read_xlsx,
)

st.set_page_config(
    page_title="Interview Peserta Magang",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
:root {--navy:#102743;--muted:#62748a;--line:#e5ebf3;--blue:#3567d8;--pale:#f5f8fd;}
html,body,[class*="css"],[data-testid="stApp"]{font-family:'DM Sans',sans-serif;color:#1b324d}
[data-testid="stAppViewContainer"]{background:#f5f8fc}
[data-testid="stHeader"]{background:#f5f8fc;opacity:.95}
.block-container{max-width:1380px;padding:2.1rem 2.8rem 3.2rem}
h1,h2,h3,h4{font-family:'Plus Jakarta Sans',sans-serif!important;letter-spacing:-.04em!important;color:#102743!important}
h1{font-size:2.0rem!important;font-weight:800!important;margin:0 0 .4rem!important}
h2{font-size:1.35rem!important}h3{font-size:1.12rem!important}
p{line-height:1.6}
[data-testid="stSidebar"]{background:#fff;border-right:1px solid #e4eaf2}
[data-testid="stSidebar"] [data-testid="stVerticalBlock"]{gap:.30rem}
[data-testid="stSidebar"]{min-width:265px;max-width:290px}
/* The navigation is made of plain clickable buttons, without radio circles. */
[data-testid="stSidebar"] .st-key-main_navigation{margin-bottom:-.12rem}
[data-testid="stSidebar"] .st-key-main_navigation button{
    min-height:43px!important;height:43px!important;width:100%;
    border:1px solid transparent!important;border-radius:10px!important;
    background:transparent!important;color:#4c607a!important;
    justify-content:flex-start!important;text-align:left!important;
    padding:0 13px!important;font-size:.86rem!important;font-weight:600!important;
    box-shadow:none!important;transition:background .15s ease,color .15s ease;
}
[data-testid="stSidebar"] .st-key-main_navigation button:hover{
    background:#f0f5fc!important;color:#1b4ca9!important;
}
[data-testid="stSidebar"] .st-key-main_navigation button[kind="primary"]{
    background:#eaf1ff!important;color:#2355b3!important;
    border:1px solid #dce8ff!important;box-shadow:inset 3px 0 0 #3567d8!important;
    font-weight:800!important;
}
[data-testid="stSidebar"] .st-key-main_navigation button [data-testid="stIconMaterial"]{font-size:1.15rem}
[data-testid="stSidebar"] .st-key-signout_panel button{
    width:100%;min-height:40px!important;border-radius:10px!important;
    border:1px solid #e0e8f2!important;background:#fff!important;
    color:#536780!important;justify-content:flex-start!important;
    font-weight:650!important;
}
[data-testid="stSidebar"] .st-key-signout_panel button:hover{background:#fff0f0!important;color:#bd3f4a!important;border-color:#f4d9dd!important}
[data-testid="stSidebar"] [data-testid="stExpander"]{box-shadow:none!important;border:1px solid #e4ebf3!important}
[data-testid="stSidebar"] [data-testid="stExpander"] summary{font-size:.83rem!important;font-weight:700!important}
[data-testid="stSidebar"] [data-testid="stFileUploader"]{background:#f7f9fc;border-radius:12px;border:1px dashed #c9d5e6;padding:5px}
[data-testid="stForm"]{background:#fff;border:1px solid #e2eaf3;border-radius:18px;padding:24px 28px;box-shadow:0 8px 28px rgba(27,60,104,.025)}
[data-testid="stExpander"]{border-radius:12px;border:1px solid #e4ebf3;background:#fff}
[data-testid="stDataFrame"]{border:1px solid #e7edf3;border-radius:13px;overflow:hidden;background:#fff}
[data-testid="stTabs"] [data-baseweb="tab-list"]{gap:10px}
[data-testid="stTabs"] [data-baseweb="tab"]{border-radius:9px 9px 0 0;padding:10px 14px;font-weight:600}
[data-testid="stButton"] button,[data-testid="stFormSubmitButton"] button,[data-testid="stDownloadButton"] button{border-radius:10px;font-weight:700;min-height:42px}
[data-testid="stButton"] button[kind="primary"],[data-testid="stFormSubmitButton"] button[kind="primary"],[data-testid="stDownloadButton"] button[kind="primary"]{background:#315fd4;border:1px solid #315fd4;color:#fff}
[data-testid="stButton"] button[kind="primary"]:hover,[data-testid="stFormSubmitButton"] button[kind="primary"]:hover{background:#214aae;border-color:#214aae}
[data-testid="stTextInput"] input,[data-testid="stNumberInput"] input,[data-testid="stTextArea"] textarea{border-radius:9px}
[data-testid="stAlert"]{border-radius:12px}
[data-testid="stMetric"]{border-radius:13px;border:1px solid var(--line);background:#fff;padding:17px 18px}
[data-testid="stMetricValue"]{color:var(--navy);font-weight:800!important}
.kicker{font-size:.72rem;font-weight:800;letter-spacing:.145em;color:#3567d8;text-transform:uppercase;margin-bottom:9px}
.intro{color:#63758a;font-size:.95rem;margin:0 0 1.55rem;max-width:780px}
.topline{display:flex;justify-content:space-between;align-items:center;margin:0 0 20px}
.topline .workspace{font-size:.77rem;font-weight:800;color:#607991;letter-spacing:.13em;text-transform:uppercase}
.topline .availability{background:#e7f7f1;color:#117e62;font-size:.73rem;font-weight:800;padding:7px 13px;border:1px solid #c7eddf;border-radius:90px}
.brand{display:flex;align-items:center;gap:11px;margin:12px 0 6px}
.brand-icon{width:42px;height:42px;background:linear-gradient(140deg,#2454bc,#3d7fe7);color:#fff;border-radius:12px;display:flex;align-items:center;justify-content:center;font-size:23px;box-shadow:0 5px 15px #2f61c122}
.brand-title{font-family:'Plus Jakarta Sans',sans-serif;font-size:1rem;color:#102743;font-weight:800;line-height:1.35}
.brand-subtitle{font-size:.74rem;color:#8492a4;margin-top:4px}
.nav-caption{font-size:.66rem;font-weight:800;letter-spacing:.14em;text-transform:uppercase;color:#8a99ab;margin:19px 0 9px}
.sidebar-divider{height:1px;background:#e8edf4;margin:16px 0 10px}
.sidebar-profile{display:flex;align-items:center;gap:10px;padding:12px;background:#f7f9fd;border:1px solid #e9eef5;border-radius:12px;margin:4px 0 6px}
.profile-avatar{display:flex;align-items:center;justify-content:center;width:32px;height:32px;border-radius:50%;background:#dce8ff;color:#2357b7;font-weight:800;font-size:.86rem}
.profile-name{color:#193859;font-size:.82rem;font-weight:800;overflow-wrap:anywhere}
.profile-role{color:#7c8da1;font-size:.7rem;margin-top:2px}
.logout-copy{font-size:.86rem;line-height:1.6;color:#64768a;margin:4px 0 14px}
.logout-note{font-size:.81rem;line-height:1.6;color:#986a14;background:#fff8e6;padding:13px 14px;border:1px solid #f6e9bc;border-radius:10px;margin:0 0 18px}
.hero{background:linear-gradient(111deg,#102947 0%,#183e71 58%,#2857a0 100%);border-radius:19px;padding:32px 36px 35px;display:flex;justify-content:space-between;align-items:center;gap:26px;overflow:hidden;position:relative;margin-bottom:23px;box-shadow:0 14px 34px #12335d15}
.hero:after{content:'';position:absolute;right:-115px;top:-105px;width:380px;height:380px;border:1px solid #ffffff2e;border-radius:50%;box-shadow:0 0 0 55px #ffffff07,0 0 0 112px #ffffff05}
.hero-main{position:relative;z-index:1;max-width:680px}
.hero .eyebrow{font-size:.72rem;color:#c3dcff;letter-spacing:.16em;font-weight:800;margin-bottom:11px}
.hero h1{font-size:2rem!important;line-height:1.25;color:#fff!important;margin:0 0 12px!important}
.hero p{color:#d1e0f4;font-size:.92rem;margin:0;max-width:570px}
.hero-right{z-index:1;flex-shrink:0;min-width:178px;border:1px solid #ffffff35;border-radius:14px;background:#ffffff13;padding:17px 20px}
.hero-right .count{font-size:1.6rem;font-weight:800;line-height:1.2;color:#fff}
.hero-right .tiny{font-size:.76rem;color:#d1e0f4}
.stats{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:15px;margin:0 0 30px}
.stat{background:#fff;border:1px solid #e1eaf4;border-radius:15px;padding:19px 20px 17px;box-shadow:0 6px 25px #1c4a7807;min-width:0}
.stat-head{display:flex;align-items:center;justify-content:space-between;gap:7px}
.stat-label{color:#687b91;font-size:.78rem;font-weight:600}
.stat-icon{height:34px;width:34px;border-radius:10px;background:#eef3ff;color:#3265d4;display:flex;align-items:center;justify-content:center;font-size:19px}
.stat-value{color:#112b4c;font-family:'Plus Jakarta Sans',sans-serif;font-size:2rem;letter-spacing:-.05em;font-weight:800;margin-top:11px;line-height:1.2}
.stat-foot{color:#8494a6;font-size:.72rem;margin-top:7px}
.section-heading{font-family:'Plus Jakarta Sans',sans-serif;font-size:1.18rem;font-weight:800;color:#12304f;margin:2px 0 14px}
.panel{background:white;border:1px solid #e2eaf3;border-radius:16px;padding:22px 23px;margin-bottom:17px;box-shadow:0 7px 28px #1a487305}
.panel-title{font-family:'Plus Jakarta Sans',sans-serif;font-weight:800;color:#14304c;font-size:1rem;margin:0 0 6px}
.panel-desc{font-size:.77rem;color:#8190a2;margin-bottom:20px}
.status-row{display:flex;align-items:center;justify-content:space-between;padding:13px 0;border-bottom:1px solid #edf1f5;gap:10px}
.status-row:last-child{border-bottom:none}
.status-meta{display:flex;align-items:center;gap:10px;font-size:.83rem;font-weight:600;color:#405873}
.dot{width:9px;height:9px;border-radius:99px;flex-shrink:0}
.count-pill{background:#f0f4f9;color:#24415f;padding:5px 10px;font-weight:800;font-size:.8rem;border-radius:8px}
.progress-track{height:9px;border-radius:99px;background:#e8eef7;overflow:hidden;margin-top:12px}
.progress-fill{height:100%;border-radius:99px;background:linear-gradient(90deg,#2d65d6,#63a3f5)}
.progress-caption{display:flex;justify-content:space-between;color:#7488a0;font-size:.76rem;margin-top:9px}
.empty-box{background:#f8faff;border:1px dashed #cfdded;border-radius:13px;padding:23px;text-align:center;color:#73879b;font-size:.84rem;line-height:1.7;margin:5px 0}
.step{display:flex;align-items:flex-start;gap:12px;padding:11px 0}
.step-num{flex:0 0 28px;height:28px;border-radius:9px;background:#edf3ff;color:#3867cf;font-weight:800;font-size:.81rem;display:flex;align-items:center;justify-content:center}
.step-body{font-size:.86rem;color:#31516e;line-height:1.4}
.step-body small{display:block;color:#8998aa;margin-top:4px;font-size:.76rem}
.podium{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px;margin-bottom:18px}
.podium-item{background:#fff;border:1px solid #e3eaf2;border-radius:14px;padding:19px 17px;min-width:0}
.podium-rank{font-size:.74rem;letter-spacing:.08em;font-weight:800;color:#5b7db4;margin-bottom:9px}
.podium-name{font-family:'Plus Jakarta Sans',sans-serif;font-size:1rem;font-weight:800;color:#193553;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.podium-meta{font-size:.77rem;color:#8190a3;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;margin-top:4px}
.podium-score{font-family:'Plus Jakarta Sans',sans-serif;font-size:1.6rem;font-weight:800;color:#2f62d3;margin-top:13px}
.podium-score span{font-size:.8rem;color:#8596ac}
.login-shell{margin:46px auto 20px;padding:0 4px;text-align:center}
.st-key-login_card [data-testid="stForm"]{padding:20px 23px 23px!important;border-radius:16px!important;box-shadow:0 13px 40px #1c476a12!important}
.st-key-login_card [data-testid="stForm"] [data-testid="stVerticalBlock"]{gap:.68rem!important}
.st-key-login_card [data-testid="stForm"] button{min-height:42px!important}
.login-icon{width:52px;height:52px;display:flex;align-items:center;justify-content:center;background:linear-gradient(120deg,#2359c1,#5290ef);color:#fff;border-radius:14px;margin:0 auto 17px;font-size:24px;box-shadow:0 11px 23px #2f66ca24}
.login-title{text-align:center;font-family:'Plus Jakarta Sans',sans-serif;font-size:1.48rem;font-weight:800;color:#12304d;margin-bottom:8px}
.login-desc{text-align:center;color:#75869b;font-size:.84rem;line-height:1.6;margin:0 auto 20px;max-width:350px}
.login-footer{text-align:center;color:#8494a7;font-size:.73rem;margin:17px 0}
@media(max-width:1000px){.block-container{padding:1.7rem 1rem 2.5rem}.stats{grid-template-columns:repeat(2,1fr)}.hero{padding:23px;display:block}.hero-right{display:none}}
@media(max-width:520px){.stats{grid-template-columns:1fr 1fr;gap:8px}.stat{padding:13px}.stat-value{font-size:1.5rem}.podium{grid-template-columns:1fr}.hero h1{font-size:1.55rem!important}}
</style>
""", unsafe_allow_html=True)


def secret(name: str) -> str:
    try:
        return str(st.secrets.get(name, "")).strip()
    except Exception:  # No secrets file, unavailable secrets, or platform configuration issue.
        return ""


APP_USERNAME = secret("APP_USERNAME")
APP_PASSWORD = secret("APP_PASSWORD")

# Fail closed: no candidate/profile screen may render until credentials are configured.
if not APP_USERNAME or not APP_PASSWORD:
    st.markdown('<div class="login-shell"><div class="login-icon">🔐</div><div class="login-title">Akses belum dikonfigurasi</div><div class="login-desc">Atur APP_USERNAME dan APP_PASSWORD di Streamlit Cloud → Settings → Secrets. Data kandidat tidak ditampilkan sebelum akses aktif.</div></div>', unsafe_allow_html=True)
    st.code('APP_USERNAME = "nama_akun_pewawancara"\nAPP_PASSWORD = "password_kuat_dan_unik"', language="toml")
    st.stop()

credential_revision = hashlib.sha256(f"{APP_USERNAME}\0{APP_PASSWORD}".encode()).hexdigest()
if st.session_state.get("authenticated_revision") != credential_revision:
    st.session_state["authenticated"] = False

if not st.session_state.get("authenticated", False):
    # Center both the heading AND the actual Streamlit form, not just the HTML header.
    # The middle column stays approximately 420-490px on a desktop screen.
    _, login_column, _ = st.columns([1.05, 1.10, 1.05], gap="small")
    with login_column:
        st.markdown('<div class="login-shell"><div class="login-icon">🔐</div><div class="login-title">Selamat datang kembali</div><div class="login-desc">Masuk untuk mengelola kandidat, melakukan wawancara, dan melihat hasil seleksi.</div></div>', unsafe_allow_html=True)
        if st.session_state.get("failed_logins", 0) >= 5:
            st.error("Terlalu banyak percobaan pada sesi ini. Tutup sesi browser dan coba kembali. Untuk proteksi lebih kuat, gunakan autentikasi akun organisasi (SSO/OIDC).")
            st.stop()
        with st.container(key="login_card"):
            with st.form("login_form", clear_on_submit=False):
                username = st.text_input("Username", placeholder="Masukkan username", autocomplete="username")
                password = st.text_input("Password", type="password", placeholder="Masukkan password", autocomplete="current-password")
                login_submit = st.form_submit_button("Masuk ke Sistem", type="primary", use_container_width=True)
        if login_submit:
            if hmac.compare_digest(username, APP_USERNAME) and hmac.compare_digest(password, APP_PASSWORD):
                st.session_state["authenticated"] = True
                st.session_state["authenticated_revision"] = credential_revision
                st.session_state["failed_logins"] = 0
                st.rerun()
            else:
                st.session_state["failed_logins"] = st.session_state.get("failed_logins", 0) + 1
                st.error("Username atau password tidak sesuai.")
        st.markdown('<div class="login-footer">INTERNAL RECRUITMENT PORTAL · AKSES TERBATAS</div>', unsafe_allow_html=True)
    st.stop()

# Fail closed: never use ephemeral session data as a silent database fallback.
if not configured():
    st.error("Koneksi PostgreSQL belum dikonfigurasi. Isi DATABASE_URL melalui Streamlit Cloud → Settings → Secrets.")
    st.stop()
try:
    db_candidates, db_interviews = load_records()
except DatabaseOperationError as exc:
    st.error(str(exc))
    st.stop()
st.session_state["candidates"] = db_candidates
st.session_state["interviews"] = db_interviews
if "page" not in st.session_state:
    st.session_state["page"] = "Mulai di Sini"

@st.cache_data(show_spinner=False)
def rubric():
    return get_rubric_and_questions()

weights, high, mid, questions, scale = rubric()
C = st.session_state["candidates"]
I = st.session_state["interviews"]
ranked = ranking_rows(C, I, weights, high, mid)


def section(kicker: str, title: str, description: str):
    st.markdown(f'<div class="kicker">{html.escape(kicker)}</div>', unsafe_allow_html=True)
    st.title(title)
    st.markdown(f'<div class="intro">{html.escape(description)}</div>', unsafe_allow_html=True)


def metric_card(label: str, value: str, note: str, icon: str):
    st.markdown(f"""<div class="stat"><div class="stat-head"><div class="stat-label">{html.escape(label)}</div><div class="stat-icon">{icon}</div></div><div class="stat-value">{html.escape(str(value))}</div><div class="stat-foot">{html.escape(note)}</div></div>""", unsafe_allow_html=True)


def blank_state(message: str):
    st.markdown(f'<div class="empty-box">{html.escape(message)}</div>', unsafe_allow_html=True)


def status_row(label: str, count: int, color: str):
    st.markdown(f'<div class="status-row"><div class="status-meta"><span class="dot" style="background:{color}"></span>{html.escape(label)}</div><span class="count-pill">{count}</span></div>', unsafe_allow_html=True)


def panel_head(title: str, subtitle: str):
    st.markdown(f'<div class="panel-title">{html.escape(title)}</div><div class="panel-desc">{html.escape(subtitle)}</div>', unsafe_allow_html=True)


def _date(text):
    if isinstance(text, date):
        return text
    if not text:
        return None
    try:
        return date.fromisoformat(str(text)[:10])
    except (TypeError, ValueError):
        return None


def _int(text, fallback=0):
    try:
        return int(float(text))
    except (TypeError, ValueError):
        return fallback


def _float(text, fallback=0.0):
    try:
        return float(text)
    except (TypeError, ValueError):
        return fallback


def _select(label: str, value: str, options: list[str]):
    choices = options.copy()
    if value and value not in choices:
        choices.append(value)
    return st.selectbox(label, choices, index=choices.index(value) if value in choices else 0)


def _shortlist_table(records: list[dict]):
    st.dataframe(records, use_container_width=True, hide_index=True)


@st.dialog("Konfirmasi keluar", width="small")
def show_logout_dialog():
    st.markdown('<p class="logout-copy">Anda akan keluar dari sistem Interview Peserta Magang.</p><div class="logout-note">Data kandidat dan hasil interview yang sudah disimpan tetap tersedia di PostgreSQL.</div>', unsafe_allow_html=True)
    cancel, confirm = st.columns(2, gap="small")
    with cancel:
        if st.button("Batal", key="cancel_logout", use_container_width=True):
            st.rerun()
    with confirm:
        if st.button("Ya, keluar", key="confirm_logout", type="primary", use_container_width=True):
            st.session_state.clear()
            st.rerun()


with st.sidebar:
    st.markdown('<div class="brand"><div class="brand-icon">◈</div><div><div class="brand-title">Internship Hub</div><div class="brand-subtitle">Recruitment & Assessment</div></div></div>', unsafe_allow_html=True)
    st.markdown('<div class="sidebar-divider"></div><div class="nav-caption">MENU UTAMA</div>', unsafe_allow_html=True)
    nav_items = [
        ("Mulai di Sini", "Dashboard", ":material/dashboard:", "nav_dashboard"),
        ("Data Kandidat", "Data Kandidat", ":material/groups:", "nav_candidates"),
        ("Penilaian Interview", "Penilaian Interview", ":material/assignment_turned_in:", "nav_interview"),
        ("Summary & Ranking", "Summary & Ranking", ":material/leaderboard:", "nav_ranking"),
        ("Panduan & Rubrik", "Panduan & Rubrik", ":material/menu_book:", "nav_guide"),
    ]
    with st.container(key="main_navigation"):
        for page_name, display_name, page_icon, widget_key in nav_items:
            is_active = st.session_state.get("page") == page_name
            if st.button(display_name, key=widget_key, icon=page_icon,
                         use_container_width=True, type="primary" if is_active else "secondary"):
                st.session_state["page"] = page_name
                st.rerun()
    menu = st.session_state["page"]

    st.markdown('<div class="sidebar-divider"></div><div class="nav-caption">PENGELOLAAN DATA</div>', unsafe_allow_html=True)
    with st.expander("Backup & Impor Excel", expanded=False, icon=":material/folder_open:"):
        uploaded = st.file_uploader("Pilih file backup Excel", type=["xlsx"], help="Impor menggabungkan data berdasarkan ID. ID yang sudah ada akan diperbarui, data lain tetap disimpan.")
        if uploaded is not None:
            st.caption("Impor akan menambahkan kandidat baru dan memperbarui ID yang sama. Data lain tidak dihapus.")
            confirm_import = st.checkbox("Saya memahami bahwa data dengan ID sama akan diperbarui", key="db_confirm_import")
            if st.button("Impor ke PostgreSQL", use_container_width=True, disabled=not confirm_import):
                try:
                    new_c, new_i = read_xlsx(uploaded.getvalue())
                    import_backup(new_c, new_i)
                    st.success(f"{len(new_c)} kandidat dan {len(new_i)} hasil interview berhasil diimpor ke PostgreSQL.")
                    st.rerun()
                except (ValueError, DatabaseOperationError) as exc:
                    st.error(f"Impor gagal: {exc}")
        try:
            excel_bytes = export_xlsx(C, I)
            st.download_button(
                "Download Backup Excel",
                data=excel_bytes,
                file_name=f"Hasil_Interview_Magang_{date.today().isoformat()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                icon=":material/download:",
            )
        except ValueError as exc:
            st.error(str(exc))
        st.caption("Data tersimpan di PostgreSQL. Download Excel sebagai backup berkala.")

    st.markdown('<div class="sidebar-divider"></div><div class="nav-caption">AKUN PEWAWANCARA</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="sidebar-profile"><div class="profile-avatar">' + html.escape(APP_USERNAME[:1].upper()) +
        '</div><div><div class="profile-name">' + html.escape(APP_USERNAME) +
        '</div><div class="profile-role">Akun aktif</div></div></div>',
        unsafe_allow_html=True,
    )
    with st.container(key="signout_panel"):
        if st.button("Keluar dari Sistem", key="signout_button", icon=":material/logout:", use_container_width=True):
            show_logout_dialog()

# Data may have been replaced in the sidebar by an import. Rebind to the newest dict.
C = st.session_state["candidates"]
I = st.session_state["interviews"]
ranked = ranking_rows(C, I, weights, high, mid)

if menu == "Mulai di Sini":
    fully_scored = len(ranked)
    recommended = sum(r["Rekomendasi"] == "Direkomendasikan" for r in ranked)
    pending = len(C) - fully_scored
    avg = sum(r["Skor"] for r in ranked) / fully_scored if fully_scored else 0
    rate = (fully_scored / len(C) * 100) if C else 0

    st.markdown('<div class="topline"><span class="workspace">RECRUITMENT / OVERVIEW</span><span class="availability">● Workspace aktif</span></div>', unsafe_allow_html=True)
    st.markdown(f"""<div class="hero"><div class="hero-main"><div class="eyebrow">INTERNSHIP RECRUITMENT  ·  2026</div><h1>Kelola kandidat.<br>Temukan talenta terbaik.</h1><p>Satu ruang kerja untuk pendataan peserta magang, wawancara, evaluasi kompetensi, dan pemeringkatan kandidat yang objektif.</p></div><div class="hero-right"><div class="tiny">Progress Interview</div><div class="count">{rate:.0f}%</div><div class="tiny">{fully_scored} dari {len(C)} peserta dinilai</div></div></div>""", unsafe_allow_html=True)

    c1,c2,c3,c4 = st.columns(4, gap="medium")
    with c1: metric_card("Total Kandidat", len(C), "Terdaftar dalam sistem", "♙")
    with c2: metric_card("Interview Selesai", fully_scored, "Seluruh aspek telah dinilai", "✓")
    with c3: metric_card("Direkomendasikan", recommended, "Berdasarkan skor otomatis", "✦")
    with c4: metric_card("Rata-rata Skor", f"{avg:.1f}", "Dari total 100 poin", "◈")

    left,right = st.columns([1.25,1], gap="large")
    with left:
        st.markdown('<div class="section-heading">Monitoring Seleksi</div>', unsafe_allow_html=True)
        with st.container(border=True):
            panel_head("Progres Interview", "Pemantauan kelengkapan proses penilaian")
            st.markdown(f'<div style="font-family:Plus Jakarta Sans,sans-serif;font-size:2rem;font-weight:800;color:#12304d">{fully_scored}<span style="font-size:.95rem;color:#8497aa;font-weight:500"> / {len(C)} kandidat</span></div>', unsafe_allow_html=True)
            st.markdown(f'<div class="progress-track"><div class="progress-fill" style="width:{rate:.2f}%"></div></div><div class="progress-caption"><span>{rate:.0f}% selesai</span><span>{pending} belum lengkap</span></div>', unsafe_allow_html=True)
            st.write("")
            status_row("Interview selesai", fully_scored, "#23b38c")
            status_row("Belum lengkap / belum mulai", pending, "#f0ab48")
            status_row("Total data peserta", len(C), "#426bd4")
        with st.container(border=True):
            panel_head("Kandidat Berperingkat", "Skor tertinggi dari hasil interview lengkap")
            if ranked:
                for r in ranked[:5]:
                    st.markdown(f'<div class="status-row"><div class="status-meta"><span class="count-pill">#{r["Ranking"]}</span>{html.escape(r["Nama Kandidat"][:46])}</div><span style="font-weight:800;color:#2657b7">{r["Skor"]:.1f}</span></div>', unsafe_allow_html=True)
            else:
                blank_state("Peringkat akan muncul setelah penilaian tujuh aspek pertama selesai.")
    with right:
        st.markdown('<div class="section-heading">Ringkasan Keputusan</div>', unsafe_allow_html=True)
        with st.container(border=True):
            panel_head("Rekomendasi Otomatis", "Hasil sistem berdasarkan bobot pada template Excel")
            status_row("Direkomendasikan", recommended, "#23b38c")
            status_row("Dipertimbangkan", sum(r["Rekomendasi"] == "Dipertimbangkan" for r in ranked), "#e0a23a")
            status_row("Belum Direkomendasikan", sum(r["Rekomendasi"] == "Belum Direkomendasikan" for r in ranked), "#eb7180")
        with st.container(border=True):
            panel_head("Keputusan Pewawancara", "Hasil final yang ditetapkan secara manual")
            for label,color in [("Lolos","#23b38c"),("Cadangan","#e0a23a"),("Tidak Lolos","#eb7180"),("Menunggu","#94a3b8")]:
                status_row(label, sum((I.get(cid,{ }).get("keputusan") or "Menunggu") == label for cid in C), color)
        with st.container(border=True):
            panel_head("Langkah Selanjutnya", "Ikuti alur yang terstruktur")
            for no,title,desc in [("01","Daftarkan kandidat","Input identitas, perguruan tinggi dan pengalaman"),("02","Lakukan interview","Isi jawaban, catatan, dan tujuh aspek nilai"),("03","Tinjau hasil","Bandingkan ranking dan putuskan status akhir")]:
                st.markdown(f'<div class="step"><div class="step-num">{no}</div><div class="step-body"><b>{title}</b><small>{desc}</small></div></div>', unsafe_allow_html=True)
    st.caption("Data kandidat dan hasil interview yang disimpan tersedia kembali saat login. Download Excel untuk cadangan berkala.")

elif menu == "Data Kandidat":
    section(
        "CANDIDATE DATABASE", "Data Kandidat",
        "Lengkapi data satu peserta dari identitas hingga preferensi magang, kemudian simpan sekali di bagian paling bawah.",
    )

    # After saving a NEW candidate, stay on the newly saved profile instead of
    # unexpectedly showing the blank form for the NEXT candidate (MG-002, etc.).
    options = ["➕ Tambah Kandidat Baru"] + sorted(C.keys())
    pending = st.session_state.pop("candidate_pending_selection", None)
    if pending in options:
        st.session_state["candidate_choice"] = pending
    if st.session_state.get("candidate_choice") not in options:
        st.session_state["candidate_choice"] = options[0]

    choice = st.selectbox(
        "Pilih peserta untuk mengisi atau mengedit data",
        options,
        format_func=lambda x: f"{x} — {C[x].get('nama', '')}" if x in C else x,
        key="candidate_choice",
    )
    existing = choice in C
    cid = choice if existing else candidate_id_next(C)
    initial = C[choice].copy() if existing else {}
    if message := st.session_state.pop("candidate_saved_notice", None):
        st.success(message)

    st.caption(f"ID Kandidat: **{cid}** · Kapasitas ekspor Excel: {len(C)}/{MAX_KANDIDAT}")
    if existing:
        st.info("Anda sedang membuka profil kandidat yang sudah disimpan. Perbaiki data yang diperlukan, lalu klik **Simpan Perubahan** di bagian akhir formulir.")
    else:
        st.info("Isi seluruh bagian formulir berikut untuk **satu kandidat**. Data baru dikirim ke PostgreSQL setelah Anda menekan **Simpan Data Kandidat** di bagian paling bawah.")

    with st.form(f"form_candidate_{cid}", clear_on_submit=False):
        result = {"id": cid}

        st.markdown("### 01 · Identitas & Kontak")
        st.caption("Informasi dasar dan kontak kandidat.")
        a, b = st.columns(2, gap="large")
        with a:
            result["nama"] = st.text_input("Nama Lengkap *", value=initial.get("nama", ""))
            result["telepon"] = st.text_input("No. HP / WA", value=initial.get("telepon", ""))
            result["alamat"] = st.text_area("Alamat Domisili", value=initial.get("alamat", ""), height=85)
        with b:
            result["email"] = st.text_input("Email", value=initial.get("email", ""))
            result["kota"] = st.text_input("Kota / Kabupaten", value=initial.get("kota", ""))

        st.divider()
        st.markdown("### 02 · Pendidikan")
        st.caption("Asal perguruan tinggi, jenjang, dan informasi akademik.")
        a, b = st.columns(2, gap="large")
        with a:
            result["kampus"] = st.text_input("Asal Perguruan Tinggi", value=initial.get("kampus", ""))
            result["prodi"] = st.text_input("Program Studi / Jurusan", value=initial.get("prodi", ""))
            result["ipk"] = st.number_input(
                "IPK (0,00–4,00)", min_value=0.0, max_value=4.0,
                value=min(max(_float(initial.get("ipk")), 0.0), 4.0),
                step=0.01, format="%.2f",
            )
        with b:
            result["jenjang"] = _select("Jenjang", initial.get("jenjang", ""), ["", "D3", "D4", "S1", "S2", "Lainnya"])
            result["semester"] = st.number_input(
                "Semester", min_value=0, max_value=16,
                value=min(max(_int(initial.get("semester")), 0), 16), step=1,
            )
            result["status_kuliah"] = _select(
                "Status Perkuliahan", initial.get("status_kuliah", ""),
                ["", "Aktif", "Cuti", "Lulus", "Lainnya"],
            )

        st.divider()
        st.markdown("### 03 · Pengalaman & Kompetensi")
        st.caption("Pengalaman organisasi, proyek, keterampilan, dan sertifikasi.")
        a, b = st.columns(2, gap="large")
        with a:
            result["pengalaman_kerja"] = st.text_area(
                "Pengalaman Magang / Kerja", value=initial.get("pengalaman_kerja", ""), height=110,
            )
            result["pengalaman_organisasi"] = st.text_area(
                "Pengalaman Organisasi", value=initial.get("pengalaman_organisasi", ""), height=110,
            )
            result["portofolio"] = st.text_area(
                "Proyek / Portofolio", value=initial.get("portofolio", ""), height=110,
            )
        with b:
            result["keahlian"] = st.text_area(
                "Keahlian Utama", value=initial.get("keahlian", ""), height=110,
            )
            result["tools"] = st.text_area(
                "Tools / Software", value=initial.get("tools", ""), height=110,
            )
            result["sertifikasi"] = st.text_area(
                "Sertifikasi / Pelatihan", value=initial.get("sertifikasi", ""), height=110,
            )
        result["link"] = st.text_input("Tautan LinkedIn / Portofolio", value=initial.get("link", ""))

        st.divider()
        st.markdown("### 04 · Preferensi Magang & Berkas")
        st.caption("Bidang yang diminati, ketersediaan, dan motivasi magang.")
        a, b = st.columns(2, gap="large")
        with a:
            result["bidang"] = st.text_input("Posisi / Bidang Magang", value=initial.get("bidang", ""))
            result["wfo"] = _select(
                "Bersedia WFO", initial.get("wfo", ""),
                ["", "Ya", "Tidak", "Hybrid", "Negosiasi"],
            )
            result["durasi"] = st.number_input(
                "Durasi (bulan)", min_value=0, max_value=24, step=1,
                value=min(max(_int(initial.get("durasi")), 0), 24),
            )
        with b:
            result["mulai"] = st.date_input(
                "Dapat Mulai Tanggal", value=_date(initial.get("mulai")),
                format="DD/MM/YYYY",
            )
            result["status_cv"] = _select(
                "Status CV / Berkas", initial.get("status_cv", ""),
                ["", "Lengkap", "Belum Lengkap", "Proses Verifikasi"],
            )
        result["motivasi"] = st.text_area(
            "Motivasi / Harapan Magang", value=initial.get("motivasi", ""), height=105,
        )

        st.divider()
        st.markdown("**Selesai mengisi?** Periksa kembali nama, alamat, perguruan tinggi, dan pengalaman sebelum menyimpan.")
        submitted = st.form_submit_button(
            "💾 Simpan Perubahan Kandidat" if existing else "💾 Simpan Data Kandidat",
            type="primary", use_container_width=True,
        )

    if submitted:
        if not result["nama"].strip():
            st.error("Nama lengkap wajib diisi sebelum menyimpan.")
        elif not existing and len(C) >= MAX_KANDIDAT:
            st.error("Kapasitas 150 kandidat pada template ekspor Excel telah tercapai.")
        else:
            result["nama"] = result["nama"].strip()
            result["mulai"] = result["mulai"].isoformat() if result["mulai"] else ""
            try:
                # One DB write/transaction for all 25 candidate fields.
                save_candidate(result, create=not existing)
            except DatabaseOperationError as exc:
                st.error(str(exc))
            else:
                st.session_state["candidate_saved_notice"] = (
                    f"Data {result['nama']} ({cid}) berhasil disimpan lengkap ke PostgreSQL. "
                    "Anda dapat mengeditnya lagi tanpa membuat ID baru."
                )
                st.session_state["candidate_pending_selection"] = cid
                st.rerun()

    st.divider()
    st.subheader("Daftar Kandidat")
    search = st.text_input("Cari nama, kampus, atau posisi", placeholder="Masukkan kata kunci...")
    table = []
    for key, val in sorted(C.items()):
        searchable = " ".join(str(val.get(field, "")) for field in ["nama", "kampus", "prodi", "bidang"]).lower()
        if search and search.lower() not in searchable:
            continue
        score, status = calculate(I.get(key), weights, high, mid)
        table.append({
            "ID": key,
            "Nama": val.get("nama", ""),
            "Perguruan Tinggi": val.get("kampus", ""),
            "Program Studi": val.get("prodi", ""),
            "IPK": val.get("ipk", ""),
            "Posisi": val.get("bidang", ""),
            "Status Interview": status,
        })
    if table:
        _shortlist_table(table)
    else:
        st.info("Belum ada data kandidat. Silakan tambah kandidat melalui form di atas.")


elif menu == "Penilaian Interview":
    section("INTERVIEW ASSESSMENT", "Penilaian Interview", "Catat hasil wawancara, tentukan skor kompetensi dengan rubrik yang konsisten, kemudian pilih keputusan akhir pewawancara.")
    if not C:
        st.info("Belum ada kandidat. Isi menu Data Kandidat terlebih dahulu.")
    else:
        cid = st.selectbox("Pilih peserta magang", sorted(C.keys()), format_func=lambda x: f"{C[x].get('nama', '')}  ·  {x}")
        person = C[cid]
        old = I.get(cid, {})
        st.markdown(f"**{person.get('nama', '')}** · {person.get('kampus', '—')} · {person.get('prodi', '—')} · {person.get('bidang', '—')}")
        st.caption("Skor baru dihitung jika seluruh tujuh aspek sudah dinilai. Pilihan 'Belum Dinilai' tidak dihitung sebagai nilai 0.")
        with st.form(f"interview_{cid}"):
            a, b = st.columns(2)
            with a:
                interview_date = st.date_input("Tanggal Interview", value=_date(old.get("tanggal")) or date.today(), format="DD/MM/YYYY")
            with b:
                interviewer = st.text_input("Pewawancara", value=old.get("pewawancara", ""))
            st.markdown("#### Nilai Kompetensi")
            current_score = {}
            left, right = st.columns(2, gap="large")
            for i, (key, label, _) in enumerate(ASPEK):
                with (left if i < 4 else right):
                    val = old.get(key, 0)
                    opts = [0, 1, 2, 3, 4, 5]
                    current_score[key] = st.selectbox(
                        f"{label} · Bobot {weights[key]:.0%}",
                        options=opts,
                        index=opts.index(val) if val in opts else 0,
                        format_func=lambda x: "— Belum Dinilai —" if x == 0 else f"{x} / 5",
                    )
            st.divider()
            st.markdown("#### Hasil Wawancara & Tindak Lanjut")
            a, b = st.columns(2)
            with a:
                strength = st.text_area("Kelebihan / Kekuatan", value=old.get("kelebihan", ""), height=105)
                final = _select("Keputusan Akhir Pewawancara", old.get("keputusan", "Menunggu"), DECISIONS)
            with b:
                needs_work = st.text_area("Hal yang Perlu Ditingkatkan", value=old.get("peningkatan", ""), height=105)
            notes = st.text_area("Catatan / Tindak Lanjut", value=old.get("catatan", ""), height=115)
            submitted = st.form_submit_button("💾 Simpan Hasil Interview", type="primary", use_container_width=True)
        if submitted:
            saved = {
                "id": cid, "tanggal": interview_date.isoformat(), "pewawancara": interviewer,
                "keputusan": final, "kelebihan": strength, "peningkatan": needs_work, "catatan": notes,
                **current_score,
            }
            try:
                save_interview(saved)
                I[cid] = saved
                sc, recom = calculate(saved, weights, high, mid)
                if sc is None:
                    st.success("Draft wawancara tersimpan di PostgreSQL. Lengkapi tujuh aspek penilaian untuk menampilkan skor dan ranking.")
                else:
                    st.success(f"Wawancara tersimpan di PostgreSQL · Skor: {sc:.1f}/100 · {recom}.")
            except DatabaseOperationError as exc:
                st.error(str(exc))
        show = I.get(cid)
        if show:
            sc, recom = calculate(show, weights, high, mid)
            with st.container(border=True):
                st.markdown("**Hasil Penilaian Tersimpan**")
                if sc is None:
                    st.info("Masih berupa draft: tujuh aspek nilai belum lengkap.")
                else:
                    c1, c2 = st.columns(2)
                    c1.metric("Skor Akhir", f"{sc:.1f} / 100")
                    c2.metric("Rekomendasi Sistem", recom)
                    st.caption(f"Ambang rekomendasi sesuai Excel: ≥{high:g} direkomendasikan, ≥{mid:g} dipertimbangkan, di bawahnya belum direkomendasikan. Keputusan akhir diisi pewawancara.")
        with st.expander("Lihat contoh pertanyaan wawancara (sesuai template)"):
            for q in questions:
                st.markdown(f"**{q['no']}. {q['area']}** — {q['pertanyaan']}")
                st.caption(f"Pendalaman: {q['pendalaman']} · Fokus: {q['dinilai']}")

elif menu == "Summary & Ranking":
    section("EVALUATION / ANALYTICS", "Summary & Ranking Kandidat", "Pantau hasil wawancara, bandingkan skor, dan susun shortlist berdasarkan bobot penilaian yang sama dengan template Excel.")
    total = len(C)
    # Use columns so cards remain responsive in Streamlit layout.
    c1,c2,c3,c4 = st.columns(4, gap="medium")
    with c1: metric_card("Total Kandidat",total,"Seluruh peserta terdaftar","♙")
    with c2: metric_card("Dinilai Lengkap",len(ranked),"Sudah masuk ranking","✓")
    with c3: metric_card("Belum Lengkap",total-len(ranked),"Butuh tindak lanjut","◷")
    with c4: metric_card("Skor Tertinggi",f"{ranked[0]['Skor']:.1f}" if ranked else "—","Nilai maksimum 100","✦")

    if ranked:
        st.markdown('<div class="section-heading">Top Kandidat</div>',unsafe_allow_html=True)
        podium = []
        for idx,r in enumerate(ranked[:3]):
            medal = ["01 / TOP SCORE","02 / RUNNER UP","03 / TOP 3"][idx]
            podium.append(f'<div class="podium-item"><div class="podium-rank">{medal}</div><div class="podium-name">{html.escape(str(r["Nama Kandidat"]))}</div><div class="podium-meta">{html.escape(str(r["Perguruan Tinggi"]))}</div><div class="podium-score">{r["Skor"]:.1f}<span> / 100</span></div></div>')
        st.markdown('<div class="podium">'+''.join(podium)+'</div>',unsafe_allow_html=True)
        st.markdown('<div class="section-heading">Perbandingan Skor</div>',unsafe_allow_html=True)
        plot = [{"Nama":r["Nama Kandidat"],"Skor":r["Skor"],"Rekomendasi":r["Rekomendasi"]} for r in ranked[:12]]
        palette = alt.Scale(domain=["Direkomendasikan","Dipertimbangkan","Belum Direkomendasikan"], range=["#315fd4","#e0a23a","#e77e85"])
        chart = (alt.Chart(alt.Data(values=plot)).mark_bar(cornerRadiusEnd=6,size=20)
            .encode(x=alt.X("Skor:Q",scale=alt.Scale(domain=[0,100]),title="Nilai / 100"),
                    y=alt.Y("Nama:N",sort="-x",title=None),
                    color=alt.Color("Rekomendasi:N",scale=palette,legend=alt.Legend(orient="bottom")),
                    tooltip=[alt.Tooltip("Nama:N"),alt.Tooltip("Skor:Q",format=".1f"),"Rekomendasi:N"])
            .properties(height=max(220,len(plot)*36),background="#ffffff").configure_view(stroke=None)
            .configure_axis(labelColor="#5c6e86",titleColor="#5c6e86",gridColor="#eaf0f6"))
        with st.container(border=True):
            st.altair_chart(chart,use_container_width=True)
        st.markdown('<div class="section-heading">Tabel Peringkat Lengkap</div>',unsafe_allow_html=True)
        f1,f2 = st.columns(2)
        with f1:
            cats=["Semua","Direkomendasikan","Dipertimbangkan","Belum Direkomendasikan"]
            reco_filter=st.selectbox("Filter rekomendasi",cats)
        with f2:
            positions=sorted({str(r["Posisi Magang"]) for r in ranked if str(r["Posisi Magang"]).strip()})
            pos_filter=st.selectbox("Filter posisi magang",["Semua",*positions])
        filtered=[r for r in ranked if (reco_filter=="Semua" or r["Rekomendasi"]==reco_filter) and (pos_filter=="Semua" or r["Posisi Magang"]==pos_filter)]
        keys=["Ranking","ID","Nama Kandidat","Perguruan Tinggi","Program Studi","Posisi Magang","IPK","Skor","Rekomendasi","Keputusan Akhir"]
        st.dataframe([{k:r[k] for k in keys} for r in filtered],use_container_width=True,hide_index=True,
                     column_config={"Skor":st.column_config.ProgressColumn("Skor",min_value=0,max_value=100,format="%.1f")})
        st.caption("Peringkat dihitung dari seluruh kandidat yang dinilai lengkap. Bila nilainya sama, peringkat sama. Keputusan akhir tetap ditetapkan pewawancara.")
        if filtered:
            stream=StringIO()
            writer=csv.DictWriter(stream,fieldnames=keys)
            writer.writeheader()
            for r in filtered: writer.writerow({k:r[k] for k in keys})
            st.download_button("⬇  Download Ranking (CSV)",stream.getvalue().encode("utf-8-sig"),file_name="ranking_peserta_magang.csv",mime="text/csv")
    else:
        blank_state("Belum tersedia ranking. Tambahkan kandidat, kemudian lengkapi tujuh aspek penilaian melalui menu Penilaian Interview.")

    st.divider()
    st.markdown('<div class="section-heading">Status Seluruh Kandidat</div>',unsafe_allow_html=True)
    progresses=[]
    for cid,p in sorted(C.items()):
        interview=I.get(cid)
        sc,reco=calculate(interview,weights,high,mid)
        filled=sum(1 for key,_,_ in ASPEK if interview and interview.get(key) in range(1,6))
        progresses.append({"ID":cid,"Nama":p.get("nama",""),"Aspek Dinilai":filled,"Status":"Selesai" if sc is not None else ("Draft" if interview else "Belum Interview"),"Skor":sc,"Keputusan":(interview or {}).get("keputusan","—")})
    if progresses:
        st.dataframe(progresses,use_container_width=True,hide_index=True,
                     column_config={"Aspek Dinilai":st.column_config.ProgressColumn("Aspek Dinilai",min_value=0,max_value=7,format="%d/7")})
    else:
        blank_state("Data peserta belum tersedia. Silakan mulai melalui menu Data Kandidat.")

elif menu == "Panduan & Rubrik":
    section("ASSESSMENT GUIDELINES", "Panduan & Rubrik", "Referensi pewawancara: bobot kompetensi, skala nilai, dan pertanyaan yang digunakan secara konsisten untuk setiap peserta.")
    st.subheader("Bobot dan Fokus Penilaian")
    focus = [
        "Menyampaikan gagasan secara jelas dan menjawab pertanyaan dengan runtut.",
        "Alasan magang, tujuan belajar, dan pemahaman terhadap peran.",
        "Dasar pengetahuan, keterampilan relevan, dan penerapannya.",
        "Logika berpikir, analisis persoalan, dan pendekatan solusi.",
        "Kolaborasi, tanggung jawab, dan komunikasi dalam tim.",
        "Inisiatif, sikap terhadap umpan balik, dan adaptasi.",
        "Etika, kedisiplinan, komitmen, dan kejelasan ketersediaan.",
    ]
    st.dataframe([{"Aspek": label, "Bobot": f"{weights[key]:.0%}", "Fokus": fokus} for (key, label, _), fokus in zip(ASPEK, focus)], hide_index=True, use_container_width=True)
    st.caption(f"Skor akhir = Σ(nilai 1–5 × bobot) × 20. Direkomendasikan: ≥{high:g}; dipertimbangkan: ≥{mid:g}.")
    st.subheader("Skala Penilaian 1–5")
    st.dataframe([{"Nilai": x["nilai"], "Kategori": x["kategori"], "Panduan": x["panduan"]} for x in scale], hide_index=True, use_container_width=True)
    st.subheader("12 Pertanyaan Panduan Interview")
    for q in questions:
        with st.expander(f"{q['no']}. {q['area']} — {q['pertanyaan']}"):
            st.markdown(f"**Pertanyaan pendalaman:** {q['pendalaman']}")
            st.markdown(f"**Hal yang dinilai:** {q['dinilai']}")
    st.info("Gunakan pertanyaan inti yang seragam untuk menjaga konsistensi. Skor sistem adalah alat bantu, bukan keputusan akhir penerimaan peserta magang.")
