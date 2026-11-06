"""Sistem Pendataan & Interview Peserta Magang — Streamlit.

Run: streamlit run app.py
Storage: st.session_state + manual Excel import/export; NO server database.
"""
from __future__ import annotations

import csv
from datetime import date, datetime
from io import StringIO
from pathlib import Path
import hmac

import streamlit as st

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
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html, body, [class*="css"], [data-testid="stApp"] { font-family: Inter, sans-serif; }
[data-testid="stAppViewContainer"] {background: linear-gradient(126deg,#fff 0%,#f7fbff 55%,#eaf3ff 100%);}
[data-testid="stHeader"] {background: transparent;}
[data-testid="stSidebar"] {background:#f3f7fd; border-right:1px solid #e3ebf5;}
.block-container {max-width:1320px; padding-top:2rem; padding-bottom:3rem;}
h1,h2,h3 {color:#132b4b!important; letter-spacing:-.035em;}
[data-testid="stMetric"] {background:#fff;border:1px solid #e5edf7;border-radius:14px;padding:16px 20px;box-shadow:0 3px 16px rgba(23,55,102,.04);}
[data-testid="stMetricValue"] {font-size:1.7rem;color:#183960;}
[data-testid="stForm"] {background:rgba(255,255,255,.9);padding:20px 24px;border:1px solid #e5edf7;border-radius:16px;}
[data-testid="stExpander"] {border:1px solid #dbe8f5;border-radius:12px;background:#fff;}
[data-testid="stDataFrame"] {border-radius:12px;overflow:hidden;}
.intro {color:#61738d;font-size:1rem;margin:-.35rem 0 1.4rem 0;}
.kicker {font-size:.78rem;letter-spacing:.12em;text-transform:uppercase;font-weight:800;color:#3469aa;margin-bottom:4px;}
.tagline {background:#fff;border:1px solid #e2eaf5;border-radius:12px;padding:16px 20px;margin-bottom:18px;}
.small-label {color:#697b91;font-size:.83rem;}
hr {border-color:#e5edf7;}
</style>
""", unsafe_allow_html=True)


def app_password() -> str:
    try:
        return str(st.secrets.get("APP_PASSWORD", ""))
    except Exception:
        return ""


PASSWORD = app_password()
if PASSWORD:
    if not st.session_state.get("authenticated", False):
        st.title("🔐 Akses Pewawancara")
        st.caption("Masukkan sandi untuk membuka sistem seleksi magang.")
        with st.form("login"):
            password_entered = st.text_input("Sandi aplikasi", type="password")
            submitted_login = st.form_submit_button("Masuk", type="primary")
        if submitted_login:
            if hmac.compare_digest(password_entered, PASSWORD):
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("Sandi tidak sesuai.")
        st.stop()

if "candidates" not in st.session_state:
    st.session_state["candidates"] = {}
if "interviews" not in st.session_state:
    st.session_state["interviews"] = {}
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
    st.markdown(f'<div class="kicker">{kicker}</div>', unsafe_allow_html=True)
    st.title(title)
    st.markdown(f'<div class="intro">{description}</div>', unsafe_allow_html=True)


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


with st.sidebar:
    st.markdown("### 📋 Internship Recruitment")
    st.caption("Pendataan & Penilaian Interview")
    st.divider()
    menu = st.radio(
        "NAVIGASI",
        ["Mulai di Sini", "Data Kandidat", "Penilaian Interview", "Summary & Ranking", "Panduan & Rubrik"],
        key="page",
    )
    st.divider()
    st.markdown("**Data & Backup**")
    uploaded = st.file_uploader("Impor data dari template Excel", type=["xlsx"], help="Memuat file akan mengganti data di sesi ini. Simpan hasil sebelumnya terlebih dahulu.")
    if uploaded is not None:
        if st.button("Muat Data Excel", use_container_width=True):
            try:
                new_c, new_i = read_xlsx(uploaded.getvalue())
                st.session_state["candidates"] = new_c
                st.session_state["interviews"] = new_i
                st.success(f"{len(new_c)} kandidat dan {len(new_i)} hasil interview berhasil dimuat.")
                st.rerun()
            except Exception as exc:
                st.error(f"Impor gagal: {exc}")
    try:
        excel_bytes = export_xlsx(C, I)
        st.download_button(
            "⬇️ Download Backup Excel",
            data=excel_bytes,
            file_name=f"Hasil_Interview_Magang_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )
    except ValueError as exc:
        st.error(str(exc))
    st.caption("Belum memakai database. Data baru tersimpan permanen setelah file Excel diunduh.")
    if not PASSWORD:
        st.warning("Belum ada sandi akses. Jika aplikasi dipublikasikan, atur APP_PASSWORD di Streamlit Secrets dan batasi pembagian tautan.")
    elif st.button("Keluar", use_container_width=True):
        st.session_state["authenticated"] = False
        st.rerun()

# Data may have been replaced in the sidebar by an import. Rebind to the newest dict.
C = st.session_state["candidates"]
I = st.session_state["interviews"]
ranked = ranking_rows(C, I, weights, high, mid)

if menu == "Mulai di Sini":
    section("OVERVIEW", "Dashboard Interview Peserta Magang", "Ringkasan jumlah kandidat dan hasil penilaian, mengikuti struktur sheet 'Mulai di Sini' pada template Excel.")
    fully_scored = len(ranked)
    recommended = sum(r["Rekomendasi"] == "Direkomendasikan" for r in ranked)
    avg = sum(r["Skor"] for r in ranked) / fully_scored if fully_scored else 0
    cols = st.columns(4)
    for col, label, val in zip(cols, ["Total Pelamar", "Sudah Dinilai", "Direkomendasikan", "Rata-rata Nilai"], [len(C), fully_scored, recommended, f"{avg:.1f}/100"]):
        with col:
            st.metric(label, val)
    st.write("")
    left, right = st.columns([1.35, 1], gap="large")
    with left:
        st.subheader("Ringkasan Rekomendasi")
        statuses = ["Direkomendasikan", "Dipertimbangkan", "Belum Direkomendasikan"]
        totals = {s: sum(r["Rekomendasi"] == s for r in ranked) for s in statuses}
        st.dataframe([{"Hasil Sistem": label, "Jumlah": totals[label]} for label in statuses], hide_index=True, use_container_width=True)
        st.subheader("Keputusan Akhir")
        decisions = ["Lolos", "Cadangan", "Tidak Lolos", "Menunggu"]
        st.dataframe([{"Keputusan Pewawancara": x, "Jumlah": sum((I.get(cid, {}).get("keputusan") or "Menunggu") == x for cid in C)} for x in decisions], hide_index=True, use_container_width=True)
    with right:
        st.subheader("Kandidat Teratas")
        if ranked:
            st.dataframe([{"Ranking": r["Ranking"], "Nama": r["Nama Kandidat"], "Skor": r["Skor"]} for r in ranked[:8]], hide_index=True, use_container_width=True)
        else:
            st.info("Belum ada kandidat dengan tujuh aspek penilaian yang lengkap.")
        st.subheader("Alur Penggunaan")
        st.markdown("**1. Data Kandidat** → simpan profil peserta.\n\n**2. Penilaian Interview** → isi tujuh nilai dan catatan.\n\n**3. Summary & Ranking** → bandingkan hasil.\n\n**4. Download Backup Excel** → simpan seluruh data.")

elif menu == "Data Kandidat":
    section("PROFIL PESERTA", "Data Kandidat", "Form menggantikan 25 kolom pada sheet Data Kandidat. Satu peserta menggunakan satu ID MG-xxx.")
    options = ["➕ Tambah Kandidat Baru"] + sorted(C.keys())
    choice = st.selectbox("Pilih kandidat untuk dilihat/diedit", options, format_func=lambda x: f"{x} — {C[x].get('nama','')}" if x in C else x)
    existing = choice in C
    cid = choice if existing else candidate_id_next(C)
    initial = C[choice].copy() if existing else {}
    st.caption(f"ID Kandidat: **{cid}** · Kapasitas template Excel: {len(C)}/{MAX_KANDIDAT}")
    with st.form(f"form_candidate_{cid}"):
        t1, t2, t3, t4 = st.tabs(["Identitas & Kontak", "Pendidikan", "Pengalaman & Kompetensi", "Preferensi Magang & Berkas"])
        result = {"id": cid}
        with t1:
            a, b = st.columns(2)
            with a:
                result["nama"] = st.text_input("Nama Lengkap *", value=initial.get("nama", ""))
                result["telepon"] = st.text_input("No. HP / WA", value=initial.get("telepon", ""))
                result["alamat"] = st.text_area("Alamat Domisili", value=initial.get("alamat", ""), height=80)
            with b:
                result["email"] = st.text_input("Email", value=initial.get("email", ""))
                result["kota"] = st.text_input("Kota / Kabupaten", value=initial.get("kota", ""))
        with t2:
            a, b = st.columns(2)
            with a:
                result["kampus"] = st.text_input("Asal Perguruan Tinggi", value=initial.get("kampus", ""))
                result["prodi"] = st.text_input("Program Studi / Jurusan", value=initial.get("prodi", ""))
                result["ipk"] = st.number_input("IPK (0,00–4,00)", min_value=0.0, max_value=4.0, value=min(max(_float(initial.get("ipk")), 0.0), 4.0), step=0.01, format="%.2f")
            with b:
                result["jenjang"] = _select("Jenjang", initial.get("jenjang", ""), ["", "D3", "D4", "S1", "S2", "Lainnya"])
                result["semester"] = st.number_input("Semester", min_value=0, max_value=16, value=min(max(_int(initial.get("semester")), 0), 16), step=1)
                result["status_kuliah"] = _select("Status Perkuliahan", initial.get("status_kuliah", ""), ["", "Aktif", "Cuti", "Lulus", "Lainnya"])
        with t3:
            a, b = st.columns(2)
            with a:
                result["pengalaman_kerja"] = st.text_area("Pengalaman Magang / Kerja", value=initial.get("pengalaman_kerja", ""), height=110)
                result["pengalaman_organisasi"] = st.text_area("Pengalaman Organisasi", value=initial.get("pengalaman_organisasi", ""), height=110)
                result["portofolio"] = st.text_area("Proyek / Portofolio", value=initial.get("portofolio", ""), height=110)
            with b:
                result["keahlian"] = st.text_area("Keahlian Utama", value=initial.get("keahlian", ""), height=110)
                result["tools"] = st.text_area("Tools / Software", value=initial.get("tools", ""), height=110)
                result["sertifikasi"] = st.text_area("Sertifikasi / Pelatihan", value=initial.get("sertifikasi", ""), height=110)
            result["link"] = st.text_input("Tautan LinkedIn / Portofolio", value=initial.get("link", ""))
        with t4:
            a, b = st.columns(2)
            with a:
                result["bidang"] = st.text_input("Posisi / Bidang Magang", value=initial.get("bidang", ""))
                result["wfo"] = _select("Bersedia WFO", initial.get("wfo", ""), ["", "Ya", "Tidak", "Hybrid", "Negosiasi"])
                result["durasi"] = st.number_input("Durasi (bulan)", min_value=0, max_value=24, step=1, value=min(max(_int(initial.get("durasi")), 0), 24))
            with b:
                result["mulai"] = st.date_input("Dapat Mulai Tanggal", value=_date(initial.get("mulai")), format="DD/MM/YYYY")
                result["status_cv"] = _select("Status CV / Berkas", initial.get("status_cv", ""), ["", "Lengkap", "Belum Lengkap", "Proses Verifikasi"])
            result["motivasi"] = st.text_area("Motivasi / Harapan Magang", value=initial.get("motivasi", ""), height=105)
        submitted = st.form_submit_button("💾 Simpan Data Kandidat", type="primary", use_container_width=True)
    if submitted:
        if not result["nama"].strip():
            st.error("Nama lengkap wajib diisi.")
        elif not existing and len(C) >= MAX_KANDIDAT:
            st.error("Kapasitas 150 kandidat telah tercapai pada template Excel ini.")
        else:
            result["nama"] = result["nama"].strip()
            result["mulai"] = result["mulai"].isoformat() if result["mulai"] else ""
            C[cid] = result
            st.success(f"Profil {result['nama']} ({cid}) berhasil disimpan dalam sesi. Download backup Excel untuk menyimpannya permanen.")
    st.divider()
    st.subheader("Daftar Kandidat")
    search = st.text_input("Cari nama, kampus, atau posisi", placeholder="Masukkan kata kunci...")
    table = []
    for key, val in sorted(C.items()):
        searchable = " ".join(str(val.get(field, "")) for field in ["nama", "kampus", "prodi", "bidang"]).lower()
        if search and search.lower() not in searchable:
            continue
        score, status = calculate(I.get(key), weights, high, mid)
        table.append({"ID": key, "Nama": val.get("nama", ""), "Perguruan Tinggi": val.get("kampus", ""), "Program Studi": val.get("prodi", ""), "IPK": val.get("ipk", ""), "Posisi": val.get("bidang", ""), "Status Interview": status})
    if table:
        _shortlist_table(table)
    else:
        st.info("Belum ada data kandidat. Silakan tambah kandidat melalui form di atas.")

elif menu == "Penilaian Interview":
    section("WAWANCARA & ASESMEN", "Penilaian Interview", "Pilih kandidat yang sudah didaftarkan, beri nilai 1–5 pada setiap aspek dan tuliskan hasil wawancara.")
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
            I[cid] = saved
            sc, recom = calculate(saved, weights, high, mid)
            if sc is None:
                st.success("Catatan wawancara disimpan sebagai draft. Lengkapi tujuh aspek penilaian untuk menampilkan skor dan ranking.")
            else:
                st.success(f"Wawancara berhasil disimpan · Skor: {sc:.1f}/100 · {recom}.")
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
    section("HASIL SELEKSI", "Summary & Ranking", "Ranking dihitung berdasarkan bobot dari template Excel. Hanya kandidat dengan tujuh nilai lengkap yang masuk peringkat.")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Kandidat", len(C))
    c2.metric("Sudah Dinilai", len(ranked))
    c3.metric("Belum Lengkap", len(C) - len(ranked))
    c4.metric("Nilai Tertinggi", f"{ranked[0]['Skor']:.1f}" if ranked else "—")
    if not ranked:
        st.info("Belum ada ranking. Input data kandidat dan lengkapi tujuh aspek nilai di menu Penilaian Interview.")
    else:
        st.subheader("Peringkat Kandidat")
        f1, f2 = st.columns(2)
        with f1:
            cats = ["Semua", "Direkomendasikan", "Dipertimbangkan", "Belum Direkomendasikan"]
            reco_filter = st.selectbox("Filter rekomendasi", cats)
        with f2:
            positions = sorted({str(r["Posisi Magang"]) for r in ranked if str(r["Posisi Magang"]).strip()})
            pos_filter = st.selectbox("Filter posisi magang", ["Semua", *positions])
        filtered = [r for r in ranked if (reco_filter == "Semua" or r["Rekomendasi"] == reco_filter) and (pos_filter == "Semua" or r["Posisi Magang"] == pos_filter)]
        ordered_keys = ["Ranking", "ID", "Nama Kandidat", "Perguruan Tinggi", "Program Studi", "Posisi Magang", "IPK", "Skor", "Rekomendasi", "Keputusan Akhir"]
        st.dataframe([{k: r[k] for k in ordered_keys} for r in filtered], use_container_width=True, hide_index=True)
        st.caption("Peringkat tetap mengikuti seluruh kandidat yang telah dinilai, meskipun tabel sedang difilter. Nilai sama memperoleh ranking yang sama.")
        if filtered:
            st.subheader("Grafik Perbandingan Nilai")
            chart = [{"Kandidat": f"{r['Ranking']}. {r['Nama Kandidat'][:25]}", "Nilai": r["Skor"]} for r in filtered[:12]]
            st.bar_chart(chart, x="Kandidat", y="Nilai", horizontal=True)
            stream = StringIO()
            writer = csv.DictWriter(stream, fieldnames=ordered_keys)
            writer.writeheader()
            for r in filtered:
                writer.writerow({k: r[k] for k in ordered_keys})
            st.download_button("⬇️ Download Ranking (CSV)", stream.getvalue().encode("utf-8-sig"), file_name="ranking_peserta_magang.csv", mime="text/csv")
    st.divider()
    st.subheader("Progress Wawancara Semua Kandidat")
    progresses = []
    for cid, p in sorted(C.items()):
        interview = I.get(cid)
        sc, reco = calculate(interview, weights, high, mid)
        filled = sum(1 for key, _, _ in ASPEK if interview and interview.get(key) in range(1,6))
        progresses.append({"ID": cid, "Nama": p.get("nama", ""), "Aspek Dinilai": f"{filled}/7", "Status": "Selesai" if sc is not None else ("Draft" if interview else "Belum Interview"), "Skor": sc, "Keputusan": (interview or {}).get("keputusan", "—")})
    if progresses:
        st.dataframe(progresses, use_container_width=True, hide_index=True)

elif menu == "Panduan & Rubrik":
    section("STANDAR EVALUASI", "Panduan & Rubrik", "Materi ini diambil langsung dari sheet Panduan & Rubrik pada template Excel, sehingga aspek, bobot, dan ambang nilai tetap konsisten.")
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
