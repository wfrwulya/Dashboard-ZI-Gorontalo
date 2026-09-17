
import os, re, io, json, hmac, sqlite3, datetime as dt
from pathlib import Path

import pandas as pd
import openpyxl
import streamlit as st
import plotly.express as px

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
SAMPLE_DIR = APP_DIR / "sample_data"
DB_PATH = DATA_DIR / "zi_dashboard.db"
DATA_DIR.mkdir(exist_ok=True)

# Deployment: set ADMIN_PASSWORD in Streamlit Secrets.
# Local development may use the ADMIN_PASSWORD environment variable.
ADMIN_PASSWORD = st.secrets.get("ADMIN_PASSWORD", os.environ.get("ADMIN_PASSWORD", ""))
admin_password = ADMIN_PASSWORD

os.makedirs("data", exist_ok=True)

st.set_page_config(
    page_title="Monitoring ZI Kanwil BPN Gorontalo",
    page_icon="📊",
    layout="wide",
)


st.markdown("""
<style>
    .stApp { background-color: #f5f7fa; }
    section[data-testid="stSidebar"] { background-color: #ffffff; }
    .kpi-card {
        background: white; padding: 20px; border-radius: 14px;
        border: 1px solid #e6e9ef; box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        min-height: 120px;
    }
    .kpi-title { font-size: 14px; color: #6b7280; margin-bottom: 8px; }
    .kpi-value { font-size: 30px; font-weight: 700; color: #111827; }
    .section-title { font-size: 20px; font-weight: 700; margin: 20px 0 10px; }
    .hero-card {
        background: white; padding: 24px 28px; border-radius: 16px;
        border: 1px solid #e6e9ef; box-shadow: 0 2px 10px rgba(0,0,0,0.04);
        margin-bottom: 18px;
    }
    .status-ok { color:#166534; background:#dcfce7; padding:5px 10px;
        border-radius:999px; font-weight:700; display:inline-block; }
    .status-no { color:#991b1b; background:#fee2e2; padding:5px 10px;
        border-radius:999px; font-weight:700; display:inline-block; }
    .tl-chip { padding:4px 9px; border-radius:999px; font-weight:700; font-size:13px; }
    .tl-ok { background:#dcfce7; color:#166534; }
    .tl-no { background:#fee2e2; color:#991b1b; }
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

AREA_NAMES = [
    "MANAJEMEN PERUBAHAN",
    "PENATAAN TATALAKSANA",
    "PENATAAN SISTEM MANAJEMEN SDM APARATUR",
    "PENGUATAN AKUNTABILITAS",
    "PENGUATAN PENGAWASAN",
    "PENINGKATAN KUALITAS PELAYANAN PUBLIK",
]

def db():
    con = sqlite3.connect(DB_PATH)
    con.execute("""
        CREATE TABLE IF NOT EXISTS summary (
            satker TEXT PRIMARY KEY,
            total_rb REAL,
            pengungkit REAL,
            pengungkit_pct REAL,
            hasil REAL,
            hasil_pct REAL,
            min_area_pct REAL,
            status TEXT,
            updated_at TEXT,
            source_file TEXT,
            area_json TEXT,
            result_json TEXT
        )
    """)
    con.execute("""
        CREATE TABLE IF NOT EXISTS points (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            satker TEXT,
            section TEXT,
            component_no TEXT,
            component TEXT,
            sub_no TEXT,
            subcomponent TEXT,
            code TEXT,
            indicator TEXT,
            answer TEXT,
            score REAL,
            score_pct REAL,
            review_note TEXT,
            tl_r INTEGER,
            tl_s INTEGER,
            tl_t INTEGER,
            pic TEXT,
            link TEXT,
            source_row INTEGER
        )
    """)
    con.commit()
    return con

def numeric_score_cells(ws, row, minv=None, maxv=None):
    out = []
    for c in range(8, 14):
        v = ws.cell(row, c).value
        if isinstance(v, (int, float)):
            if (minv is None or v >= minv) and (maxv is None or v <= maxv):
                out.append(v)
    return out

def extract_summary(file_bytes, satker, source_file):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    if "Utama" not in wb.sheetnames:
        raise ValueError("Sheet 'Utama' tidak ditemukan.")
    ws = wb["Utama"]

    total_rb = pengungkit = pengungkit_pct = hasil = hasil_pct = None
    for r in range(1, ws.max_row + 1):
        b = ws.cell(r, 2).value
        if not isinstance(b, str):
            continue
        label = b.strip().upper()
        if label == "NILAI EVALUASI REFORMASI BIROKRASI":
            vals = numeric_score_cells(ws, r, minv=1.01)
            total_rb = vals[-1] if vals else None
        elif label == "TOTAL PENGUNGKIT":
            vals = numeric_score_cells(ws, r, minv=1.01)
            pengungkit = vals[-1] if vals else None
            pcts = numeric_score_cells(ws, r, minv=0, maxv=1)
            pengungkit_pct = pcts[-1] if pcts else None
        elif label == "TOTAL HASIL":
            vals = numeric_score_cells(ws, r, minv=1.01)
            hasil = vals[-1] if vals else None
            pcts = numeric_score_cells(ws, r, minv=0, maxv=1)
            hasil_pct = pcts[-1] if pcts else None

    if total_rb is None or pengungkit is None:
        raise ValueError(
            "Nilai ringkasan tidak terbaca. Pastikan file LKE sudah dibuka dan disimpan di Excel "
            "agar formula memiliki nilai hasil perhitungan."
        )

    areas = []
    for r in range(6, 12):
        nums = numeric_score_cells(ws, r, minv=1.01)
        pcts = numeric_score_cells(ws, r, minv=0, maxv=1)
        areas.append({
            "no": str(ws.cell(r, 4).value).strip(".") if ws.cell(r, 4).value else str(r - 5),
            "name": str(ws.cell(r, 5).value).strip() if ws.cell(r, 5).value else AREA_NAMES[r - 6],
            "weight": ws.cell(r, 8).value,
            "score": nums[-1] if nums else None,
            "pct": pcts[-1] if pcts else None,
        })

    results = []
    for r in (15, 18):
        nums = numeric_score_cells(ws, r, minv=1.01)
        pcts = numeric_score_cells(ws, r, minv=0, maxv=1)
        name = ws.cell(r, 4).value or ws.cell(r, 5).value
        results.append({
            "name": str(name).strip() if name else "",
            "weight": ws.cell(r, 8).value,
            "score": nums[-1] if nums else None,
            "pct": pcts[-1] if pcts else None,
        })

    min_area_pct = min(a["pct"] for a in areas if isinstance(a["pct"], (int, float)))
    status = (
        "MEMENUHI"
        if total_rb >= 75 and pengungkit >= 40 and min_area_pct >= 0.60
        else "BELUM MEMENUHI"
    )

    return {
        "satker": satker,
        "total_rb": total_rb,
        "pengungkit": pengungkit,
        "pengungkit_pct": pengungkit_pct,
        "hasil": hasil,
        "hasil_pct": hasil_pct,
        "min_area_pct": min_area_pct,
        "status": status,
        "areas": areas,
        "results": results,
        "source_file": source_file,
        "updated_at": dt.datetime.now().isoformat(timespec="seconds"),
    }

def extract_points(file_bytes, satker):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    if "Jawaban" not in wb.sheetnames:
        raise ValueError("Sheet 'Jawaban' tidak ditemukan.")
    ws = wb["Jawaban"]

    section = None
    area_no = area_name = None
    sub_no = sub_name = None
    result_no = result_name = None
    rows = []

    for r in range(2, 205):
        c, d, e, f, g = [ws.cell(r, col).value for col in range(3, 8)]
        if c in ("PENGUNGKIT", "HASIL"):
            section = c

        if section == "PENGUNGKIT":
            if d and isinstance(d, str) and re.match(r"^\d+\.$", d.strip()) and e:
                area_no = d.strip().rstrip(".")
                area_name = str(e).strip()
                sub_no = sub_name = None

            if e and isinstance(e, str) and re.match(r"^[ivx]+\.$", e.strip().lower()):
                sub_no = e.strip().rstrip(".")
                sub_name = str(f).strip() if f else ""

            j = ws.cell(r, 10).value
            n = ws.cell(r, 14).value
            if g is not None and j is not None and n is not None:
                fclean = str(f).strip().rstrip(".") if f is not None else ""
                helper = (
                    isinstance(g, str)
                    and g.strip().startswith("- ")
                    and fclean in ("", "-")
                )
                if not helper:
                    rows.append({
                        "satker": satker,
                        "section": "PENGUNGKIT",
                        "component_no": area_no,
                        "component": area_name,
                        "sub_no": sub_no or "",
                        "subcomponent": sub_name or "",
                        "code": f"A.{area_no}.{sub_no}.{fclean or '-'}",
                        "indicator": str(g).strip(),
                        "answer": str(n),
                        "score": ws.cell(r, 15).value,
                        "score_pct": ws.cell(r, 16).value,
                        "review_note": ws.cell(r, 17).value,
                        "tl_r": int(ws.cell(r, 18).value is True),
                        "tl_s": int(ws.cell(r, 19).value is True),
                        "tl_t": int(ws.cell(r, 20).value is True),
                        "pic": ws.cell(r, 22).value,
                        "link": ws.cell(r, 13).value,
                        "source_row": r,
                    })

        elif section == "HASIL":
            if (
                c
                and isinstance(c, str)
                and re.match(r"^[ivx]+\.$", c.strip().lower())
                and d
            ):
                result_no = c.strip().rstrip(".")
                result_name = str(d).strip()

            j = ws.cell(r, 10).value
            n = ws.cell(r, 14).value
            if e is not None and j is not None and n is not None:
                dclean = str(d).strip().rstrip(".") if d is not None else "-"
                rows.append({
                    "satker": satker,
                    "section": "HASIL",
                    "component_no": result_no,
                    "component": result_name,
                    "sub_no": "",
                    "subcomponent": "",
                    "code": f"B.{result_no}.{dclean}",
                    "indicator": str(e).strip(),
                    "answer": str(n),
                    "score": ws.cell(r, 15).value,
                    "score_pct": ws.cell(r, 16).value,
                    "review_note": ws.cell(r, 17).value,
                    "tl_r": int(ws.cell(r, 18).value is True),
                    "tl_s": int(ws.cell(r, 19).value is True),
                    "tl_t": int(ws.cell(r, 20).value is True),
                    "pic": ws.cell(r, 22).value,
                    "link": ws.cell(r, 13).value,
                    "source_row": r,
                })
    return rows

def load_data():
    con = db()
    summary = pd.read_sql_query("SELECT * FROM summary", con)
    points = pd.read_sql_query("SELECT * FROM points", con)
    con.close()
    return summary, points

def save_import(summary, points):
    con = db()
    satker = summary["satker"]
    con.execute("DELETE FROM summary WHERE satker = ?", (satker,))
    con.execute("DELETE FROM points WHERE satker = ?", (satker,))
    con.execute(
        """INSERT INTO summary VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
        (
            summary["satker"], summary["total_rb"], summary["pengungkit"],
            summary["pengungkit_pct"], summary["hasil"], summary["hasil_pct"],
            summary["min_area_pct"], summary["status"], summary["updated_at"],
            summary["source_file"], json.dumps(summary["areas"], ensure_ascii=False),
            json.dumps(summary["results"], ensure_ascii=False),
        ),
    )
    for p in points:
        con.execute(
            """INSERT INTO points
            (satker,section,component_no,component,sub_no,subcomponent,code,indicator,
             answer,score,score_pct,review_note,tl_r,tl_s,tl_t,pic,link,source_row)
             VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (
                p["satker"], p["section"], p["component_no"], p["component"],
                p["sub_no"], p["subcomponent"], p["code"], p["indicator"],
                p["answer"], p["score"], p["score_pct"], p["review_note"],
                p["tl_r"], p["tl_s"], p["tl_t"], p["pic"], p["link"], p["source_row"],
            ),
        )
    con.commit()
    con.close()


def check_icon(value):
    return "✅" if bool(value) else "❌"

def kpi_card(title, value):
    st.markdown(
        f'<div class="kpi-card"><div class="kpi-title">{title}</div>'
        f'<div class="kpi-value">{value}</div></div>',
        unsafe_allow_html=True,
    )

def status_badge(status):
    return (
        '<span class="status-ok">✓ MEMENUHI</span>'
        if status == "MEMENUHI"
        else '<span class="status-no">✕ BELUM MEMENUHI</span>'
    )

def tl_badge(value):
    return (
        '<span class="tl-chip tl-ok">✓</span>'
        if bool(value)
        else '<span class="tl-chip tl-no">✕</span>'
    )


def seed_sample_data():
    summary, points = load_data()
    if not summary.empty:
        return
    if not SAMPLE_DIR.exists():
        return

    mapping = [
        ("LKE Kantah Kab Gorontalo.xlsx", "Kantah Kabupaten Gorontalo"),
        ("LKE Kanwil BPN Gorontalo 2026. (1).xlsx", "Kanwil BPN Provinsi Gorontalo"),
    ]
    for filename, satker in mapping:
        path = SAMPLE_DIR / filename
        if path.exists():
            raw = path.read_bytes()
            s = extract_summary(raw, satker, filename)
            p = extract_points(raw, satker)
            save_import(s, p)

def admin_password():
    value = os.getenv("ZI_ADMIN_PASSWORD", "")
    try:
        value = st.secrets.get("ADMIN_PASSWORD", value)
    except Exception:
        pass
    return value

def is_admin():
    return st.session_state.get("admin_authenticated", False)

def status_counts(df):
    if df.empty:
        return 0, 0
    return int((df["status"] == "MEMENUHI").sum()), int((df["status"] != "MEMENUHI").sum())

def tl_summary(points, satker=None):
    df = points.copy()
    if satker:
        df = df[df["satker"] == satker]
    if df.empty:
        return pd.DataFrame()

    out = []
    for component, g in df.groupby("component", sort=False):
        total = len(g)
        r = int(g["tl_r"].sum())
        s = int(g["tl_s"].sum())
        t = int(g["tl_t"].sum())
        pct = (r + s + t) / (total * 3) if total else 0
        out.append({
            "Komponen": component,
            "Total Poin": total,
            "Catatan TPI": r,
            "Eviden 2024-2025": s,
            "Penamaan Eviden": t,
            "Checklist": r + s + t,
            "TL %": pct,
        })
    return pd.DataFrame(out)

def fmt_pct(v):
    return f"{v:.1%}" if isinstance(v, (int, float)) else "-"

seed_sample_data()
summary, points = load_data()

st.sidebar.title("Monitoring ZI")
page = st.sidebar.radio(
    "Menu",
    ["Dashboard", "Progress Satker", "Monitoring Perbaikan", "Admin"],
)
st.sidebar.caption(
    "Viewer dapat mengakses Dashboard, Progress, dan Monitoring Perbaikan. "
    "Fitur upload dikunci untuk Admin."
)

if summary.empty:
    st.warning("Belum ada data LKE. Admin perlu mengunggah file LKE terlebih dahulu.")
    if page != "Admin":
        st.stop()

if page == "Dashboard":
    st.markdown(
        '<div class="hero-card"><div style="font-size:14px;color:#6b7280;">'
        'MONITORING PEMBANGUNAN ZI</div><div style="font-size:30px;font-weight:800;'
        'color:#111827;">Dashboard Kinerja ZI</div><div style="font-size:15px;'
        'color:#6b7280;margin-top:6px;">Kantor Wilayah BPN Provinsi Gorontalo</div></div>',
        unsafe_allow_html=True,
    )

    n = len(summary)
    memenuhi, belum = status_counts(summary)
    avg_rb = summary["total_rb"].mean() if n else 0

    satker_tl = []
    for satker in summary["satker"]:
        p = points[points["satker"] == satker]
        total_checks = len(p) * 3
        pct = p[["tl_r", "tl_s", "tl_t"]].sum().sum() / total_checks if total_checks else 0
        satker_tl.append({"Satker": satker, "TL %": pct})
    avg_tl = pd.DataFrame(satker_tl)["TL %"].mean() if satker_tl else 0

    c1, c2, c3, c4 = st.columns(4)
    with c1: kpi_card("Jumlah Satker Diusulkan", n)
    with c2: kpi_card("Memenuhi Syarat WBK", f"{memenuhi} / {n}")
    with c3: kpi_card("Rata-rata Nilai RB", f"{avg_rb:.2f}")
    with c4: kpi_card("Rata-rata Tindak Lanjut", f"{avg_tl:.1%}")

    st.markdown('<div class="section-title">Status WBK & Progress Tindak Lanjut</div>',
                unsafe_allow_html=True)
    left, right = st.columns([1, 2])
    with left:
        status_df = pd.DataFrame({
            "Status": ["Memenuhi", "Belum Memenuhi"],
            "Jumlah": [memenuhi, belum],
        })
        fig_status = px.pie(status_df, names="Status", values="Jumlah", hole=0.65)
        fig_status.update_traces(
            textinfo="label+percent",
            hovertemplate="<b>%{label}</b><br>%{value} Satker<extra></extra>",
        )
        fig_status.update_layout(
            showlegend=True, height=320, margin=dict(l=10, r=10, t=30, b=10)
        )
        st.plotly_chart(fig_status, use_container_width=True)
    with right:
        tl_df = pd.DataFrame(satker_tl).sort_values("TL %") if satker_tl else pd.DataFrame()
        if not tl_df.empty:
            fig_tl = px.bar(tl_df, x="TL %", y="Satker", orientation="h",
                            text="TL %", range_x=[0, 1])
            fig_tl.update_traces(texttemplate="%{text:.1%}", textposition="outside")
            fig_tl.update_layout(xaxis_tickformat=".0%", height=320,
                                 margin=dict(l=10, r=30, t=30, b=10))
            st.plotly_chart(fig_tl, use_container_width=True)

    st.markdown('<div class="section-title">Rata-rata Capaian 6 Area Pengungkit</div>',
                unsafe_allow_html=True)
    area_rows = []
    for _, row in summary.iterrows():
        for a in json.loads(row["area_json"]):
            if isinstance(a.get("pct"), (int, float)):
                area_rows.append({"Satker": row["satker"], "Area": a["name"], "Capaian": a["pct"]})
    area_df = pd.DataFrame(area_rows)
    if not area_df.empty:
        avg_area = area_df.groupby("Area", as_index=False)["Capaian"].mean().sort_values("Capaian")
        fig_area = px.bar(avg_area, x="Capaian", y="Area", orientation="h",
                          text="Capaian", range_x=[0, 1])
        fig_area.update_traces(texttemplate="%{text:.1%}", textposition="outside",
                               hovertemplate="<b>%{y}</b><br>Capaian: %{x:.1%}<extra></extra>")
        fig_area.add_vline(x=0.60, line_dash="dash", annotation_text="Ambang 60%",
                           annotation_position="top")
        fig_area.update_layout(xaxis_tickformat=".0%", height=420,
                               margin=dict(l=10, r=40, t=30, b=10))
        st.plotly_chart(fig_area, use_container_width=True)

    st.markdown('<div class="section-title">Peringkat Nilai Satker</div>', unsafe_allow_html=True)
    rank = summary.sort_values("total_rb", ascending=False).copy()
    rank.insert(0, "Peringkat", range(1, len(rank) + 1))
    rank["Nilai RB"] = rank["total_rb"].round(2)
    rank["Pengungkit"] = rank["pengungkit"].round(2)
    rank["Min Area"] = rank["min_area_pct"].map(fmt_pct)
    rank["Status"] = rank["status"].map({"MEMENUHI": "Memenuhi", "BELUM MEMENUHI": "Belum Memenuhi"})
    rank["TL %"] = [next(x["TL %"] for x in satker_tl if x["Satker"] == s) for s in rank["satker"]]
    rank["TL %"] = rank["TL %"].map(fmt_pct)
    st.dataframe(
        rank[["Peringkat","satker","Nilai RB","Pengungkit","Min Area","TL %","Status"]]
        .rename(columns={"satker":"Satker"}),
        use_container_width=True, hide_index=True
    )

    st.markdown('<div class="section-title">Satker Belum Memenuhi</div>', unsafe_allow_html=True)
    not_ok = summary[summary["status"] != "MEMENUHI"].copy()
    if not not_ok.empty:
        rows = []
        for _, r in not_ok.iterrows():
            reasons = []
            if r["total_rb"] < 75: reasons.append("Nilai RB < 75")
            if r["pengungkit"] < 40: reasons.append("Pengungkit < 40")
            if r["min_area_pct"] < 0.60: reasons.append("Ada area < 60%")
            rows.append({"Satker":r["satker"], "Nilai RB":round(r["total_rb"],2),
                         "Pengungkit":round(r["pengungkit"],2),
                         "Min Area":fmt_pct(r["min_area_pct"]),
                         "Kendala":", ".join(reasons)})
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    else:
        st.success("Seluruh satker memenuhi kriteria status saat ini.")

    st.markdown('<div class="section-title">Kriteria Status Memenuhi Syarat WBK</div>',
                unsafe_allow_html=True)
    st.dataframe(pd.DataFrame({
        "Kriteria":["Nilai RB","Nilai Pengungkit","Nilai minimal setiap area pengungkit"],
        "Batas":["≥ 75","≥ 40","≥ 60%"],
    }), use_container_width=True, hide_index=True)

elif page == "Progress Satker":
    st.markdown(
        '<div class="hero-card"><div style="font-size:14px;color:#6b7280;">DRILL-DOWN</div>'
        '<div style="font-size:28px;font-weight:800;color:#111827;">Progress Pembangunan ZI Satker</div>'
        '<div style="font-size:15px;color:#6b7280;margin-top:6px;">Pilih satu satker untuk melihat '
        'area, hasil, dan tindak lanjut level poin.</div></div>',
        unsafe_allow_html=True,
    )

    satkers = summary["satker"].tolist()
    st.markdown("### Pilih Satker")
    selected = st.selectbox("Satker", satkers, label_visibility="collapsed")
    srow = summary[summary["satker"] == selected].iloc[0]
    selected_points = points[points["satker"] == selected].copy()
    areas = json.loads(srow["area_json"])
    results = json.loads(srow["result_json"])

    st.markdown(
        f'<div class="hero-card"><div style="font-size:22px;font-weight:800;">🏢 {selected}</div>'
        f'<div style="margin-top:10px;">{status_badge(srow["status"])}</div></div>',
        unsafe_allow_html=True,
    )

    c1,c2,c3,c4 = st.columns(4)
    with c1: kpi_card("Nilai RB", f'{srow["total_rb"]:.2f}')
    with c2: kpi_card("Nilai Pengungkit", f'{srow["pengungkit"]:.2f} / 60')
    with c3: kpi_card("Nilai Hasil", f'{srow["hasil"]:.2f} / 40')
    with c4: kpi_card("Capaian Area Terendah", fmt_pct(srow["min_area_pct"]))

    st.markdown('<div class="section-title">1. Capaian 6 Area Pengungkit</div>', unsafe_allow_html=True)
    area_view = pd.DataFrame(areas)
    for _, row in area_view.iterrows():
        pct = float(row.get("pct") or 0)
        st.markdown(
            f'<div style="display:flex;justify-content:space-between;margin-top:8px;margin-bottom:4px;">'
            f'<b>{row.get("no","")} · {row.get("name","")}</b><b>{pct:.1%}</b></div>',
            unsafe_allow_html=True,
        )
        st.progress(min(max(pct,0),1))

    area_gap = area_view[pd.to_numeric(area_view["pct"], errors="coerce") < 0.60].copy()
    if not area_gap.empty:
        st.warning("Terdapat area pengungkit di bawah ambang 60%.")
        gap = area_gap[["no","name","pct"]].copy()
        gap.columns = ["No.","Area","Capaian"]
        gap["Capaian"] = gap["Capaian"].map(fmt_pct)
        st.dataframe(gap, use_container_width=True, hide_index=True)

    st.markdown('<div class="section-title">2. Komponen Hasil</div>', unsafe_allow_html=True)
    result_df = pd.DataFrame(results)[["name","weight","score","pct"]]
    result_df.columns = ["Komponen Hasil","Bobot","Nilai","% Capaian"]
    result_df["% Capaian"] = result_df["% Capaian"].map(fmt_pct)
    st.dataframe(result_df, use_container_width=True, hide_index=True)

    st.markdown('<div class="section-title">3. Status Tindak Lanjut</div>', unsafe_allow_html=True)
    tldf = tl_summary(selected_points, selected)
    if not tldf.empty:
        chart_df = tldf[tldf["Komponen"].isin([a["name"] for a in areas])].copy()
        if not chart_df.empty:
            fig = px.bar(chart_df.sort_values("TL %"), x="TL %", y="Komponen",
                         orientation="h", text="TL %", range_x=[0,1])
            fig.update_traces(texttemplate="%{text:.1%}", textposition="outside")
            fig.update_layout(xaxis_tickformat=".0%", height=380,
                              margin=dict(l=10,r=30,t=30,b=10))
            st.plotly_chart(fig, use_container_width=True)
        display = tldf.copy()
        display["TL %"] = display["TL %"].map(fmt_pct)
        st.dataframe(display, use_container_width=True, hide_index=True)
    else:
        st.info("Belum ada data tindak lanjut.")

    st.markdown('<div class="section-title">4. Detail Tindak Lanjut Level Poin LKE</div>',
                unsafe_allow_html=True)
    filter_status = st.radio("Tampilkan", ["Semua","Belum Lengkap","Sudah Lengkap"], horizontal=True)
    filter_tl = st.multiselect("Filter Kekurangan", [
        "Perbaikan Catatan TPI","Eviden 2024-2025","Penyesuaian Nama Eviden"
    ])
    filtered = selected_points.copy()
    if filter_status == "Belum Lengkap":
        filtered = filtered[(filtered["tl_r"]==0)|(filtered["tl_s"]==0)|(filtered["tl_t"]==0)]
    elif filter_status == "Sudah Lengkap":
        filtered = filtered[(filtered["tl_r"]==1)&(filtered["tl_s"]==1)&(filtered["tl_t"]==1)]
    if filter_tl:
        mask = pd.Series(False, index=filtered.index)
        if "Perbaikan Catatan TPI" in filter_tl: mask |= filtered["tl_r"]==0
        if "Eviden 2024-2025" in filter_tl: mask |= filtered["tl_s"]==0
        if "Penyesuaian Nama Eviden" in filter_tl: mask |= filtered["tl_t"]==0
        filtered = filtered[mask]

    st.caption(f"{len(filtered)} poin ditampilkan.")
    if filtered.empty:
        st.success("Tidak ada poin sesuai filter.")
    else:
        for component in list(dict.fromkeys(filtered["component"].tolist())):
            g = filtered[filtered["component"] == component].copy()
            pct = g[["tl_r","tl_s","tl_t"]].sum().sum()/(len(g)*3) if len(g) else 0
            with st.expander(f"{component} • TL {pct:.1%} • {len(g)} poin"):
                for _, row in g.iterrows():
                    st.markdown(
                        f'<div class="hero-card" style="padding:16px;margin-bottom:10px;">'
                        f'<div style="font-size:18px;font-weight:800;">{row["code"]}</div>'
                        f'<div style="margin-top:6px;color:#374151;">{row["indicator"]}</div>'
                        f'<div style="margin-top:12px;">Catatan TPI {tl_badge(row["tl_r"])} '
                        f'&nbsp;&nbsp; Eviden 2024-2025 {tl_badge(row["tl_s"])} '
                        f'&nbsp;&nbsp; Penamaan Eviden {tl_badge(row["tl_t"])}</div>'
                        f'<div style="margin-top:10px;color:#6b7280;"><b>PIC:</b> {row["pic"] or "-"}</div>'
                        f'</div>', unsafe_allow_html=True
                    )
                    if row["link"]:
                        st.link_button("🔗 Buka Eviden", str(row["link"]))

elif page == "Monitoring Perbaikan":
    st.markdown(
        '<div class="hero-card"><div style="font-size:14px;color:#6b7280;">CONTROL CENTER</div>'
        '<div style="font-size:28px;font-weight:800;color:#111827;">Monitoring Perbaikan</div>'
        '<div style="font-size:15px;color:#6b7280;margin-top:6px;">Seluruh poin yang masih '
        'memerlukan tindak lanjut dari semua satker.</div></div>',
        unsafe_allow_html=True,
    )
    all_points = points.copy()
    if all_points.empty:
        st.info("Belum ada data poin.")
        st.stop()

    all_points["Status TL"] = all_points.apply(
        lambda r: "Sudah Lengkap" if r["tl_r"] and r["tl_s"] and r["tl_t"] else "Belum Lengkap", axis=1
    )
    all_points["Jenis Kekurangan"] = all_points.apply(
        lambda r: ", ".join([
            x for x, flag in [
                ("Perbaikan Catatan TPI", not bool(r["tl_r"])),
                ("Eviden 2024-2025", not bool(r["tl_s"])),
                ("Penyesuaian Nama Eviden", not bool(r["tl_t"])),
            ] if flag
        ]) or "Tidak ada", axis=1
    )

    f1,f2,f3 = st.columns(3)
    with f1: satker_filter = st.multiselect("Satker", sorted(all_points["satker"].dropna().unique()))
    with f2: area_filter = st.multiselect("Area/Komponen", sorted(all_points["component"].dropna().unique()))
    with f3: status_filter = st.selectbox("Status", ["Belum Lengkap","Semua","Sudah Lengkap"])
    weakness_filter = st.multiselect("Jenis Kekurangan", [
        "Perbaikan Catatan TPI","Eviden 2024-2025","Penyesuaian Nama Eviden"
    ])
    pic_search = st.text_input("Cari PIC", placeholder="Ketik nama PIC...")

    monitor = all_points.copy()
    if satker_filter: monitor = monitor[monitor["satker"].isin(satker_filter)]
    if area_filter: monitor = monitor[monitor["component"].isin(area_filter)]
    if status_filter != "Semua": monitor = monitor[monitor["Status TL"] == status_filter]
    if weakness_filter:
        monitor = monitor[monitor["Jenis Kekurangan"].apply(lambda x: any(w in x for w in weakness_filter))]
    if pic_search.strip():
        monitor = monitor[monitor["pic"].fillna("").astype(str).str.contains(pic_search.strip(), case=False, na=False)]

    c1,c2,c3 = st.columns(3)
    with c1: kpi_card("Poin Ditampilkan", len(monitor))
    with c2: kpi_card("Poin Belum Lengkap", int((monitor["Status TL"]=="Belum Lengkap").sum()))
    with c3:
        pct_complete = (monitor["Status TL"]=="Sudah Lengkap").mean() if len(monitor) else 0
        kpi_card("Tingkat Kelengkapan", f"{pct_complete:.1%}")

    display = monitor[["satker","component","code","indicator","Jenis Kekurangan","pic","link"]].copy()
    display.columns = ["Satker","Area/Komponen","Kode","Poin LKE","Jenis Kekurangan","PIC","Link Eviden"]
    st.dataframe(display, use_container_width=True, hide_index=True,
                 column_config={"Link Eviden": st.column_config.LinkColumn("Link Eviden", display_text="Buka Eviden")})

elif page == "Admin":
    st.markdown(
        '<div class="hero-card"><div style="font-size:14px;color:#6b7280;">ADMINISTRATION</div>'
        '<div style="font-size:28px;font-weight:800;color:#111827;">Manajemen Data LKE</div>'
        '<div style="font-size:15px;color:#6b7280;margin-top:6px;">Upload, validasi, dan perbarui '
        'data satker pada dashboard.</div></div>', unsafe_allow_html=True
    )
    if not is_admin():
        password = st.text_input("Password Admin", type="password")
        if st.button("Login", type="primary"):
            expected = admin_password()
            if expected and hmac.compare_digest(password, expected):
                st.session_state["admin_authenticated"] = True
                st.rerun()
            else:
                st.error("Password tidak sesuai atau ADMIN_PASSWORD belum dikonfigurasi.")
        st.info("Viewer tidak memerlukan login. Hanya Admin yang dapat mengakses fitur upload.")
        st.stop()

    if st.button("Logout"):
        st.session_state["admin_authenticated"] = False
        st.rerun()

    st.success("Mode Admin aktif.")
    st.markdown('<div class="section-title">Upload / Update LKE</div>', unsafe_allow_html=True)
    uploaded = st.file_uploader(
        "Upload file LKE (.xlsx)", type=["xlsx"], accept_multiple_files=True,
        help="Upload file LKE per satker. File baru menambahkan satker; nama sama akan diperbarui."
    )
    if uploaded:
        entries = []
        for f in uploaded:
            inferred = re.sub(r"\.xlsx$", "", f.name, flags=re.I)
            inferred = re.sub(r"^LKE\s*", "", inferred, flags=re.I)
            if "Kanwil" in inferred:
                inferred = "Kanwil BPN Provinsi Gorontalo"
            elif "Kantah" in inferred:
                inferred = inferred.replace("Kab ", "Kabupaten ")
            name = st.text_input(f"Nama Satker — {f.name}", value=inferred, key=f"satker_{f.name}")
            entries.append((f, name.strip()))

        if st.button("Import / Update Dashboard", type="primary"):
            errors = []
            for f, satker in entries:
                try:
                    raw = f.getvalue()
                    s = extract_summary(raw, satker, f.name)
                    p = extract_points(raw, satker)
                    save_import(s, p)
                except Exception as e:
                    errors.append(f"{f.name}: {e}")
            if errors:
                for e in errors: st.error(e)
            else:
                st.success("Data berhasil diimpor. Satker baru ditambahkan dan satker dengan nama sama diperbarui.")
                st.rerun()

    current, _ = load_data()
    st.markdown('<div class="section-title">Data Satker Saat Ini</div>', unsafe_allow_html=True)
    if not current.empty:
        st.dataframe(
            current[["satker","total_rb","pengungkit","min_area_pct","status","updated_at","source_file"]]
            .rename(columns={"satker":"Satker","total_rb":"Nilai RB","pengungkit":"Pengungkit",
                             "min_area_pct":"Min Area","status":"Status","updated_at":"Update",
                             "source_file":"File Sumber"}),
            use_container_width=True, hide_index=True
        )
    st.warning(
        "Untuk deployment internet, gunakan password yang kuat dan simpan di secrets/environment variable. "
        "Untuk lingkungan Google Workspace internal, autentikasi akun Google/allowlist admin lebih disarankan "
        "daripada password bersama."
    )

