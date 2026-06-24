from flask import Flask, render_template, request, redirect, url_for, send_file, jsonify, session
import os, json, glob
from datetime import datetime
from werkzeug.utils import secure_filename
import pandas as pd

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "gofo_saturnv_2026_secure")
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024  # 32MB max upload

UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), "uploads")
REPORTS_FOLDER = os.path.join(os.path.dirname(__file__), "reports")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(REPORTS_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

CITY_CONFIG = {
    "ORD": {
        "name": "Chicago",
        "routes": {
            "ORD01-032": {"stop": 1.60, "extra": 0.25},
            "ORD01-007": {"stop": 1.70, "extra": 0.40},
        },
        "default_route": "ORD01-032",
        "bonos_fijos": {"susy": 800.0},
    },
    "SDF": {
        "name": "Louisville",
        "routes": {
            "SDF01-014": {"stop": 1.60, "extra": 0.50},
            "SDF01-027": {"stop": 1.70, "extra": 0.50},
        },
        "default_route": "SDF01-014",
        "bonos_fijos": {"yadrian": 800.0},
    },
    "MEM": {
        "name": "Memphis",
        "routes": {
            "MEM01-014": {"stop": 1.60, "extra": 0.40},
            "MEM01-020": {"stop": 1.79, "extra": 0.40},
            "MEM01-024": {"stop": 1.89, "extra": 0.40},
        },
        "default_route": "MEM01-014",
        "bonos_fijos": {"areli": 500.0},
    },
    "BTR": {
        "name": "Baton Rouge",
        "routes": {
            "BTR01-007": {"stop": 1.50, "extra": 0.25},
        },
        "default_route": "BTR01-007",
        "bonos_fijos": {"heather": 500.0},
    },
    "SHV": {
        "name": "Shreveport",
        "routes": {
            "SHV01-002": {"stop": 1.50, "extra": 0.25},
        },
        "default_route": "SHV01-002",
        "bonos_fijos": {"germai": 500.0},
    },
}

def detect_city(filepath):
    fname = os.path.basename(filepath).upper()
    for code in CITY_CONFIG:
        if fname.startswith(code + "-"):
            return code
    return None

def extract_drivers(filepath):
    city_code = detect_city(filepath)
    if not city_code:
        return None, []
    cfg = CITY_CONFIG[city_code]
    xl  = pd.ExcelFile(filepath)

    dsp = pd.read_excel(xl, sheet_name="DSP Summary", header=None).iloc[2]
    revenue_total = float(dsp[6])
    claims_total  = float(dsp[7])
    period        = str(dsp[4])

    ds_raw = pd.read_excel(xl, sheet_name="Driver Summary", header=None)
    driver_fee = {}
    for _, r in ds_raw.iloc[2:].iterrows():
        name = str(r[1]).strip()
        if not name or name == "nan": continue
        if any(x in name.lower() for x in ["payroll","gofo","profit","total"]): break
        driver_fee[name] = float(r[6]) if pd.notna(r[6]) else 0.0

    det_raw = pd.read_excel(xl, sheet_name="Details of Delivery Fees", header=None)
    det = det_raw.iloc[1:].copy()
    det.columns = det_raw.iloc[0].values
    det = det[det["Courier"].notna()].copy()
    det["Courier"] = det["Courier"].astype(str).str.strip()
    det["STOP Point Details"] = pd.to_numeric(det["STOP Point Details"], errors="coerce").fillna(1)
    det["Region/route"] = det["Region/route"].astype(str).str.strip()

    try:
        cl = pd.read_excel(xl, sheet_name="Claims Detail")
        cl["amt"] = pd.to_numeric(cl["The total expenses-exclusive of taxes"], errors="coerce").fillna(0)
        penalties = cl[cl["amt"] < 0].copy()
    except:
        penalties = pd.DataFrame(columns=["Courier","amt","Claim Type","Waybill No.","Remarks"])

    drivers = []
    for driver, gross_fee in driver_fee.items():
        d_det = det[det["Courier"] == driver]
        if len(d_det) == 0 and gross_fee >= 0:
            continue

        bono = 0.0
        for key, val in cfg["bonos_fijos"].items():
            if key.lower() in driver.lower():
                bono = val

        pen   = float(penalties[penalties["Courier"] == driver]["amt"].sum()) if len(penalties) > 0 else 0.0
        n_pen = int(len(penalties[penalties["Courier"] == driver])) if len(penalties) > 0 else 0

        rutas_data = []
        pago_stops_total = 0.0
        pago_extra_total = 0.0

        for ruta, grp in d_det.groupby("Region/route"):
            route_cfg    = cfg["routes"].get(ruta, cfg["routes"][cfg["default_route"]])
            tarifa_stop  = route_cfg["stop"]
            tarifa_extra = route_cfg["extra"]
            stops  = int((grp["STOP Point Details"] == 1).sum())
            extras = int((grp["STOP Point Details"] > 1).sum())
            ps = round(stops  * tarifa_stop,  2)
            pe = round(extras * tarifa_extra, 2)
            pago_stops_total += ps
            pago_extra_total += pe
            rutas_data.append({
                "ruta": ruta, "tarifa_stop": tarifa_stop, "tarifa_extra": tarifa_extra,
                "stops": stops, "extras": extras, "pago_stops": ps, "pago_extra": pe,
            })

        if len(d_det) == 0 and gross_fee < 0:
            # No-work driver with claim - company loss
            drivers.append({
                "driver": driver, "city": city_code, "city_name": cfg["name"],
                "rutas": "—", "rutas_data": [],
                "gross_fee": gross_fee, "pago_base": 0.0,
                "pago_stops": 0.0, "pago_extra": 0.0,
                "bono": 0.0, "penalizaciones": pen,
                "pago_total": 0.0, "revenue": gross_fee,
                "profit": gross_fee, "margin": 0.0,
                "n_penalizaciones": n_pen,
                "ajuste_bono": 0.0, "ajuste_cobro": 0.0,
                "ajuste_nota": "NO TRABAJÓ — pérdida compañía",
            })
            continue

        pago_base  = round(pago_stops_total + pago_extra_total, 2)
        pago_total = round(pago_base + bono + pen, 2)
        profit     = round(gross_fee - pago_total, 2)
        rutas_str  = " + ".join(sorted(set(r["ruta"] for r in rutas_data)))

        drivers.append({
            "driver": driver, "city": city_code, "city_name": cfg["name"],
            "rutas": rutas_str, "rutas_data": rutas_data,
            "gross_fee": gross_fee, "pago_base": pago_base,
            "pago_stops": round(pago_stops_total, 2),
            "pago_extra": round(pago_extra_total, 2),
            "bono": bono, "penalizaciones": pen,
            "pago_total": pago_total, "revenue": gross_fee,
            "profit": profit, "margin": round((profit/gross_fee*100) if gross_fee > 0 else 0, 1),
            "n_penalizaciones": n_pen,
            "ajuste_bono": 0.0, "ajuste_cobro": 0.0, "ajuste_nota": "",
        })

    return city_code, drivers, revenue_total, claims_total, period

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload():
    files = request.files.getlist("files")
    for f in glob.glob(os.path.join(UPLOAD_FOLDER, "*.xlsx")):
        os.remove(f)

    saved = []
    for f in files:
        if f.filename.endswith(".xlsx"):
            fname = secure_filename(f.filename)
            path  = os.path.join(UPLOAD_FOLDER, fname)
            f.save(path)
            saved.append(fname)

    if not saved:
        return redirect(url_for("index"))

    all_drivers = {}
    errors = []
    period = ""
    city_revenue = {}

    for fname in saved:
        path = os.path.join(UPLOAD_FOLDER, fname)
        try:
            city_code, drivers, revenue, claims, p = extract_drivers(path)
            if city_code:
                all_drivers[city_code] = drivers
                city_revenue[city_code] = {"revenue": revenue, "claims": claims}
                if not period:
                    period = p
        except Exception as e:
            errors.append(f"{fname}: {str(e)}")

    session["drivers"]      = all_drivers
    session["period"]       = period
    session["errors"]       = errors
    session["city_revenue"] = city_revenue
    return redirect(url_for("adjustments"))

@app.route("/adjustments", methods=["GET","POST"])
def adjustments():
    all_drivers = session.get("drivers", {})
    period      = session.get("period", "")
    if not all_drivers:
        return redirect(url_for("index"))

    if request.method == "POST":
        for city_code, drivers in all_drivers.items():
            for drv in drivers:
                key = f"{city_code}__{drv['driver']}"
                drv["ajuste_bono"]  = float(request.form.get(f"bono__{key}",  0) or 0)
                drv["ajuste_cobro"] = float(request.form.get(f"cobro__{key}", 0) or 0)
                drv["ajuste_nota"]  = request.form.get(f"nota__{key}", "")
                for rd in drv["rutas_data"]:
                    ruta_key = f"tarifa__{key}__{rd['ruta']}"
                    new_t = request.form.get(ruta_key, "")
                    if new_t:
                        try:
                            rd["tarifa_stop"] = float(new_t)
                            rd["pago_stops"]  = round(rd["stops"] * rd["tarifa_stop"], 2)
                        except: pass
                drv["pago_base"] = round(sum(r["pago_stops"]+r["pago_extra"] for r in drv["rutas_data"]), 2)
        session["drivers"] = all_drivers
        return redirect(url_for("preview"))

    return render_template("adjustments.html", all_drivers=all_drivers, period=period)

@app.route("/preview")
def preview():
    all_drivers   = session.get("drivers", {})
    period        = session.get("period", "")
    city_revenue  = session.get("city_revenue", {})
    if not all_drivers:
        return redirect(url_for("index"))

    city_summaries = {}
    for city_code, drivers in all_drivers.items():
        for drv in drivers:
            drv["pago_total"] = round(
                drv["pago_base"] + drv["bono"] +
                drv["ajuste_bono"] - abs(drv["ajuste_cobro"]) +
                drv["penalizaciones"], 2)
            drv["profit"] = round(drv["revenue"] - drv["pago_total"], 2)

        rev = city_revenue.get(city_code, {}).get("revenue", sum(d["revenue"] for d in drivers))
        pay = round(sum(d["pago_total"] for d in drivers), 2)
        city_summaries[city_code] = {
            "city_name": drivers[0]["city_name"] if drivers else "",
            "revenue":   round(rev, 2),
            "payroll":   pay,
            "profit":    round(rev - pay, 2),
            "margin":    round(((rev-pay)/rev*100) if rev > 0 else 0, 1),
            "drivers":   drivers,
        }

    grand_rev    = round(sum(c["revenue"] for c in city_summaries.values()), 2)
    grand_pay    = round(sum(c["payroll"] for c in city_summaries.values()), 2)
    grand_profit = round(grand_rev - grand_pay, 2)
    grand_margin = round((grand_profit/grand_rev*100) if grand_rev > 0 else 0, 1)

    return render_template("preview.html",
        city_summaries=city_summaries, period=period,
        grand_rev=grand_rev, grand_pay=grand_pay,
        grand_profit=grand_profit, grand_margin=grand_margin)

@app.route("/generate")
def generate():
    all_drivers  = session.get("drivers", {})
    period       = session.get("period", "Week")
    city_revenue = session.get("city_revenue", {})
    if not all_drivers:
        return redirect(url_for("index"))

    city_results = {}
    for city_code, drivers in all_drivers.items():
        for drv in drivers:
            drv["pago_total"] = round(
                drv["pago_base"] + drv["bono"] +
                drv["ajuste_bono"] - abs(drv["ajuste_cobro"]) +
                drv["penalizaciones"], 2)
            drv["profit"] = round(drv["revenue"] - drv["pago_total"], 2)
            drv["margin"] = round((drv["profit"]/drv["revenue"]*100) if drv["revenue"] > 0 else 0, 1)

        rev = city_revenue.get(city_code, {}).get("revenue", sum(d["revenue"] for d in drivers))
        pay = round(sum(d["pago_total"] for d in drivers), 2)
        city_results[city_code] = {
            "city_name":     drivers[0]["city_name"] if drivers else "",
            "period":        period,
            "revenue":       round(rev, 2),
            "claims":        city_revenue.get(city_code, {}).get("claims", 0),
            "total_payroll": pay,
            "profit":        round(rev - pay, 2),
            "margin":        round(((rev-pay)/rev*100) if rev > 0 else 0, 1),
            "drivers_list":  drivers,
        }

    from report_builder import build_report
    period_clean = period.replace(" ","_").replace("/","-")
    out = os.path.join(REPORTS_FOLDER, f"Reporte_GOFO_{period_clean}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")
    build_report(city_results, out)
    return send_file(out, as_attachment=True, download_name=os.path.basename(out))

@app.route("/finances")
def finances():
    return render_template("finances.html")

@app.route("/bank", methods=["GET","POST"])
def bank():
    analysis = None
    error    = None
    if request.method == "POST":
        f = request.files.get("csv_file")
        if f and f.filename.endswith(".csv"):
            path = os.path.join(UPLOAD_FOLDER, "bank_statement.csv")
            f.save(path)
            try:
                from bank_analyzer import analyze_csv
                analysis = analyze_csv(path)
                session["bank_analysis"] = analysis
            except Exception as e:
                error = str(e)
        else:
            error = "Por favor sube un archivo .csv de Wells Fargo"
    elif "bank_analysis" in session:
        analysis = session.get("bank_analysis")
    return render_template("bank.html", analysis=analysis, error=error)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5050))
    app.run(debug=False, host="0.0.0.0", port=port)
