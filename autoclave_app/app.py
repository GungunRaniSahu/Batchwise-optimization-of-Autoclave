from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import pandas as pd
import os, json, shutil, math
from datetime import datetime
from openpyxl import load_workbook
from flask_login import (
    LoginManager, login_user, logout_user, login_required, current_user
)

from config import Config
from auth import (
    verify_user, get_user, create_user, user_exists, any_users_exist,
    verify_recovery_pin, reset_password,
)
from scheduler import start_scheduler

BASE_DIR = Config.BASE_DIR if hasattr(Config, "BASE_DIR") else os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = Config.EXCEL_FILE
BACKUP_DIR = Config.BACKUP_DIR
SETTINGS_FILE = Config.SETTINGS_FILE
OUTPUT_FILE = Config.OUTPUT_FILE
LOGO_FILE = Config.LOGO_FILE

os.makedirs(BACKUP_DIR, exist_ok=True)
os.makedirs(os.path.dirname(LOGO_FILE), exist_ok=True)

app = Flask(__name__)
app.config.from_object(Config)
app.secret_key = Config.SECRET_KEY

# ---------------------------------------------------------------------------
# Login manager
# ---------------------------------------------------------------------------
login_manager = LoginManager(app)
login_manager.login_view = "login"
login_manager.login_message = "Please sign in to continue."
login_manager.login_message_category = "warning"


@login_manager.user_loader
def load_user(username):
    return get_user(Config.USERS_FILE, username)


# ---------------------------------------------------------------------------
# Settings persistence (unchanged business logic from the original app)
# ---------------------------------------------------------------------------
DEFAULT_SETTINGS = {
    "autoclave_length_mm": 6000.0,
    "autoclave_width_mm": 2500.0,
    "total_ports": 24,
    "big_ports": 4,
    "small_ports": 2,
    "max_big_parts": 2,
    # Change this after confirming your engineering definition of BIG.
    "big_area_threshold_in2": 10000.0
}


def load_settings():
    if not os.path.exists(SETTINGS_FILE):
        save_settings(DEFAULT_SETTINGS)
        return DEFAULT_SETTINGS.copy()
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        out = DEFAULT_SETTINGS.copy()
        out.update(data)
        return out
    except Exception:
        return DEFAULT_SETTINGS.copy()


def save_settings(settings):
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(settings, f, indent=2)


def backup_excel():
    if os.path.exists(EXCEL_FILE):
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        shutil.copy2(EXCEL_FILE, os.path.join(BACKUP_DIR, f"parts_backup_{stamp}.xlsx"))


def normalize_columns(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    aliases = {
        "Availability": "Availability",
        "Availability ": "Availability",
        "TOOL NUMBER": "TOOL NUMBER",
        "TOOL NUMBER ": "TOOL NUMBER",
        "WIDTH(INCH)": "WIDTH(INCH)",
        "WIDTH (INCH)": "WIDTH(INCH)",
        "LENGTH (INCH)": "LENGTH (INCH)",
        "SURFACE AREA(INCH SQ)": "SURFACE AREA(INCH SQ)"
    }
    df.rename(columns=aliases, inplace=True)
    required = ["Part Number", "L (INCHES)", "W (INCHES)", "H (INCHES)",
                "Availability", "TOOL NUMBER", "LENGTH (INCH)",
                "WIDTH(INCH)", "SURFACE AREA(INCH SQ)"]
    for col in required:
        if col not in df.columns:
            df[col] = ""
    return df


def load_parts():
    if not os.path.exists(EXCEL_FILE):
        raise FileNotFoundError("parts.xlsx was not found in the application folder.")
    df = pd.read_excel(EXCEL_FILE, sheet_name=0)
    return normalize_columns(df)


def save_parts(df):
    backup_excel()
    df.to_excel(EXCEL_FILE, index=False)


def num(v):
    try:
        return float(v)
    except Exception:
        return 0.0


def inch_to_mm(v):
    return num(v) * 25.4


def get_part_type(row, settings):
    area = num(row.get("SURFACE AREA(INCH SQ)", 0))
    return "BIG" if area >= settings["big_area_threshold_in2"] else "SMALL"


def rect_fits(rect_l, rect_w, placements, L, W):
    """Simple 2D rectangle packing with 90-degree rotation.
    Returns (fits, x, y, placed_l, placed_w).
    """
    candidates = {(0.0, 0.0)}
    for x, y, l, w, _ in placements:
        candidates.add((x + l, y))
        candidates.add((x, y + w))
        candidates.add((x + l, y + w))

    orientations = [(rect_l, rect_w)]
    if abs(rect_l - rect_w) > 1e-9:
        orientations.append((rect_w, rect_l))

    def overlaps(x, y, l, w, p):
        px, py, pl, pw, _ = p
        return not (x + l <= px or px + pl <= x or y + w <= py or py + pw <= y)

    for x, y in sorted(candidates, key=lambda q: (q[1], q[0])):
        for l, w in orientations:
            if x + l > L or y + w > W:
                continue
            if any(overlaps(x, y, l, w, p) for p in placements):
                continue
            return True, x, y, l, w
    return False, None, None, None, None


def optimize_batches(df, settings):
    df = df.copy()
    df["Availability"] = df["Availability"].astype(str).str.strip().str.upper()
    available = df[df["Availability"] == "YES"].copy()

    if available.empty:
        return [], pd.DataFrame(), []

    available["_tool_L_mm"] = available["LENGTH (INCH)"].apply(inch_to_mm)
    available["_tool_W_mm"] = available["WIDTH(INCH)"].apply(inch_to_mm)
    available["_area_in2"] = available["SURFACE AREA(INCH SQ)"].apply(num)
    available["_type"] = available.apply(lambda r: get_part_type(r, settings), axis=1)
    available["_ports"] = available["_type"].map({"BIG": settings["big_ports"], "SMALL": settings["small_ports"]})

    invalid = []
    valid_rows = []
    for idx, r in available.iterrows():
        row = r.to_dict()
        if not str(row["TOOL NUMBER"]).strip() or row["_tool_L_mm"] <= 0 or row["_tool_W_mm"] <= 0:
            invalid.append({"Part Number": row["Part Number"], "Reason": "Missing/invalid tool dimensions"})
            continue
        if row["_tool_L_mm"] > settings["autoclave_length_mm"] and row["_tool_W_mm"] > settings["autoclave_width_mm"]:
            invalid.append({"Part Number": row["Part Number"], "Reason": "Tool exceeds autoclave in both orientations"})
            continue
        valid_rows.append(row)

    remaining = valid_rows[:]
    remaining.sort(key=lambda r: (0 if r["_type"] == "BIG" else 1, -r["_area_in2"]))

    batches = []
    batch_no = 1

    while remaining:
        bigs = [r for r in remaining if r["_type"] == "BIG"]
        if bigs:
            seed = bigs[0]
        else:
            seed = remaining[0]

        selected = [seed]
        remaining.remove(seed)
        used_ports = int(seed["_ports"])
        placements = []
        ok, x, y, pl, pw = rect_fits(seed["_tool_L_mm"], seed["_tool_W_mm"], placements,
                                     settings["autoclave_length_mm"], settings["autoclave_width_mm"])
        if not ok:
            invalid.append({"Part Number": seed["Part Number"], "Reason": "Tool could not be placed"})
            continue
        placements.append((x, y, pl, pw, seed["Part Number"]))
        big_count = 1 if seed["_type"] == "BIG" else 0

        while remaining:
            candidates = []
            for r in remaining:
                if used_ports + int(r["_ports"]) > settings["total_ports"]:
                    continue
                if r["_type"] == "BIG" and big_count >= settings["max_big_parts"]:
                    continue
                ok, cx, cy, cl, cw = rect_fits(r["_tool_L_mm"], r["_tool_W_mm"], placements,
                                               settings["autoclave_length_mm"], settings["autoclave_width_mm"])
                if not ok:
                    continue
                area = r["_area_in2"]
                score = (1000000 if r["_type"] == "BIG" else 0) + area * 10 + int(r["_ports"]) * 100
                candidates.append((score, r, cx, cy, cl, cw))

            if not candidates:
                break

            candidates.sort(key=lambda z: z[0], reverse=True)
            _, chosen, cx, cy, cl, cw = candidates[0]
            selected.append(chosen)
            remaining.remove(chosen)
            used_ports += int(chosen["_ports"])
            if chosen["_type"] == "BIG":
                big_count += 1
            placements.append((cx, cy, cl, cw, chosen["Part Number"]))

        batches.append({
            "batch_number": f"BATCH-{batch_no:03d}",
            "parts": selected,
            "used_ports": used_ports,
            "port_utilization": used_ports / settings["total_ports"] * 100,
            "placements": placements,
            "big_count": big_count,
            "tool_area_in2": sum(r["_area_in2"] for r in selected),
        })
        batch_no += 1

    report_rows = []
    summary_rows = []
    for b in batches:
        for r in b["parts"]:
            report_rows.append({
                "Batch Number": b["batch_number"],
                "Part Number": r["Part Number"],
                "Part Type": r["_type"],
                "Availability": r["Availability"],
                "Tool Number": r["TOOL NUMBER"],
                "Part L (in)": r["L (INCHES)"],
                "Part W (in)": r["W (INCHES)"],
                "Part H (in)": r["H (INCHES)"],
                "Tool L (in)": r["LENGTH (INCH)"],
                "Tool W (in)": r["WIDTH(INCH)"],
                "Tool Area (in²)": r["_area_in2"],
                "Ports": r["_ports"],
            })
        summary_rows.append({
            "Batch Number": b["batch_number"],
            "Parts": len(b["parts"]),
            "BIG Parts": b["big_count"],
            "SMALL Parts": len(b["parts"]) - b["big_count"],
            "Ports Used": b["used_ports"],
            "Port Capacity": settings["total_ports"],
            "Port Utilization %": round(b["port_utilization"], 2),
            "Tool Area (in²)": round(b["tool_area_in2"], 2),
            "Tool Area (m²)": round(b["tool_area_in2"] * 0.00064516, 3),
            "Layout Status": "PASS"
        })

    return batches, pd.DataFrame(report_rows), invalid


def run_and_save():
    settings = load_settings()
    df = load_parts()
    batches, report, invalid = optimize_batches(df, settings)
    summary = pd.DataFrame([
        {
            "Total Parts in Excel": len(df),
            "Available Parts": int((df["Availability"].astype(str).str.strip().str.upper() == "YES").sum()),
            "Batched Parts": sum(len(b["parts"]) for b in batches),
            "Total Batches": len(batches),
            "Autoclave Capacity (ports)": settings["total_ports"],
            "Autoclave L (mm)": settings["autoclave_length_mm"],
            "Autoclave W (mm)": settings["autoclave_width_mm"],
        }
    ])
    batch_summary = pd.DataFrame([{
        "Batch Number": b["batch_number"],
        "Parts": len(b["parts"]),
        "BIG Parts": b["big_count"],
        "SMALL Parts": len(b["parts"]) - b["big_count"],
        "Ports Used": b["used_ports"],
        "Port Capacity": settings["total_ports"],
        "Port Utilization %": round(b["port_utilization"], 2),
        "Tool Area (in²)": round(b["tool_area_in2"], 2),
        "Tool Area (m²)": round(b["tool_area_in2"] * 0.00064516, 3),
        "Layout Status": "PASS"
    } for b in batches])
    invalid_df = pd.DataFrame(invalid)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        summary.to_excel(writer, index=False, sheet_name="Run Summary")
        batch_summary.to_excel(writer, index=False, sheet_name="Batch Summary")
        report.to_excel(writer, index=False, sheet_name="Batch Details")
        invalid_df.to_excel(writer, index=False, sheet_name="Excluded Parts")

    return df, batches, invalid


# ---------------------------------------------------------------------------
# Auth routes
# ---------------------------------------------------------------------------
@app.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect(url_for("index"))
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        remember = bool(request.form.get("remember"))
        user = verify_user(Config.USERS_FILE, username, password)
        if user:
            login_user(user, remember=remember)
            flash(f"Welcome back, {username}.", "success")
            next_page = request.args.get("next")
            return redirect(next_page or url_for("index"))
        flash("Invalid username or password.", "error")
    return render_template("login.html", logo_exists=os.path.exists(LOGO_FILE))


@app.route("/logout")
@login_required
def logout():
    logout_user()
    flash("You have been signed out.", "success")
    return redirect(url_for("login"))


@app.route("/register", methods=["GET", "POST"])
def register():
    if current_user.is_authenticated:
        return redirect(url_for("index"))

    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        confirm = request.form.get("confirm_password", "")
        recovery_pin = request.form.get("recovery_pin", "").strip()

        if not username or not password:
            flash("Username and password are required.", "error")
            return render_template("register.html", logo_exists=os.path.exists(LOGO_FILE))
        if password != confirm:
            flash("Passwords do not match.", "error")
            return render_template("register.html", logo_exists=os.path.exists(LOGO_FILE))
        if len(password) < 8:
            flash("Password must be at least 8 characters.", "error")
            return render_template("register.html", logo_exists=os.path.exists(LOGO_FILE))
        if user_exists(Config.USERS_FILE, username):
            flash(f"Username '{username}' is already taken.", "error")
            return render_template("register.html", logo_exists=os.path.exists(LOGO_FILE))
        if not recovery_pin or len(recovery_pin) < 4:
            flash("Please set a recovery PIN of at least 4 digits — it's used for 'Forgot password'.", "error")
            return render_template("register.html", logo_exists=os.path.exists(LOGO_FILE))

        # The very first account created on a fresh install becomes admin;
        # everyone after that is an operator by default.
        role = "admin" if not any_users_exist(Config.USERS_FILE) else "operator"
        create_user(Config.USERS_FILE, username, password, role=role, recovery_pin=recovery_pin)
        flash(f"Account created for '{username}'. You can now sign in.", "success")
        return redirect(url_for("login"))

    return render_template("register.html", logo_exists=os.path.exists(LOGO_FILE))


@app.route("/forgot-password", methods=["GET", "POST"])
def forgot_password():
    if current_user.is_authenticated:
        return redirect(url_for("index"))

    if request.method == "POST":
        stage = request.form.get("stage", "verify")
        username = request.form.get("username", "").strip()

        if stage == "verify":
            pin = request.form.get("recovery_pin", "").strip()
            if verify_recovery_pin(Config.USERS_FILE, username, pin):
                return render_template("forgot_password.html", logo_exists=os.path.exists(LOGO_FILE),
                                        stage="reset", username=username, verified_pin=pin)
            flash("Username or recovery PIN is incorrect.", "error")
            return render_template("forgot_password.html", logo_exists=os.path.exists(LOGO_FILE), stage="verify")

        elif stage == "reset":
            pin = request.form.get("recovery_pin", "").strip()
            new_password = request.form.get("new_password", "")
            confirm = request.form.get("confirm_password", "")
            # Re-check the PIN so this step can't be reached by guessing the username alone.
            if not verify_recovery_pin(Config.USERS_FILE, username, pin):
                flash("Your session expired — please verify your recovery PIN again.", "error")
                return render_template("forgot_password.html", logo_exists=os.path.exists(LOGO_FILE), stage="verify")
            if len(new_password) < 8:
                flash("Password must be at least 8 characters.", "error")
                return render_template("forgot_password.html", logo_exists=os.path.exists(LOGO_FILE),
                                        stage="reset", username=username, verified_pin=pin)
            if new_password != confirm:
                flash("Passwords do not match.", "error")
                return render_template("forgot_password.html", logo_exists=os.path.exists(LOGO_FILE),
                                        stage="reset", username=username, verified_pin=pin)
            reset_password(Config.USERS_FILE, username, new_password)
            flash("Password updated. You can now sign in.", "success")
            return redirect(url_for("login"))

    return render_template("forgot_password.html", logo_exists=os.path.exists(LOGO_FILE), stage="verify")


# ---------------------------------------------------------------------------
# Application routes (all require login)
# ---------------------------------------------------------------------------
@app.route("/")
@login_required
def index():
    try:
        df = load_parts()
    except Exception as e:
        flash(str(e), "error")
        df = pd.DataFrame(columns=["Part Number", "L (INCHES)", "W (INCHES)", "H (INCHES)", "Availability",
                                    "TOOL NUMBER", "LENGTH (INCH)", "WIDTH(INCH)", "SURFACE AREA(INCH SQ)"])
    settings = load_settings()
    records = df.fillna("").to_dict(orient="records")
    return render_template(
        "index.html", parts=records, settings=settings,
        logo_exists=os.path.exists(LOGO_FILE), output_exists=os.path.exists(OUTPUT_FILE),
        retention_days=Config.DATA_RETENTION_DAYS,
    )


@app.post("/update_availability")
@login_required
def update_availability():
    part_number = request.form.get("part_number", "").strip()
    availability = request.form.get("availability", "NO").strip().upper()
    if availability not in ("YES", "NO"):
        flash("Availability must be YES or NO.", "error")
        return redirect(url_for("index"))
    df = load_parts()
    mask = df["Part Number"].astype(str).str.strip().str.upper() == part_number.upper()
    if not mask.any():
        flash(f"Part Number {part_number} not found.", "error")
        return redirect(url_for("index"))
    df.loc[mask, "Availability"] = availability
    save_parts(df)
    flash(f"{part_number} availability updated to {availability}.", "success")
    return redirect(url_for("index"))


@app.post("/update_all_availability")
@login_required
def update_all_availability():
    df = load_parts()
    changed = 0
    for key, value in request.form.items():
        if key.startswith("avail__"):
            part = key.split("__", 1)[1]
            mask = df["Part Number"].astype(str).str.strip() == part.strip()
            if mask.any() and value.upper() in ("YES", "NO"):
                df.loc[mask, "Availability"] = value.upper()
                changed += 1
    if changed:
        save_parts(df)
    flash(f"Updated availability for {changed} part(s).", "success")
    return redirect(url_for("index"))


@app.post("/add_part")
@login_required
def add_part():
    df = load_parts()
    data = {c: request.form.get(c, "").strip() for c in df.columns}
    part = data.get("Part Number", "")
    if not part:
        flash("Part Number is required.", "error")
        return redirect(url_for("index"))
    exists = df["Part Number"].astype(str).str.strip().str.upper().eq(part.upper()).any()
    if exists:
        flash(f"Part Number {part} already exists.", "error")
        return redirect(url_for("index"))
    df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)
    save_parts(df)
    flash(f"New part {part} added.", "success")
    return redirect(url_for("index"))


@app.post("/save_settings")
@login_required
def save_settings_route():
    settings = {
        "autoclave_length_mm": num(request.form.get("autoclave_length_mm")),
        "autoclave_width_mm": num(request.form.get("autoclave_width_mm")),
        "total_ports": int(num(request.form.get("total_ports"))),
        "big_ports": int(num(request.form.get("big_ports"))),
        "small_ports": int(num(request.form.get("small_ports"))),
        "max_big_parts": int(num(request.form.get("max_big_parts"))),
        "big_area_threshold_in2": num(request.form.get("big_area_threshold_in2"))
    }
    if min(settings.values()) <= 0:
        flash("All settings must be greater than zero.", "error")
        return redirect(url_for("index"))
    save_settings(settings)
    flash("Autoclave settings saved.", "success")
    return redirect(url_for("index"))


@app.post("/optimize")
@login_required
def optimize():
    try:
        _, batches, invalid = run_and_save()
        batched = sum(len(b["parts"]) for b in batches)
        if invalid:
            flash(f"Optimization complete: {len(batches)} batch(es), {batched} part(s) batched, "
                  f"{len(invalid)} excluded for data/fit issues.", "warning")
        else:
            flash(f"Optimization complete: {len(batches)} batch(es), {batched} part(s) batched.", "success")
    except Exception as e:
        flash(f"Optimization failed: {e}", "error")
    return redirect(url_for("index"))


@app.get("/download")
@login_required
def download():
    if not os.path.exists(OUTPUT_FILE):
        flash("Run optimization first.", "error")
        return redirect(url_for("index"))
    return send_file(OUTPUT_FILE, as_attachment=True, download_name="Optimized_Batches.xlsx")


if __name__ == "__main__":
    # Starts the 60-day-retention background cleanup job.
    start_scheduler()
    app.run(debug=True)