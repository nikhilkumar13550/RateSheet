"""
VHA Proposed Rates Generator – Flask Backend
Run:  python app.py
Then visit: http://localhost:5050
"""
import json
from flask import Flask, request, jsonify, send_file, send_from_directory
from io import BytesIO
import traceback

from processor import parse_and_clean, compute_report_data, compute_sold_rate_groups, DEFAULT_RATES, DEFAULT_ADJUSTMENTS
from generator import generate_excel, generate_sold_rate_sheet

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20 MB


@app.route("/")
def index():
    return send_from_directory(".", "index.html")


@app.route("/api/parse", methods=["POST"])
def api_parse():
    """Step 1 – parse & clean the uploaded L&V Excel. Returns summary JSON."""
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    f = request.files["file"]
    if not f.filename.lower().endswith((".xlsx", ".xls")):
        return jsonify({"error": "Please upload an Excel file (.xlsx or .xls)"}), 400

    try:
        file_bytes = f.read()
        parsed     = parse_and_clean(file_bytes)
        df         = parsed["df"]

        # Build a JSON-serialisable preview of the cleaned data
        preview = df.head(50).fillna("").astype(str).to_dict(orient="records")

        # Summary by benefit
        benefit_summary = (
            df.groupby("Benefit")
            .agg(Plans=("Plan", lambda x: ",".join(sorted(x.unique()))),
                 Total_Lives=("Lives", "sum"),
                 Total_Volumes=("Volumes", "sum"),
                 Rows=("Lives", "count"))
            .reset_index()
            .to_dict(orient="records")
        )

        return jsonify({
            "ok":               True,
            "client_name":      parsed["client_name"],
            "period_str":       parsed["period_str"],
            "contract_numbers": parsed["contract_numbers"],
            "total_rows":       len(df),
            "preview":          preview,
            "benefit_summary":  benefit_summary,
            "default_rates":    DEFAULT_RATES,
            "default_adjustments": DEFAULT_ADJUSTMENTS,
        })
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/api/generate", methods=["POST"])
def api_generate():
    """Step 2 – generate the Proposed Rates Excel output."""
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    f = request.files["file"]
    try:
        file_bytes  = f.read()
        form_data   = request.form

        # ── Parse rate overrides from form ─────────────────────────────────
        adjustments = {}
        raw_adj = form_data.get("adjustments")
        if raw_adj:
            try:
                adjustments = json.loads(raw_adj)
            except Exception:
                pass

        # ── Parse current rate overrides ───────────────────────────────────
        rates_override = {}
        raw_rates = form_data.get("rates")
        if raw_rates:
            try:
                rates_override = json.loads(raw_rates)
            except Exception:
                pass

        # ── Run pipeline ───────────────────────────────────────────────────
        parsed      = parse_and_clean(file_bytes)
        report_data = compute_report_data(
            parsed,
            rates=rates_override if rates_override else None,
            adjustments=adjustments if adjustments else None,
        )
        excel_bytes = generate_excel(report_data)

        return send_file(
            BytesIO(excel_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="VHA_Proposed_Rates.xlsx",
        )
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/api/preview", methods=["POST"])
def api_preview():
    """Return computed benefit groups as JSON for the UI preview table."""
    if "file" not in request.files:
        return jsonify({"error": "No file"}), 400
    f = request.files["file"]
    try:
        file_bytes  = f.read()
        form_data   = request.form
        adjustments = {}
        raw_adj = form_data.get("adjustments")
        if raw_adj:
            try:
                adjustments = json.loads(raw_adj)
            except Exception:
                pass

        parsed      = parse_and_clean(file_bytes)
        report_data = compute_report_data(
            parsed,
            adjustments=adjustments if adjustments else None,
        )

        # Serialise (replace None with null-friendly values)
        def clean(obj):
            if isinstance(obj, dict):
                return {k: clean(v) for k, v in obj.items()}
            if isinstance(obj, list):
                return [clean(i) for i in obj]
            if obj is None:
                return None
            return obj

        return jsonify({"ok": True, "report": clean(report_data)})
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/api/generate-sold-rates", methods=["POST"])
def api_generate_sold_rates():
    """Generate the Sold Rate Sheet Excel (ManuConnect format)."""
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    f = request.files["file"]
    try:
        file_bytes  = f.read()
        adjustments = {}
        raw_adj = request.form.get("adjustments")
        if raw_adj:
            try:
                adjustments = json.loads(raw_adj)
            except Exception:
                pass

        parsed      = parse_and_clean(file_bytes)
        report_data = compute_report_data(parsed, adjustments=adjustments or None)
        sold_rows   = compute_sold_rate_groups(parsed, adjustments=adjustments or None)
        excel_bytes = generate_sold_rate_sheet(report_data, sold_rows)

        return send_file(
            BytesIO(excel_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="VHA_Sold_Rate_Sheet.xlsx",
        )
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    print("\n🌿 VHA Proposed Rates Generator")
    print("   Visit: http://localhost:5050\n")
    app.run(debug=True, port=5050)
