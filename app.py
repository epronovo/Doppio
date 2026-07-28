from flask import Flask, render_template, request, jsonify
from MpGd_Save_Updates import monitor_update_folder
from MpGd_Build_Template import export_templates_for_guide
import traceback

app = Flask(__name__)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/run-updates", methods=["POST"])
def run_updates():
    try:
        monitor_update_folder()
        return jsonify(status="success", message="Mapping updates processed.")
    except Exception as e:
        return jsonify(status="error", message=str(e)), 500

@app.route("/build-templates", methods=["POST"])
def build_templates():
    try:
        guide_table = request.json.get("guide_table")
        if not guide_table:
            return jsonify(status="error", message="Missing guide_table"), 400

        export_templates_for_guide(guide_table)

        return jsonify(
            status="success",
            message=f"Templates built for {guide_table}"
        )
    except Exception as e:
        return jsonify(
            status="error",
            message=str(e),
            trace=traceback.format_exc()
        ), 500

if __name__ == "__main__":
    app.run(debug=True)
