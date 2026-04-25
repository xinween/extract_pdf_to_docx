import os
import subprocess
import uuid
from flask import Flask, request, send_file, render_template, flash
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = "agent_ready_2026"
app.config["UPLOAD_FOLDER"] = "/app/uploads"
app.config["OUTPUT_FOLDER"] = "/app/output"

os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
os.makedirs(app.config["OUTPUT_FOLDER"], exist_ok=True)

# ======================
# 你现在的 Skill 脚本路径
# ======================
EXTRACT_SCRIPT = "/app/skill/script/extract_from_pdfs.py"
GEN_DOCX_SCRIPT = "/app/skill/script/create_calibration_doc.py"

# ======================
# 未来 AGENT 接口（预留）
# ======================
USE_AGENT = False
try:
    from core.agent import PLCAgent
    agent = PLCAgent()
    USE_AGENT = True
except:
    pass

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "GET":
        return render_template("index.html")

    try:
        task_id = str(uuid.uuid4())
        task_dir = os.path.join(app.config["UPLOAD_FOLDER"], task_id)
        os.makedirs(task_dir, exist_ok=True)

        # 上传PDF
        pdf_files = request.files.getlist("pdfs")
        for f in pdf_files:
            if f.filename.endswith(".pdf"):
                f.save(os.path.join(task_dir, secure_filename(f.filename)))

        # 上传模板
        template = request.files["template"]
        tmpl_path = os.path.join(task_dir, "template.docx")
        template.save(tmpl_path)

        json_path = os.path.join(task_dir, "data.json")
        result_path = os.path.join(app.config["OUTPUT_FOLDER"], f"result_{task_id}.docx")

        # 1. 运行提取脚本
        p1 = subprocess.run(["python3", EXTRACT_SCRIPT, json_path, task_dir], capture_output=True, text=True)
        if p1.returncode !=0:
            flash(f"提取失败：{p1.stderr[:300]}")
            return render_template("index.html")

        # ======================
        # 【未来】这里接入Agent
        # ======================
        if USE_AGENT:
            agent.process_json(json_path)

        # 2. 生成文档
        p2 = subprocess.run(["python3", GEN_DOCX_SCRIPT, json_path, tmpl_path, result_path], capture_output=True, text=True)
        if p2.returncode !=0:
            flash(f"生成失败：{p2.stderr[:300]}")
            return render_template("index.html")

        return send_file(result_path, as_attachment=True)

    except Exception as e:
        flash(f"错误：{str(e)}")
        return render_template("index.html")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)