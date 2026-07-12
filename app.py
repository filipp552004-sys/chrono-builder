"""chrono-builder — микросервис сборки Excel.
  POST /build      → карта хронометража (норма времени)
  POST /build-frd  → баланс рабочего времени + Lean-анализ (ФРД / самонаблюдение)
Аутентификация: заголовок X-API-Key = env CHRONO_BUILDER_API_KEY.
"""
import os, io
from flask import Flask, request, send_file, jsonify, abort
from builder import build
from frd_builder import build_frd

app = Flask(__name__)
API_KEY = os.environ.get("CHRONO_BUILDER_API_KEY", "")
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

def _check_key():
    if not API_KEY or request.headers.get("X-API-Key", "") != API_KEY:
        abort(401)

def _xlsx_response(wb, name):
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return send_file(buf, mimetype=XLSX_MIME, as_attachment=True, download_name=name)

@app.get("/health")
def health():
    return jsonify(status="ok", service="chrono-builder")

@app.post("/build")
def build_endpoint():
    _check_key()
    data = request.get_json(force=True, silent=True)
    if not data or not isinstance(data, dict):
        return jsonify(error="Ожидается JSON-объект с данными хронометража"), 400
    if not data.get("elements"):
        return jsonify(error="Нет операций-составляющих (elements)"), 400
    try:
        return _xlsx_response(build(data), "norma.xlsx")
    except Exception as e:
        return jsonify(error=f"Ошибка сборки: {e}"), 500

@app.post("/build-frd")
def build_frd_endpoint():
    _check_key()
    data = request.get_json(force=True, silent=True)
    if not data or not isinstance(data, dict):
        return jsonify(error="Ожидается JSON-объект с данными наблюдения"), 400
    if not data.get("records"):
        return jsonify(error="Нет записей наблюдения (records)"), 400
    try:
        return _xlsx_response(build_frd(data), "balance.xlsx")
    except Exception as e:
        return jsonify(error=f"Ошибка сборки ФРД: {e}"), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))
