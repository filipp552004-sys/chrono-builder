"""chrono-builder — микросервис сборки карты хронометража (Excel).
Эндпоинт POST /build принимает JSON с данными и возвращает .xlsx (bytes).
Аутентификация: заголовок X-API-Key должен совпадать с env CHRONO_BUILDER_API_KEY.
"""
import os, io
from flask import Flask, request, send_file, jsonify, abort
from builder import build

app = Flask(__name__)
API_KEY = os.environ.get("CHRONO_BUILDER_API_KEY", "")

@app.get("/health")
def health():
    return jsonify(status="ok", service="chrono-builder")

@app.post("/build")
def build_endpoint():
    # проверка ключа
    if not API_KEY or request.headers.get("X-API-Key", "") != API_KEY:
        abort(401)
    data = request.get_json(force=True, silent=True)
    if not data or not isinstance(data, dict):
        return jsonify(error="Ожидается JSON-объект с данными хронометража"), 400
    if not data.get("elements"):
        return jsonify(error="Нет операций-составляющих (elements)"), 400
    try:
        wb = build(data)
        buf = io.BytesIO(); wb.save(buf); buf.seek(0)
        return send_file(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="norma.xlsx",
        )
    except Exception as e:
        return jsonify(error=f"Ошибка сборки: {e}"), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))
