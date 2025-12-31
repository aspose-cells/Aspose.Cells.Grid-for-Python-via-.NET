# app.py
# -------------------------------------------------------------
# Flask application that hosts Aspose.Cells GridJs services.
# -------------------------------------------------------------

from flask import (
    Flask,
    render_template,
    request,
    jsonify,
    Response,
    send_file,
    abort,
)
import io
import gzip
import mimetypes
import os

# Aspose.Cells GridJs imports
from aspose.cellsgridjs import GridJsOptions, GridJsService

app = Flask(__name__)

# -------------------------------------------------------------
# GridJs service initialization
# -------------------------------------------------------------
options = GridJsOptions()
options.file_cache_directory = "./cache"           # Cache folder (must exist)
options.base_route_name = "/GridJs"                # Base route used by the client
gridjs_service = GridJsService(options)

# Ensure cache folder exists
os.makedirs(options.file_cache_directory, exist_ok=True)

# -------------------------------------------------------------
# Helper: Guess MIME type from a filename
# -------------------------------------------------------------
def guess_mime_type_from_filename(filename: str) -> str:
    mime_type, _ = mimetypes.guess_type(filename)
    return mime_type or "application/octet-stream"

# -------------------------------------------------------------
# Root route – renders the UI
# -------------------------------------------------------------
@app.route("/")
def index():
    return render_template("index.html")

# -------------------------------------------------------------
# 1️⃣ Load spreadsheet JSON (gzipped)
# -------------------------------------------------------------
@app.route("/GridJs/LoadSpreadsheet", methods=["GET"])
def load_spreadsheet():
    uid = request.args.get("uid", "")
    # Prepare an in‑memory gzip stream
    gzip_buffer = io.BytesIO()
    with gzip.GzipFile(fileobj=gzip_buffer, mode="w") as gz:
        # Populate the gzip stream with GridJs JSON data
        gridjs_service.detail_stream_json_with_uid(gz, "./data/sample.xlsx", uid)
    gzip_buffer.seek(0)
    return Response(
        gzip_buffer.getvalue(),
        mimetype="application/json",
        headers={"Content-Encoding": "gzip"},
    )

# -------------------------------------------------------------
# 2️⃣ Update a cell (POST)
# -------------------------------------------------------------
@app.route("/GridJs/UpdateCell", methods=["POST"])
def update_cell():
    p = request.form.get("p")
    uid = request.form.get("uid")
    ret = gridjs_service.update_cell(p, uid)
    return Response(ret, content_type="text/plain; charset=utf-8")

# -------------------------------------------------------------
# 3️⃣ Add image (multipart/form‑data)
# -------------------------------------------------------------
@app.route("/GridJs/AddImage", methods=["POST"])
def add_image():
    uid = request.form.get("uid")
    p = request.form.get("p")
    is_control = request.form.get("control")
    file = request.files.get("image")          # May be None – handled below
    file_bytes = io.BytesIO(file.read()) if file else None
    ret = gridjs_service.add_image(p, uid, is_control, file_bytes)
    return jsonify(ret)

# -------------------------------------------------------------
# 4️⃣ Copy image
# -------------------------------------------------------------
@app.route("/GridJs/CopyImage", methods=["POST"])
def copy_image():
    uid = request.form.get("uid")
    p = request.form.get("p")
    ret = gridjs_service.copy_image(p, uid)
    return jsonify(ret)

# -------------------------------------------------------------
# 5️⃣ Add image by external URL
# -------------------------------------------------------------
@app.route("/GridJs/AddImageByURL", methods=["POST"])
def add_image_by_url():
    uid = request.form.get("uid")
    p = request.form.get("p")
    image_url = request.form.get("imageurl")
    ret = gridjs_service.add_image_by_url(p, uid, image_url)
    return jsonify(ret)

# -------------------------------------------------------------
# 6️⃣ Retrieve an image (GET)
# -------------------------------------------------------------
@app.route("/GridJs/Image", methods=["GET"])
def image():
    img_id = request.args.get("id")
    uid = request.args.get("uid")
    if not img_id or not uid:
        return "Missing required parameters", 400
    image_bytes = gridjs_service.image(uid, img_id)
    return send_file(
        image_bytes,
        as_attachment=False,
        download_name=img_id,
        mimetype="image/png",
    )

# -------------------------------------------------------------
# 7️⃣ Retrieve an embedded OLE object (GET)
# -------------------------------------------------------------
@app.route("/GridJs/Ole", methods=["GET"])
def ole():
    obj_id = request.args.get("id")
    uid = request.args.get("uid")
    sheet = request.args.get("sheet")
    filename = None
    file_bytes = gridjs_service.ole(uid, sheet, obj_id, filename)
    if filename:
        return send_file(
            io.BytesIO(file_bytes),
            as_attachment=True,
            download_name=filename,
            mimetype=guess_mime_type_from_filename(filename),
        )
    abort(400, "File not found")

# -------------------------------------------------------------
# 8️⃣ Batch image URLs (GET)
# -------------------------------------------------------------
@app.route("/GridJs/ImageUrl", methods=["GET"])
def image_url():
    img_id = request.args.get("id")
    uid = request.args.get("uid")
    ret = gridjs_service.image_url(options.base_route_name, img_id, uid)
    return jsonify(ret)

# -------------------------------------------------------------
# 9️⃣ Download cached file (GET)
# -------------------------------------------------------------
@app.route("/GridJs/GetFile", methods=["GET"])
def get_file():
    file_id = request.args.get("id")
    file_bytes = gridjs_service.get_file(file_id)
    return send_file(
        file_bytes,
        as_attachment=True,
        download_name=file_id,
        mimetype=guess_mime_type_from_filename(file_id),
    )

# -------------------------------------------------------------
# 🔟 Trigger file download (POST)
# -------------------------------------------------------------
@app.route("/GridJs/Download", methods=["POST"])
def download():
    p = request.form.get("p")
    uid = request.form.get("uid")
    file_name = request.form.get("file")
    ret = gridjs_service.download(p, uid, file_name)
    return jsonify(ret)

# -------------------------------------------------------------
# Application entry point
# -------------------------------------------------------------
if __name__ == "__main__":
    # Listen on all interfaces as required
    app.run(host="0.0.0.0", port=5000, debug=True)
