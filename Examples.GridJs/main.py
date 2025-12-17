# This is a demo to show how to use GridJs .
import configparser
import gzip
import io
import mimetypes
import os
from aspose.cellsgridjs import *
import requests
from flask import Flask, render_template, jsonify, request, Response, send_file, abort

config = configparser.ConfigParser()
config.read('config.ini')
app=Flask(__name__)
# your working file directory which has spreadsheet files inside wb directory，
FILE_DIRECTORY = os.path.join(os.getcwd(),'wb')
UPLOAD_FOLDER = os.path.join(os.getcwd(),'upload')
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# example usage for PdfSaveOptions
pdf_save_opt = PdfSaveOptions()
pdf_save_opt.pdf_compression = PdfCompressionCore.LZW
pdf_save_opt.all_columns_in_one_page_per_sheet = True
# choose sheet index 0 for pdf result
sheet_indices = [0]
pdf_save_opt.set_sheet_set(sheet_indices)

options = GridJsOptions()
options.custom_pdf_save_options = pdf_save_opt
# whether to load worksheets with lazy loading
options.lazy_loading = True
# set storage cache directory for GridJs
options.file_cache_directory = config.get('DEFAULT', 'CacheDir')
gridjs_service = GridJsService(options)




@app.route('/')
def index():
    filename=config.get('DEFAULT', 'FileName')
    uid = GridJsWorkbook.get_uid_for_file(filename)
    return render_template('uidload.html', filename=filename, uid=uid)


@app.route('/list')
def list():
    files = os.listdir(FILE_DIRECTORY)
    return render_template('list.html', files=files)

@app.route('/Uidtml', methods=['GET'])
def uidtml():
    filename = request.args.get('filename')
    uid = request.args.get('uid')
    return render_template('uidload.html',filename= filename,uid= uid)


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    try:
        uid = GridJsWorkbook.get_uid_for_file(file.filename)
        file.save(os.path.join(UPLOAD_FOLDER, file.filename))
        return render_template('uidload.html', filename=file.filename, uid=uid, fromupload=1)
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/GridJs2/DetailStreamJsonWithUid', methods=['GET'])
def detail_stream_json_with_uid():
    filename = request.args.get('filename')
    uid = request.args.get('uid')
    from_upload = request.args.get('fromUpload')
    if not filename:
        return jsonify({'error': 'filename is required'}), 400
    if not uid:
        return jsonify({'error': 'uid is required'}), 400
    if not from_upload:
        file_path = os.path.join(FILE_DIRECTORY, filename)
    else:
        file_path = os.path.join(UPLOAD_FOLDER, filename)
    try:

        print("\nfile path is:" + file_path)

        output = io.BytesIO()
        with gzip.GzipFile(fileobj=output, mode='wb', compresslevel=9) as gzip_stream:

            gridjs_service.detail_stream_json_with_uid(gzip_stream, file_path, uid)

        response = Response(output.getvalue(), mimetype='application/json')
        response.headers['Content-Encoding'] = 'gzip'

        return response
    except Exception as e:
        return Response(str(e), status=500)


@app.route('/GridJs2/LazyLoading', methods=['POST'])
def lazy_loading():
    sheet_name = request.form.get('name', '')
    uid = request.form.get('uid', '')
    if not sheet_name:
        return jsonify({'error': 'sheet_name is required'}), 400
    if not uid:
        return jsonify({'error': 'uid is required'}), 400

    try:

        output = io.BytesIO()
        with gzip.GzipFile(fileobj=output, mode='wb', compresslevel=9) as gzip_stream:
            gridjs_service.lazy_loading_stream_json(gzip_stream, sheet_name, uid)

        response = Response(output.getvalue(), mimetype='application/json')
        response.headers['Content-Encoding'] = 'gzip'

        return response
    except Exception as e:
        return Response(str(e), status=500)


# update action :/GridJs2/UpdateCell
@app.route('/GridJs2/UpdateCell', methods=['POST'])
def update_cell():
    # retrieve form data from the request
    p = request.form.get('p')
    uid = request.form.get('uid')

    # call the UpdateCell method and get the result
    ret = gridjs_service.update_cell(p, uid)

    # return a JSON response, as Flask defaults to returning JSON
    return Response(ret, content_type='text/plain; charset=utf-8')

# add image :/GridJs2/AddImage
@app.route('/GridJs2/AddImage', methods=['POST'])
def add_image():
    uid = request.form.get('uid')
    p = request.form.get('p')
    iscontrol = request.form.get('control')
    file = request.files.get('image')  # 使用 get() 避免 KeyError
    file_bytes = io.BytesIO(file.read()) if file else None
    ret = gridjs_service.add_image(p, uid, iscontrol, file_bytes)
                return jsonify(ret)


# copy image :/GridJs2/CopyImage
@app.route('/GridJs2/CopyImage', methods=['POST'])
def copy_image():
    uid = request.form.get('uid')
    p = request.form.get('p')
    ret = gridjs_service.copy_image(p, uid)
    return jsonify(ret)

def get_stream_from_url(url):
    response = requests.get(url)
    response.raise_for_status()  # if fail,raise HTTPError
    return io.BytesIO(response.content)

# add image by image source url:/GridJs2/AddImageByURL
@app.route('/GridJs2/AddImageByURL', methods=['POST'])
def add_image_by_url():
    uid = request.form.get('uid')
    p = request.form.get('p')
    imageurl = request.form.get('imageurl')
    ret = gridjs_service.add_image_by_url(p, uid, imageurl)
        return jsonify(ret)


# get image :/GridJs2/Image
@app.route('/GridJs2/Image', methods=['GET'])
def image():
    fileid = request.args.get('id')
    uid = request.args.get('uid')

    if fileid is None or uid is None:
        # if required parameters are missing, return an error response  
        return 'Missing required parameters', 400
    else:
        # retrieve the image stream  
        image_stream = gridjs_service.image(uid, fileid)

         # set the MIME type and attachment filename for the response (if needed)
        mimetype = 'image/png'
        attachment_filename = fileid

        #  send the file stream as the response  
        return send_file(
            image_stream,
            as_attachment=False,  # if sending as an attachment  
            download_name=attachment_filename,  # filename for download  
            mimetype=mimetype
        )


def guess_mime_type_from_filename(filename):
    # guess the MIME type based on the filename  
    mime_type, encoding = mimetypes.guess_type(filename)
    if mime_type is None:
        # if not found, return the default binary MIME type  
        mime_type = 'application/octet-stream'
    return mime_type


# get ole file: /GridJs2/Ole?uid=&id=
@app.route('/GridJs2/Ole', methods=['GET'])
def ole():
    oleid = request.args.get('id')
    uid = request.args.get('uid')
    sheet = request.args.get('sheet')
    filename = None
    filebyte = gridjs_service.ole(uid, sheet, oleid, filename)
    if filename is not None:

        # retrieve the image stream  
        ole_stream = io.BytesIO(filebyte)

    # set the MIME type and attachment filename for the response (if needed)
        mimetype = guess_mime_type_from_filename(filename)


    # send the file stream as the response 
        return send_file(
            ole_stream,
            as_attachment=True,  # if sending as an attachment  
            download_name=filename,  # filename for download
            mimetype=mimetype
        )
    else:
        # file not find
        abort(400, 'File not found')


# get batch zip image file url : /GridJs2/ImageUrl?uid=&id=
@app.route('/GridJs2/ImageUrl', methods=['GET'])
def image_url():
    id = request.args.get('id')
    uid = request.args.get('uid')
    ret = gridjs_service.image_url(Config.base_route_name, id, uid)
    return jsonify(ret)


# get file: /GridJs2/GetFile?id=&filename=
@app.route('/GridJs2/GetFile', methods=['GET'])
def get_file():
    id = request.args.get('id')

    filebt = gridjs_service.get_file(id)
    mimetype = guess_mime_type_from_filename(id)

            # set the MIME type application/zip
            # use send_file to send a file as a response  
            # as_attachment=True Send the file as an attachment，download_name Specify the filename for download  
    return send_file(filebt, as_attachment=True, download_name=id, mimetype=mimetype)


# download file :/GridJs2/Download
@app.route('/GridJs2/Download', methods=['POST'])
def download():
    p = request.form.get('p')
    uid = request.form.get('uid')
    filename = request.form.get('file')
    ret = gridjs_service.download(p, uid, filename)
    return jsonify(ret)


def do_at_start(name):
    # current_locale = locale.getencoding()
    # print(current_locale)
    # desired_culture = 'en_US.UTF-8'
    # locale.setlocale(locale.LC_ALL, desired_culture)
    print(f'Hi, {name}  {FILE_DIRECTORY} ')

    # set License for GridJs
    license_file = config.get('DEFAULT', 'LicenseFile')
    if os.path.exists(license_file):
        Config.set_license(license_file)
    # set Image route for GridJs,correspond with image()
    # GridJsWorkbook.set_image_url_base("/GridJs2/Image")
    # calc_engine = MyCalculation()
    # GridJsWorkbook.CalculateEngine = calc_engine
    print(f'{Config.file_cache_directory}')



if __name__ == '__main__':
    do_at_start('Hello GridJs python via .net')
    app.run(port=2022, host="0.0.0.0", debug=True)

