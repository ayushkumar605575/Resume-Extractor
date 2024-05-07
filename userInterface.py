from flask import Flask, render_template, request, send_file
import os
from pyzipper import AESZipFile
app = Flask(__name__)
import os
from flask import Flask, flash, request, redirect, url_for
from werkzeug.utils import secure_filename

from errorSolving import main

UPLOAD_FOLDER = 'uploads\\tmp'
ALLOWED_EXTENSIONS = {'zip', 'pdf','docx'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods = ['GET','POST'])
def mainPage():
    try: 
        os.remove("downloads\\extractedData.xlsx")
    except Exception:
        pass
    if request.method == 'POST':
        if 'file' not in request.files:
            flash(message='No file part', category='error')
            return '''<h3>Invalid File</h3><br><a href = "/"><button>Go Back</button></a>'''
        file = request.files['file']
        if file.filename == '':
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file_path = file_path.replace("/","\\")
            file.save(file_path)
            zip_file = AESZipFile(file_path)
            zip_file.extractall(UPLOAD_FOLDER)
            main(UPLOAD_FOLDER)
            return redirect('/download')
        else:
            return '''<h3>Invalid File</h3><br><a href = "/"><button>Go Back</button></a>'''
    return render_template('index.html')


@app.route('/download', methods = ['GET'])
def download():
    response = send_file(f"downloads\\extractedData.xlsx", as_attachment=True)
    return response

if __name__ == '__main__':
    try: 
        os.remove("downloads\\extractedData.xlsx")
    except Exception:
        pass
    try:
        os.mkdir("downloads")
    except Exception:
        pass
    try:
        os.mkdir("uploads")
    except Exception:
        pass
    try:
        os.mkdir("uploads\\tmp")
    except Exception:
        pass
    app.secret_key = "ndjzesuryksjuyh794hf9o47y8thrf4i9oyujhfry7489o5ywouejhfr9487y5hrwR$^%$T%$6$%^4gf3544i97ryhu49o8yhe4teg54"
    app.run(debug=True)