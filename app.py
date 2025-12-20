from flask import Flask, request, send_file
import os
import subprocess
import tempfile


from flask import (Flask, redirect, render_template, request,
                   send_from_directory, url_for)

app = Flask(__name__)


@app.route('/')
def index():
   print('Request for index page received')
   return render_template('index.html')

@app.route('/favicon.ico')
def favicon():
    return send_from_directory(os.path.join(app.root_path, 'static'),
                               'favicon.ico', mimetype='image/vnd.microsoft.icon')

@app.route('/hello', methods=['POST'])
def hello():
   name = request.form.get('name')

   if name:
       print('Request for hello page received with name=%s' % name)
       return render_template('hello.html', name = name)
   else:
       print('Request for hello page received with no name or blank name -- redirecting')
       return redirect(url_for('index'))
   
@app.route('/convert', methods=['POST'])
def convert_excel_to_pdf():
    if 'file' not in request.files:
        return "No file", 400

    file = request.files['file']
    with tempfile.TemporaryDirectory() as tmpdir:
        excel_path = os.path.join(tmpdir, file.filename)
        pdf_path = os.path.join(tmpdir, "output.pdf")
        file.save(excel_path)

        # LibreOffice CLIで変換
        subprocess.run([
            "libreoffice",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", tmpdir,
            excel_path
        ], check=True)

        return send_file(pdf_path, as_attachment=True)


if __name__ == '__main__':
   app.run()
