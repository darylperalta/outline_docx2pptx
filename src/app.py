import os


from flask import Flask, request, render_template, send_from_directory, send_file
from lyrics2pptx import Lyrics2pptx

__author__ = 'ibininja'

app = Flask(__name__)



APP_ROOT = os.path.dirname(os.path.abspath(__file__))
INPUT_TEMPLATE = os.path.join(APP_ROOT, 'doc_templates/template_lyrics.pptx')


@app.route("/")
def index():
    return render_template("upload.html")

@app.route("/upload", methods=["POST"])
def upload():
    target = os.path.join(APP_ROOT, 'input/')
    print(target)
    if not os.path.isdir(target):
            os.mkdir(target)
    else:
        print("Couldn't create upload directory: {}".format(target))
    print(request.files.getlist("file"))
    input_docx = ''
    for upload in request.files.getlist("file"):
        print(upload)
        print("{} is the file name".format(upload.filename))
        filename = upload.filename
        destination = "/".join([target, filename])
        print ("Accept incoming file:", filename)
        print ("Save it to:", destination)
        upload.save(destination)
        input_docx = os.path.join(target, filename)
        print('input dox in', input_docx)

    out_fn = os.path.join(APP_ROOT, 'downloads/out_lyrics.pptx')
    print('input template', INPUT_TEMPLATE)

    print('input docx',input_docx)
    print(len(input_docx))
    out2pptx = Lyrics2pptx(verbose=False, doc_fn = input_docx, template_pptx = INPUT_TEMPLATE, out_pptx=out_fn)
    out2pptx.create_pptx()

    # return send_from_directory("images", filename, as_attachment=True)
    return render_template("complete.html", image_name=filename)

@app.route('/download_lyrics')
def download_lyrics():
    return send_file("downloads/out_lyrics.pptx")

@app.route('/upload/<filename>')
# def send_image(filename):
#     return send_from_directory("images", filename)

# @app.route('/gallery')
# def get_gallery():
#     image_names = os.listdir('./images')
#     print(image_names)
#     return render_template("gallery.html", image_names=image_names)

# if __name__ == "__main__":
#     app.run(port=4555, debug=True)