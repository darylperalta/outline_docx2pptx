import os


from flask import Flask, request, render_template, send_from_directory, send_file
from lyrics2pptx import Lyrics2pptx
from outline2pptx import Outline2pptx


app = Flask(__name__)



APP_ROOT = os.path.dirname(os.path.abspath(__file__))
INPUT_TEMPLATE = os.path.join(APP_ROOT, 'doc_templates/template_lyrics.pptx')
TEMPLATE_OUTLINE = os.path.join(APP_ROOT, 'doc_templates/template_2020.pptx')
TEMPLATE_OUTLINE2 = os.path.join(APP_ROOT, 'doc_templates/template2_2020.pptx')


@app.route("/")
def index():
    return render_template("home.html")

@app.route("/lyrics_home")
def lyrics_home():
    return render_template("upload.html")

@app.route("/about")
def about():
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
        # print(upload)
        # print("{} is the file name".format(upload.filename))
        filename = upload.filename
        destination = "/".join([target, filename])
        # print ("Accept incoming file:", filename)
        # print ("Save it to:", destination)
        upload.save(destination)
        input_docx = os.path.join(target, filename)
        # print('input dox in', input_docx)

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

@app.route("/outline_home")
def outline_home():
    return render_template("upload_outline.html")

@app.route("/upload_outline", methods=["POST"])
def upload_outline():
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
        # print ("Save it to:", destination)
        # upload.save(destination)
        input_docx = os.path.join(target, filename)
        upload.save(input_docx)
        print('input dox in', input_docx)

    print('input template', TEMPLATE_OUTLINE)

    print('input docx',input_docx)
    print(len(input_docx))
    out_fn1 = os.path.join(APP_ROOT, 'downloads/out.pptx')
    out_fn2 = os.path.join(APP_ROOT, 'downloads/out2.pptx')
    
    out2pptx = Outline2pptx(verbose=False, doc_fn = input_docx, template_pptx = TEMPLATE_OUTLINE, out_pptx=out_fn1, template_id =0)
    out2pptx.create_pptx()
    out2pptx = Outline2pptx(verbose=False, doc_fn = input_docx, template_pptx = TEMPLATE_OUTLINE2, out_pptx=out_fn2, template_id =1)
    out2pptx.create_pptx()
    print('Powerpoint written at ', out_fn1, 'and', out_fn2)

    return render_template("complete_outline.html", out_name1='out.pptx', out_name2='out2.pptx')

@app.route('/download_outline_morning/<filename>')
def download_outline_morning(filename):
    return send_from_directory('downloads',filename)

@app.route('/download_outline_evening/<filename>')
def download_outline_evening(filename):
    return send_from_directory('downloads', filename)


# @app.route('/gallery')
# def get_gallery():
#     image_names = os.listdir('./images')
#     print(image_names)
#     return render_template("gallery.html", image_names=image_names)

if __name__ == "__main__":
    app.run(port=4555, debug=True)