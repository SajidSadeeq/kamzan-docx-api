from flask import Flask, send_from_directory, send_file
from flask import request, jsonify
import json
from docxtpl import DocxTemplate
from docx import Document
from flask_cors import CORS
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
# app.config["DEBUG"] = True
# UPLOAD_FOLDER = './uploads'
# app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER 
CORS(app)

# @app.route('/<name>')
# def index(name):
#     return '<h1>Hello </h1>'.format(name)

@app.route('/', methods=['GET'])
def home():
    return "<h1>Distant Reading Archive</h1><p>This site is a prototype API for distant reading of science fiction novels.</p>"

# Create some test data for our catalog in the form of a list of dictionaries.
books = [
    {'id': 0,
     'title': 'A Fire Upon the Deep',
     'author': 'Vernor Vinge',
     'first_sentence': 'The coldsleep itself was dreamless.',
     'year_published': '1992'},
    {'id': 1,
     'title': 'The Ones Who Walk Away From Omelas',
     'author': 'Ursula K. Le Guin',
     'first_sentence': 'With a clamor of bells that set the swallows soaring, the Festival of Summer came to the city Omelas, bright-towered by the sea.',
     'published': '1973'},
    {'id': 2,
     'title': 'Dhalgren',
     'author': 'Samuel R. Delany',
     'first_sentence': 'to wound the autumnal city.',
     'published': '1975'}
]

# A route to return all of the available entries in our catalog.
@app.route('/api/v1/resources/books', methods=['GET'])
def api_all():
    # Check if an ID was provided as part of the URL.
    # If ID is provided, assign it to a variable.
    # If no ID is provided, display an error in the browser.
    if 'id' in request.args:
        id = int(request.args['id'])
    else:
        return "Error: No id field provided. Please specify an id."
    
    results = []
    
    for book in books:
        if book['id'] == id:
            results.append(book)
    
    return jsonify(results)

# A route to return all of the available entries in our catalog.
@app.route('/api/v1/resources/firstapi', methods=['GET','POST'])
def first_api():
    if request.method == 'POST':
        data = []
        output_file_path = 'result.docx'
        # template_file_path = 'Retelit_Products.docx'
        if "document" in request.files:
            document = request.files["document"]
            template_file_path = document.filename
            # document.save(os.path.join(app.config['UPLOAD_FOLDER'],document.filename))
            document.save(os.path.join(document.filename))
            template_document = Document(template_file_path)
            
            replacesImageName = []
            files = request.files.getlist("replace_images[]")
            for file in files:
                print(file.filename)
                replacesImageName.append(file.filename)
                file.save(os.path.join(file.filename))
                
            findsImageName = []
            files = request.files.getlist("find_images[]")
            for file in files:
                print(file.filename)
                findsImageName.append(file.filename)
                file.save(os.path.join(file.filename))
            
            find_text = request.form.getlist('find_text[]')
            replace_text = request.form.getlist('replace_text[]')
            words = []
            if len(find_text) == len(replace_text):
                for i in range(len(replace_text)):
                    prepareWord = {}
                    prepareWord['find'] = find_text[i]
                    prepareWord['replace'] = replace_text[i]
                    words.append(prepareWord)
            
            if words:
                # words = request_data['inputF']
                for word in words:
                    print("word : ",word['find'])
                    for paragraph in template_document.paragraphs:
                        replace_text_in_paragraph(paragraph, word['find'], word['replace'])

                    for table in template_document.tables:
                        for col in table.columns:
                            for cell in col.cells:
                                for paragraph in cell.paragraphs:
                                    replace_text_in_paragraph(paragraph, word['find'], word['replace'])
                    
                    for section in template_document.sections:
                        for hf in section.header.paragraphs:
                            replace_text_in_paragraph(hf, word['find'], word['replace'])
                            
                    for section in template_document.sections:
                        for hf in section.footer.paragraphs:
                            replace_text_in_paragraph(hf, word['find'], word['replace'])

            template_document.save(output_file_path)
            
            doc = DocxTemplate(output_file_path)
            doc.reset_replacements()
            if len(findsImageName) == len(replacesImageName):
                for i in range(len(findsImageName)):
                    doc.replace_media(findsImageName[i],replacesImageName[i])        
            # doc.replace_media('header.jpg','header-replace.jpg')
            # doc.replace_media('main.png','main-replace.png')
            # doc.replace_media('map.png','map-replace.png')
            return_file = doc
            doc.save(output_file_path)
            # words = request.form.getlist('words')
            # return send_from_directory(directory='/', filename=return_file, as_attachment=True)
            # filename = os.path.join(app.root_path, '/', output_file_path)
            # return_data = [os.path.dirname(app.instance_path), output_file_path]
            return jsonify(isError= False,
                            message= "Success",
                            statusCode= 200,
                            data= output_file_path), 200
    return '''
    <!doctype html>
    <title>Upload new File</title>
    <h1>Upload new File</h1>
    <form method=post enctype=multipart/form-data>
      <input type=file name=file[] multiple=true>
      <input type=submit value=Upload>
    </form>
    '''
@app.route('/download/<filename>')
def downloadFile (filename):
    return send_file(filename, as_attachment=True)

def replace_text_in_paragraph(paragraph, key, value):
    if key in paragraph.text:
        inline = paragraph.runs
        for item in inline:
            if key in item.text:
                item.text = item.text.replace(key, value)
# app.run(debug=True ,port=8080,use_reloader=False)
# loop
if __name__ == "__main__":
    app.run()
#     from waitress import serve
#     serve(debug=True ,port=8080,use_reloader=False)