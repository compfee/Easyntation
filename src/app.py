from flask import Flask, jsonify
from flask import request
from text_parsing import parse_docx

app = Flask(__name__)


def prepare_doc(doc):
    list_of_pars = parse_docx(doc)
    return list_of_pars

@app.route('/predict', methods=['POST'])
def predict():
    if request.method == 'POST':
        # we will get the file from the request
        file = request.files['file']
        # convert that to bytes
        img_bytes = file.read()
        return

if __name__ == "__main__":
    app.run()