from flask import Flask, jsonify, request, send_file, make_response
from flask_cors import CORS, cross_origin
import uuid
from firebase_admin import storage, credentials, initialize_app
from dotenv import load_dotenv
from doc_creator import DocCreator, Single_docx
import os
from word_processor import WordProcessor

load_dotenv()
WordProcessor = WordProcessor()

app = Flask(__name__)
cors = CORS(app, resources={r"/api/*": {"origins": "*"}})


cred = credentials.Certificate(
    "docslaw-9e938-firebase-adminsdk-wgfbv-3771ec843b.json")
initialize_app(cred, {
    "storageBucket": "docslaw-9e938.appspot.com"
})


@app.route('/api/v1/genDocx', methods=['POST'])
@cross_origin()
def create_doc():
    # Get data from the request
    data = request.get_json()

    file_extension = ".docx"
    file_name = f"{str(uuid.uuid4())}{file_extension}"

    DocCreator(data['isUrgent'], data['indexList'],
               data['placeHolder'], file_name, data['newContent'])

    # Upload the .docx file to Firebase Storage
    bucket = storage.bucket()
    blob = bucket.blob(file_name)
    with open(file_name, "rb") as file:
        blob.upload_from_file(file)
        blob.make_public()

    # Delete the .docx file from the local system
    os.remove(file_name)

    download_url = blob.public_url

    return jsonify({"message": f"File {file_name} uploaded successfully.", "download_url": download_url}), 200


@app.route('/api/v1/initDocx', methods=['POST'])
@cross_origin()
def init_docx():
    # Get data from the request
    data = request.get_json()

    single_docx = Single_docx()

    doc_data = []

    valid_titles = [
        "Urgent Application",
        "Notice of Motion",
        "Memo of Parties",
        "Synopsis & List of Dates",
    ]

    for title in data['indexList']:
        if title in valid_titles:
            base64_content = single_docx.get_docx(title, data['placeHolder'])
            is_docs = True
        else:
            base64_content = ""
            is_docs = False

        # Create a dictionary with the title and base64 content for the document
        doc_dict = {
            'title': title,
            'content': {
                'isDocs': is_docs,
                'data': {
                    'base64String': "",
                },
            },
        }

        doc_data.append(doc_dict)

    return jsonify({"message": "ok", 'data': doc_data}), 200


if __name__ == '__main__':
    app.run(host="0.0.0.0", port=5000)
