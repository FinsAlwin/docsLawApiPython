from flask import Flask, jsonify, request, send_file, make_response
from flask_cors import CORS, cross_origin
import uuid
from firebase_admin import storage, credentials, initialize_app
from dotenv import load_dotenv
from doc_creator import DocCreator
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
               data['placeHolder'], file_name)

    # Upload the .docx file to Firebase Storage
    # bucket = storage.bucket()
    # blob = bucket.blob(file_name)
    # with open(file_name, "rb") as file:
    #     blob.upload_from_file(file)
    #     blob.make_public()

    # Delete the .docx file from the local system
    # os.remove(file_name)

    # download_url = blob.public_url

    return jsonify({"message": f"File {file_name} uploaded successfully.", "download_url": "soon"}), 200


if __name__ == '__main__':
    app.run(host="0.0.0.0", port=5000)
