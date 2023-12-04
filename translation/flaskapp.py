# translator_service.py
from flask import Flask, request, jsonify
from googletrans import Translator

app = Flask(__name__)

@app.route('/translate', methods=['POST'])
def translate_text():
    try:
        data = request.json
        text = data['text']
        src_lang = data.get('src', 'de')
        dest_lang = data.get('dest', 'en')

        translator = Translator()
        translation = translator.translate(text, src=src_lang, dest=dest_lang)
        print(translation.text)
        return jsonify({"translated_text": translation.text})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == '__main__':
    app.run(port=5006)
