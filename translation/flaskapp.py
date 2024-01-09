# translator_service.py
from flask import Flask, request, jsonify
from googletrans import Translator

app = Flask(__name__)

@app.route('/translate', methods=['POST'])
def translate_text():
    try:
        data = request.json
        print("here")
        print(data)
        print("here1")
        text = data['text']
        src_lang = data.get('src', 'de')
        dest_lang = data.get('dest', 'en')

        print(text)
        translator = Translator()
        print("here2")

        try:
            translation = translator.translate(text,  dest=dest_lang,src=src_lang)
            return jsonify({"translated_text": translation.text})
        except Exception as translate_error:
            print("An error occurred during translation", str(translate_error))
            return jsonify({"error": str(translate_error)}), 500

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == '__main__':
    app.run(port=5001)
