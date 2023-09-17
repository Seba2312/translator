from docx import Document
from docx.shared import Pt
from docx.opc.exceptions import PackageNotFoundError
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
import openai
from googletrans import Translator

openai.api_key = "your api key"


def improvegrammer(text):
    model_engine = "text-davinci-003"  #  "text-davinci-003"  is good, cheaper options give worse text, before choosing test with a small file to choose optimized
    response = openai.Completion.create(
        engine=model_engine,
        prompt=f"Please correct any grammatical or stylistic errors in the following text without altering the original message. You may rewrite sentences if that improves their clarity and makes them easier to read. Ensure that the returned text is a single, coherent paragraph free from repetition or misplaced sentences: {text}",
        max_tokens=2000  #More tokens means clearer text, but might be more expensvie
    )
    corrected_text = response.choices[0].text.strip()

    # Remove the first sentence if it starts with a lowercase letter
    if corrected_text and corrected_text[0].islower():
        corrected_text = corrected_text[corrected_text.find('. ') + 2:]
    return corrected_text


def translate_to_english(text):
    translator = Translator()
    translated = translator.translate(text, src='pt', dest='en')  #choose any languages i have chosen pt(portuguese to english)
    #do to prompt being in english it is optmized for translating english
    return translated.text


def add_tab_to_sentences(text):
    return '\t' + text


try:
    # Initialize Document objects for reading and writing
    input_document = Document(r"the location of the file to be translated, ending with .docx")
except PackageNotFoundError as e:
    print(f"PackageNotFoundError: {e}")
    exit()

# Initialize a new Document object for writing
output_document = Document()


try:
    # Loop through each paragraph in the input document, the translation is done by paragraph
    for para in input_document.paragraphs:

        # Skip footnotes based on font size and the presence of a line, if footnotes are sized 10, change accordingly
        if any(run.font.size == Pt(10) for run in para.runs):
            continue

        # Create a new paragraph in the output document
        try:
            translated_text = translate_to_english(para.text)
            if len(translated_text) > 50:  #for bug fixing and error hand
                print("Translated Text")
                print(translated_text)
                print("Fixed by spellcheck and style-check ")
                translated_text = improvegrammer(translated_text)
                print(translated_text)
                print("")

            new_text = add_tab_to_sentences(translated_text)
            new_para = output_document.add_paragraph()
            new_para.style = para.style
            new_para.alignment = para.alignment

            new_run = new_para.add_run(new_text)

            first_run = para.runs[0] if para.runs else None
            if first_run:
                new_run.bold = first_run.bold
                new_run.italic = first_run.italic
                new_run.underline = first_run.underline
                if first_run.font.size:
                    new_run.font.size = Pt(12) if first_run.font.size.pt in [11, 12] else first_run.font.size
                else:
                    new_run.font.size = Pt(12)
                new_run.font.name = 'Times New Roman'  # Set font to chosen font
        except Exception as e:
            print(f"An error occurred : {e}")
except Exception as e:
    print(f"An error occurred while processing the paragraphs: {e}")

try:
    # Save the output document
    output_document.save(r"C:\location where file created last file must have .docx (the one being created)")
except Exception as e:
    print(f"An error occurred while saving the document: {e}")
