from flask import Flask, request, render_template, send_file
from pptx import Presentation
from io import BytesIO
from gtts import gTTS
from pathlib import Path
import re
import tempfile
import os
from flask_socketio import SocketIO

app = Flask(__name__)
socketio = SocketIO(app, cors_allowed_origins="*")

# Ensure the 'uploads/' directory exists
uploads_dir = 'uploads'
if not os.path.exists(uploads_dir):
    os.makedirs(uploads_dir)


@socketio.on('connect')
def test_connect():
    print('Client connected')


@socketio.on('disconnect')
def test_disconnect():
    print('Client disconnected')


# Define the clean_up_text function
def clean_up_text(input_text):
    # Add spaces after periods to separate sentences
    cleaned_text = re.sub(r'([.!?])', r'\1 ', input_text)
    
    # Replace multiple spaces with a single space
    cleaned_text = re.sub(r'\s+', ' ', cleaned_text)
    
    # Capitalize the first letter of each sentence
    sentences = cleaned_text.split('. ')
    cleaned_text = '. '.join(sentence.capitalize() for sentence in sentences)
    
    return cleaned_text


def extract_text_from_pptx(pptx_path):
    prs = Presentation(pptx_path)
    text_runs = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    text_runs.append(run.text)
    return ''.join(text_runs).strip()


def generate_mp3(text):
    cleaned_string = clean_up_text(text)
    tts = gTTS(cleaned_string, lang='en')
    mp3_fp = BytesIO()
    tts.write_to_fp(mp3_fp)
    return mp3_fp


def save_mp3(mp3_fp, mp3_filename):
    with open(mp3_filename, 'wb') as file:
        file.write(mp3_fp.getvalue())


def process_pptx_file(pptx_file):
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_file_path = os.path.join(temp_dir, pptx_file.filename)
        pptx_file.save(temp_file_path)

        text = extract_text_from_pptx(temp_file_path)
        mp3_fp = generate_mp3(text)
        
        mp3_filename = os.path.join(uploads_dir, pptx_file.filename + '.mp3')
        save_mp3(mp3_fp, mp3_filename)

        return mp3_filename


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'pptx_file' in request.files:
            pptx_file = request.files['pptx_file']
            if pptx_file.filename != '':
                # Notify the client that the conversion has started
                socketio.emit('conversion_started', namespace='/test')

                mp3_filename = process_pptx_file(pptx_file)

                # Notify the client that the conversion is complete
                socketio.emit('conversion_completed', namespace='/test')

                return send_file(mp3_filename, as_attachment=True)
            
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)