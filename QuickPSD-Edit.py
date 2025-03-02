# Requires Python 3.x, Flask, and pywin32
from flask import Flask, send_file, request
import os
import win32com.client
import pythoncom
import time

# Flask (Web Server)
ws = Flask(__name__)

# Specify the directory where your PSD files are located
PSD_DIRECTORY = r'C:\Users\Administrator\Desktop\PSD Templates'

@ws.route("/")
def root_index():
    files = os.listdir(PSD_DIRECTORY)
    response = []
    for f in files:
        if os.path.isfile(os.path.join(PSD_DIRECTORY, f)) and f.lower().endswith('.psd'):
            response.append("""
<a href="/%s">%s</a><br/>
""" % (f, f))
    return ''.join(response)

@ws.route("/<filename>", methods=['GET'])
def view_file(filename):
    if not filename.lower().endswith('.psd'):
        return ''

    response = []
    raw = request.args.get('raw')
    if raw == '1':
        return send_file(os.path.join(PSD_DIRECTORY, filename), mimetype='image/png')

    pythoncom.CoInitialize()
    psApp = win32com.client.Dispatch("Photoshop.Application")

    doc = psApp.Open(os.path.join(PSD_DIRECTORY, filename))
    layerCount = doc.ArtLayers.Count

    response.append('<form method="post" action="/%s/submit">' % filename)

    for i in range(layerCount):
        response.append("""Layer : %s<br/>
""" % (doc.ArtLayers[i].Name))
        if doc.ArtLayers[i].Kind == 2:
            # Replace Photoshop's internal line break character with '\n' for displaying in textarea
            text_content = doc.ArtLayers[i].TextItem.Contents.replace('', '\n')
            response.append("""
- Text : %s
<input type="checkbox" name="layer_%d"> Edit Layer %d<br/>
<textarea name="text_%d" rows="4" cols="50">%s</textarea>
<br/>
""" % (text_content, i, i, i, text_content))

    response.append("""
<label for="format">Choose format:</label>
<input type="radio" id="png" name="format" value="png" checked>
<label for="png">PNG</label>
<input type="radio" id="psd" name="format" value="psd">
<label for="psd">PSD</label>
<br/>
<input type="submit" value="Submit All" />
</form>
<br/>
""")

    doc.Close(2)
    return ''.join(response)

@ws.route("/<filename>/submit", methods=['POST'])
def submit_all(filename):
    pythoncom.CoInitialize()
    psApp = win32com.client.Dispatch("Photoshop.Application")

    doc = psApp.Open(os.path.join(PSD_DIRECTORY, filename))
    layerCount = doc.ArtLayers.Count

    for i in range(layerCount):
        if 'layer_%d' % i in request.form:
            # Get the submitted text
            new_text = request.form['text_%d' % i]

            # Split the text into lines and join them back with Photoshop's line break character ''
            new_text_cleaned = ''.join(new_text.splitlines())

            # Apply the cleaned text back to the layer
            doc.ArtLayers[i].TextItem.Contents = new_text_cleaned

    format = request.form.get('format', 'png')

    if format == 'png':
        # Save the modified document as PNG
        f1 = filename + time.strftime(".%Y%m%d_%H%M%S") + '.png'
        pngFilename = os.path.join(PSD_DIRECTORY, f1)
        options = win32com.client.Dispatch('Photoshop.ExportOptionsSaveForWeb')
        options.Format = 13   # PNG
        options.PNG8 = False  # Sets it to PNG-24 bit
        doc.Export(ExportIn=pngFilename, ExportAs=2, Options=options)
    elif format == 'psd':
        # Save the modified document as PSD
        f1 = filename + time.strftime(".%Y%m%d_%H%M%S") + '.psd'
        psdFilename = os.path.join(PSD_DIRECTORY, f1)
        doc.SaveAs(psdFilename)

    doc.Close(2)

    return '<a href="/%s?raw=1">%s</a>' % (f1, f1)

if __name__ == "__main__":
    ws.run(host='0.0.0.0', port=5000, debug=True)