import os
import sys
import io
import tempfile
from flask import Flask, request, render_template_string, send_file, flash, redirect
from werkzeug.utils import secure_filename
from xml.dom.minidom import parse
from zipfile import ZipFile
import requests
import concurrent.futures
import csv
import glob
import shutil
import re
import urllib3
import secrets

# Suppress certificate validation warnings
urllib3.disable_warnings()


app = Flask(__name__)
app.secret_key = secrets.token_hex(32)
print(f'Starting pptxurlcheck with secret key {app.secret_key}')

# Allowed file extensions
ALLOWED_EXTENSIONS = {'pptx'}
MAX_REDIRECT = 10
TIMEOUT = 10
CONNECTIONS = 20

paragraphtext = ''


def parse_node(root):
    """
    Parse the given root recursively (root is intended to be the paragraph element <a:p>).
    If we encounter a link-break element a:br, add a new line to global paragraphtext.
    If we encounter an element with type TEXT_NODE, append value to paragraphtext.
    """
    global paragraphtext
    if root.childNodes:
        for node in root.childNodes:
            if node.nodeType == node.TEXT_NODE:
                paragraphtext += node.nodeValue.encode(
                    'ascii', 'ignore').decode('utf-8')
            if node.nodeType == node.ELEMENT_NODE:
                if node.tagName == 'a:br':
                    paragraphtext += '\n'
                parse_node(node)


def trim_url(url):
    """
    Remove extraneous trailing punctuation.

    First, strip common trailing punctuation (. , ; : ! ?).
    Then, while the URL ends with a ')' and there are more closing than opening
    parentheses in the URL (indicating an extra trailing parenthesis), remove one.
    """
    # Remove trailing punctuation except for parentheses
    trailing_punc = ".,;:!?"
    url = url.rstrip(trailing_punc)
    # Remove any unmatched trailing closing parentheses.
    while url.endswith(')') and url.count('(') < url.count(')'):
        url = url[:-1]
    return url


def make_csv(urlchkres, skip200=True):
    """
    Generate a CSV report from the results of the URL check.
    """

    # Create an in-memory file-like object
    output = io.StringIO()
    csvwriter = csv.writer(output)

    # Write the header row
    csvwriter.writerow(["File Number", "Page Number", "Response", "URL", "Note"])

    for urldata in urlchkres:
        filenum, pagenum, url, response, note = urldata
        if response == 200 and skip200:
            continue
        csvwriter.writerow([filenum, pagenum, response, url, note])

    # Get the CSV content as a string
    csv_string = output.getvalue()
    output.close()

    return csv_string


def parse_pptx(pptxfiles):
    """
    Parse the given PowerPoint files for URLs and return a hash of URLs indexed
    by URL with [filenum,pagenum] as the value.
    Reads slide notes and slide text boxes and other text elements.
    """
    global paragraphtext
    urlmatchre = r'(?:^|[\s(<])(?P<url>(?:(?:https?):\/\/|www\.)\S+)'

    urls = {}
    filenum = 0

    for pptxfile in pptxfiles:
        filenum += 1

        tmpd = tempfile.mkdtemp()
        try:
            ZipFile(pptxfile).extractall(path=tmpd, pwd=None)
        except:  # noqa
            flash(f'Cannot extract data from specified PowerPoint file {pptxfile}: f{sys.exc_info()}. Exiting.')

        # Parse slide content first
        path = tmpd + os.sep + 'ppt' + os.sep + 'slides' + os.sep
        for infile in glob.glob(os.path.join(path, '*.xml')):
            # parse each XML notes file from the notes folder.
            slidenum = int(re.sub(r'\D', '', infile.split(os.sep)[-1]))
            dom = parse(infile)

            # In slides, content is grouped by paragraph using <a:p>
            # Within the paragraph, there are multiple text blocks denoted as <a:t>
            # For each paragraph, concatenate all of the text blocks without whitespace,
            # then concatenate each paragraph delimited by a space.
            paragraphs = dom.getElementsByTagName('a:p')
            for paragraph in paragraphs:
                paragraphtext = ''
                parse_node(paragraph)
                urlmatches = re.finditer(urlmatchre, paragraphtext)

                for urlmatch in urlmatches:  # Now it's a tuple
                    url = urlmatch.group('url')

                    # Remove regex match artifacts at the end of the URL: www.sans.org,
                    url = trim_url(url)

                    # Add default URI for www.anything
                    if url[0:3] == 'www':
                        url = 'http://' + url

                    # Skip private IP addresses and localhost
                    privateaddr = re.compile(
                        r'(\S+127\.)|(\S+169\.254\.)|(\S+192\.168\.)|(\S+10\.)|(\S+172\.1[6-9]\.)|(\S+172\.2[0-9]\.)|(\S+172\.3[0-1]\.)|(\S+::1)')  # noqa
                    if re.match(privateaddr, url):
                        continue

                    if '://localhost' in url:
                        continue

                    # Skip .onion and .i2p domains
                    anondomain = re.compile(
                        r'\.onion$|\.onion\/|\.i2p$|\.i2p\/')
                    if re.match(anondomain, url):
                        continue

                    url = url.encode('ascii', 'ignore').decode('utf-8')

                    # Add this URL to the hash
                    if url not in urls:
                        urls[url] = [filenum, slidenum]

        # Process notes content in slides
        path = tmpd + os.sep + 'ppt' + os.sep + 'notesSlides' + os.sep
        for infile in glob.glob(os.path.join(path, '*.xml')):
            # parse each XML notes file from the notes folder.

            # Get the slide number
            slidenum = int(re.search('notesSlide(\d+)\.xml', infile).group(1))  # noqa

            # Parse slide notes, adding a space after each paragraph marker, and
            # removing XML markup
            dom = parse(infile)
            paragraphs = dom.getElementsByTagName('a:p')
            for paragraph in paragraphs:
                paragraphtext = ''
                parse_node(paragraph)

                # Parse URL content from notes text for the current paragraph
                urlmatches = re.finditer(urlmatchre, paragraphtext)

                for urlmatch in urlmatches:  # Now it's a tuple
                    url = urlmatch.group('url')
                    # Remove regex match artifacts at the end of the URL: www.sans.org,
                    url = trim_url(url)

                    # Add default URI for www.anything
                    if url[0:3] == 'www':
                        url = 'http://' + url

                    # Remove a trailing period
                    if url[-1] == '.':
                        url = url[:-1]

                    # Some authors include URLs in the form
                    # http://www.josh.net.[1], http://www.josh.net[1]. or
                    # http://www.josh.net[1] Remove the footnote and/or leading
                    # or trailing dot (this really only happens in notes)
                    footnote = re.compile(r'(\.\[\d+\]|\[\d+\]\.|\[\d+\])')
                    if re.search(footnote, url):
                        url = re.sub(footnote, '', url)

                    # Skip private IP addresses and localhost
                    privateaddr = re.compile(
                        r'(\S+127\.)|(\S+169\.254\.)|(\S+192\.168\.)|(\S+10\.)|(\S+172\.1[6-9]\.)|(\S+172\.2[0-9]\.)|(\S+172\.3[0-1]\.)|(\S+::1)')  # noqa
                    if re.match(privateaddr, url):
                        continue

                    if '://localhost' in url:
                        continue

                    # Skip .onion and .i2p domains
                    anondomain = re.compile(
                        r'\.onion$|\.onion\/|\.i2p$|\.i2p\/')
                    if re.findall(anondomain, url):
                        continue

                    url = url.encode('ascii', 'ignore').decode('utf-8')

                    # Add this URL to the hash
                    if url not in urls:
                        urls[url] = [filenum, slidenum]

        # Remove all the files created with unzip
        shutil.rmtree(tmpd)

    return urls


def signal_exit(signal, frame):
    sys.exit(0)


def test_url(url, filenum, pagenum):
    """
    Acccepts a URL, filenun, and page num as input.
    Test the given URL and return a list of
    [filenum, pagenum, url, HTTP response code, string/note]
    """

    ua = ('Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, '
          'like Gecko) Chrome/134.0.0.0 Safari/537.36')
    code = 'ERR'  # Default unless valid response
    note = ''     # Default no note

    headers = {
        'User-Agent': ua,
        }

    try:
        r = requests.head(url, timeout=TIMEOUT, verify=False, headers=headers)
        code = r.status_code
    except:  # noqa
        pass

    if (code == 200):
        return [filenum, pagenum, url, code, note]
    elif (code == 404):
        note = 'The URL returned a 404 File Not Found response (no such page on the server).'
        return [filenum, pagenum, url, code, note]

    # If we get anything other than a 200 or 404, rule out that the server is
    # returning a bad HTTP resposne for HEAD by trying a GET request.
    # Add the Range header to avoid downloading the entire file.
    headers['Range'] = 'bytes=0-0'

    try:
        r = requests.get(url, timeout=TIMEOUT, verify=False, headers=headers)
        code = r.status_code
    except requests.exceptions.HTTPError:
        note = 'An unspecified HTTP error occurred.'
    except requests.exceptions.ConnectionError:
        note = 'A connection error occurred (possible bad hostname).'
    except requests.exceptions.ConnectTimeout:
        note = ('A timeout error occurred creating a connection to the server '
                '(possible slow server or slow internet connection).')
    except requests.exceptions.ReadTimeout:
        note = 'A timeout error occurred when waiting for a read response from the server.'
    except requests.exceptions.InvalidURL:
        note = 'The URL is not valid.'
    except requests.exceptions.URLRequired:
        note = 'The URL is not valid.'
    except requests.exceptions.TooManyRedirects:
        note = 'A connection error occurred from too many server redirects.'
    except Exception:
        note = f'Unrecognized error accessing URL: {sys.exc_info()[1]}'

    if (code == 200):
        return [filenum, pagenum, url, code, note]
    elif (code == 206):
        # We're going to cheat and treat a 206 (Partial Content) as a 200 (OK)
        code = 200
        return [filenum, pagenum, url, code, note]
    elif (code == 404):
        note = 'The URL returned a 404 File Not Found response (no such page on the server).'
        return [filenum, pagenum, url, code, note]
    elif (code == 403):
        note = 'The URL returned a 403 Forbidden error (the server refuses to authorize the URL request).'
    elif (code == 400):
        note = (
            'The URL returned a 400 Bad Request error (the URL may not be intended to be visited '
            'using a standard browser).'
        )

    return [filenum, pagenum, url, code, note]


def allowed_file(filename):
    """
    Check if the file has a valid extension
    """
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def process_pptx_files(file_paths):
    """
    Accepts a list of PPTX file paths as input.
    Returns a single TXT output as a string.
    """

    # Build dictionary of URLs
    urls = parse_pptx(file_paths)

    # Check URLs with multiple threads
    urlchkres = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=CONNECTIONS) as executor:
        # The urls data object is url:[filenum, pagenum]
        futureurl = (executor.submit(
            test_url, url, urls[url][0], urls[url][1]) for url in urls)
        for future in concurrent.futures.as_completed(futureurl):
            try:
                data = future.result()
            except Exception:
                data = [sys.exc_info()]
            finally:
                urlchkres.append(data)

    # Sort list by file num, page num
    urlchkres = sorted(urlchkres, key=lambda x: (x[0], x[1]))
    return make_csv(urlchkres)


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'files[]' not in request.files:
            flash('No file part in the request.')
            return redirect(request.url)

        files = request.files.getlist('files[]')
        if not files or files[0].filename == '':
            flash('No file selected for uploading.')
            return redirect(request.url)

        # Save the uploaded files to a temporary directory
        temp_dir = tempfile.mkdtemp()
        saved_files = []
        try:
            for file in files:
                if file and allowed_file(file.filename):
                    filename = secure_filename(file.filename)
                    file_path = os.path.join(temp_dir, filename)
                    file.save(file_path)
                    saved_files.append(file_path)
                else:
                    flash(f"File {file.filename} is not a valid PPTX file.")
                    return redirect(request.url)

        except Exception as e:
            flash(f"An error occurred during processing: {str(e)}")
            return redirect(request.url)

        output_text = process_pptx_files(saved_files)

        output_io = io.BytesIO(output_text.encode('utf-8'))
        output_io.seek(0)

        # Clean up the temporary files
        for file_path in saved_files:
            os.remove(file_path)
        os.rmdir(temp_dir)

        return send_file(output_io, as_attachment=True, download_name="output.txt", mimetype='text/plain')

    else:
        # Handle GET request (display the form)
        return render_template_string(html_template)


html_template = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>PPTX URL Check</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/materialize/1.0.0/css/materialize.min.css">
    <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
    <style>
        body {
            display: flex;
            min-height: 100vh;
            flex-direction: column;
        }
        main {
            flex: 1 0 auto;
        }
        .container {
            margin-top: 50px;
        }
        .dropzone {
            border: 2px dashed #9E9E9E;
            border-radius: 5px;
            padding: 50px;
            text-align: center;
            cursor: pointer;
            margin-bottom: 20px;
        }
        .dropzone.hover {
            background-color: #EEEEEE;
        }
        #fileList {
            margin-top: 10px;
            font-style: italic;
        }
    </style>
</head>
<body>
    <nav>
      <div class="nav-wrapper teal">
        <a href="/" class="brand-logo center">PPTX URL Check</a>
      </div>
    </nav>
    <main>
      <div class="container">
        {% with messages = get_flashed_messages() %}
          {% if messages %}
            <ul class="collection red-text">
              {% for message in messages %}
                <li class="collection-item">{{ message }}</li>
              {% endfor %}
            </ul>
          {% endif %}
        {% endwith %}
        <div class="row">
          <form class="col s12" method="POST" enctype="multipart/form-data">
            <!-- Hidden file input element -->
            <input type="file" name="files[]" id="fileInput" multiple accept=".pptx" style="display: none;">
            <div class="center">
                <h5 id="pleaseWait">&nbsp;</h5>
            </div>
            <!-- Single drag-and-drop box -->
            <div class="dropzone" id="dropzone">
                Drag and drop PPTX files here (or click to select)
            </div>
            <!-- List of selected file names -->
            <div id="fileList"></div>
            <div class="center">
              <button class="btn waves-effect waves-light teal" type="submit" name="action" id="submitBtn">Submit
              </button>
            </div>
          </form>
        </div>
      </div>
    </main>
    <!-- Materialize JavaScript -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/materialize/1.0.0/js/materialize.min.js"></script>
    <script>
        var dropzone = document.getElementById('dropzone');
        var fileInput = document.getElementById('fileInput');
        var fileList = document.getElementById('fileList');

        // Open file picker when clicking the dropzone
        dropzone.addEventListener('click', function() {
            fileInput.click();
        });

        // Drag and drop events
        dropzone.addEventListener('dragover', function(e) {
            e.preventDefault();
            dropzone.classList.add('hover');
        });
        dropzone.addEventListener('dragleave', function(e) {
            e.preventDefault();
            dropzone.classList.remove('hover');
        });
        dropzone.addEventListener('drop', function(e) {
            e.preventDefault();
            dropzone.classList.remove('hover');
            const dt = e.dataTransfer;
            let files = dt.files;
            fileInput.files = files;
            updateFileList();
        });

        // When files are selected via file picker
        fileInput.addEventListener('change', updateFileList);

        function updateFileList() {
            var files = fileInput.files;
            var names = [];
            for (var i = 0; i < files.length; i++) {
                names.push(files[i].name);
            }
            fileList.innerHTML = names.length ? "Selected files:<br>" + names.join('<br>') : "";
        }

        // Show "Please wait" message on form submission
        var form = document.querySelector('form');
        form.addEventListener('submit', function(e) {
            document.getElementById('pleaseWait').innerHTML = 'Please wait...';
        });
    </script>
</body>
</html>
"""

if __name__ == '__main__':
    app.run(debug=True)
