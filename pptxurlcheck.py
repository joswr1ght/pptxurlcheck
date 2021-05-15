#!/usr/bin/env python3
# -*- coding: utf-8 -*-
#
# With code by Eric Jang ericjang2004@gmail.com
SKIP200=1

import sys
import re
import os
import shutil
import tempfile
import glob
import csv
import concurrent.futures
import requests
#import pdb
import signal
from zipfile import ZipFile
from xml.dom.minidom import parse

from requests.packages.urllib3.exceptions import InsecureRequestWarning
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

MAXREDIRECT=10
TIMEOUT=10
CONNECTIONS=20

# Remove trailing unwanted characters from the end of URL's
# This is a recursive function. Did I do it well? I don't know.
def striptrailingchar(s):
    # The valid URL charset is A-Za-z0-9-._~:/?#[]@!$&'()*+,;= and & followed by hex character
    # I don't have a better way to parse URL's from the cruft that I get from XML content, so I
    # also remove .),;'? too.  Note that this is only the end of the URL (making ? OK to remove)
    if s[-1] not in "ABCDEFGHIJKLMNOPQRSTUVWXYZZabcdefghijklmnopqrstuvwxyzz0123456789-_~#[]@!$&(*+=/":
        s = striptrailingchar(s[0:-1])
    elif s[-5:] == "&quot":
        s = striptrailingchar(s[0:-5])
    else:
        pass
    return s


# Parse the given root recursively (root is intended to be the paragraph element <a:p>
# If we encounter a link-break element a:br, add a new line to global paragraphtext
# If we encounter an element with type TEXT_NODE, append value to paragraphtext
paragraphtext=""
def parse_node(root):
    global paragraphtext
    if root.childNodes:
        for node in root.childNodes:
            if node.nodeType == node.TEXT_NODE:
                paragraphtext += node.nodeValue.encode('ascii', 'ignore').decode('utf-8')
            if node.nodeType == node.ELEMENT_NODE:
                if node.tagName == 'a:br':
                    paragraphtext += "\n" 
                parse_node(node)


# Accepts a list of PowerPoint files in pptxfiles.
# Returns a hash of links indexed by URL with [filenum,pagenum] as the value.
# Reads slide notes and slide text boxes and other text elements.
def parsepptx(pptxfiles):
    global paragraphtext
    urlmatchre = re.compile(r'((https?://[^\s<>"]+|www\.[^\s<>"]+))',re.DOTALL)
    urls = {}
    filenum=0

    for pptxfile in pptxfiles:
        filenum+=1

        tmpd = tempfile.mkdtemp()
        try:
            ZipFile(pptxfile).extractall(path=tmpd, pwd=None)
        except Exception as e:
            printerrex(f"Cannot extract data from specified PowerPoint file {pptxfile}: f{sys.exc_info()}. Exiting.")

        # Parse slide content first
        path = tmpd + os.sep + 'ppt' + os.sep + 'slides' + os.sep
        for infile in glob.glob(os.path.join(path, '*.xml')):
            #parse each XML notes file from the notes folder.
            slideText = ''
            slidenum = int(re.sub(r'\D', "", infile.split(os.sep)[-1]))
            dom = parse(infile)

            # In slides, content is grouped by paragraph using <a:p>
            # Within the paragraph, there are multiple text blocks denoted as <a:t>
            # For each paragraph, concatenate all of the text blocks without whitespace,
            # then concatenate each paragraph delimited by a space.
            paragraphs = dom.getElementsByTagName('a:p')
            for paragraph in paragraphs:
                paragraphtext = ""
                parse_node(paragraph)
                urlmatches = re.findall(urlmatchre, paragraphtext)

                for urlmatch in urlmatches:  # Now it's a tuple
                    # Remove regex match artifacts at the end of the URL: www.sans.org,
                    url = striptrailingchar(urlmatch[0])

                    # Add default URI for www.anything
                    if url[0:3] == "www":
                        url = "http://" + url

                    # Remove a trailing period
                    if url[-1] == ".":
                        url = url[:-1]

                    # Skip private IP addresses and localhost
                    privateaddr = re.compile(r'(\S+127\.)|(\S+169\.254\.)|(\S+192\.168\.)|(\S+10\.)|(\S+172\.1[6-9]\.)|(\S+172\.2[0-9]\.)|(\S+172\.3[0-1]\.)|(\S+::1)')
                    if re.match(privateaddr, url):
                        continue

                    if "://localhost" in url:
                        continue

                    # Skip .onion and .i2p domains
                    anondomain = re.compile(r'\.onion$|\.onion\/|\.i2p$|\.i2p\/')
                    if re.match(anondomain, url):
                        continue

                    url = url.encode('ascii', 'ignore').decode('utf-8')

                    # Add this URL to the hash
                    if not url in urls:
                        urls[url] = [filenum, slidenum]


        # Process notes content in slides
        path = tmpd + os.sep + 'ppt' + os.sep + 'notesSlides' + os.sep
        for infile in glob.glob(os.path.join(path, '*.xml')):
            # parse each XML notes file from the notes folder.

            # Get the slide number
            slidenum = int(re.search("notesSlide(\d+)\.xml", infile).group(1))

            # Parse slide notes, adding a space after each paragraph marker, and
            # removing XML markup
            dom = parse(infile)
            paragraphs = dom.getElementsByTagName('a:p')
            for paragraph in paragraphs:
                paragraphtext = ""
                parse_node(paragraph)

                # Parse URL content from notes text for the current paragraph
                urlmatches = re.findall(urlmatchre, paragraphtext)

                for urlmatch in urlmatches:  # Now it's a tuple
                    # Remove regex match artifacts at the end of the URL: www.sans.org,
                    url = striptrailingchar(urlmatch[0])

                    # Add default URI for www.anything
                    if url[0:3] == "www":
                        url = "http://" + url

                    # Remove a trailing period
                    if url[-1] == ".":
                        url = url[:-1]

                    # Some authors include URLs in the form
                    # http://www.josh.net.[1], http://www.josh.net[1]. or
                    # http://www.josh.net[1] Remove the footnote and/or leading
                    # or trailing dot (this really only happens in notes)
                    footnote=re.compile(r"(\.\[\d+\]|\[\d+\]\.|\[\d+\])")
                    if re.search(footnote, url):
                        url=re.sub(footnote, "", url)

                    # Skip private IP addresses and localhost
                    privateaddr = re.compile(r'(\S+127\.)|(\S+169\.254\.)|(\S+192\.168\.)|(\S+10\.)|(\S+172\.1[6-9]\.)|(\S+172\.2[0-9]\.)|(\S+172\.3[0-1]\.)|(\S+::1)')
                    if re.match(privateaddr, url):
                        continue

                    if "://localhost" in url:
                        continue

                    # Skip .onion and .i2p domains
                    anondomain = re.compile(r'\.onion$|\.onion\/|\.i2p$|\.i2p\/')
                    if re.findall(anondomain, url):
                        continue

                    url = url.encode('ascii', 'ignore').decode('utf-8')

                    # Add this URL to the hash
                    if not url in urls:
                        urls[url] = [filenum, slidenum]

        # Remove all the files created with unzip
        shutil.rmtree(tmpd)

    return urls

def signal_exit(signal, frame):
    sys.exit(0)

# Acccepts a URL, filenun, and page num as input
# Returns a list of [filenum, pagenum, url, HTTP response code, string/note]
def testurl(url, filenum, pagenum):
    code="ERR" # Default unless valid response
    note="" # Default no note

    headers = { 'User-Agent' : 'Mozilla/5.0 (Macintosh; Intel Mac OS X 11_2_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.114 Safari/537.36' }
    try:
        r=requests.get(url, timeout=TIMEOUT, verify=False, headers=headers)
        code=r.status_code
    except requests.exceptions.HTTPError as e:
        note="An unspecified HTTP error occurred."
    except requests.exceptions.ConnectionError as e:
        note="A connection error occurred (possible bad hostname)."
    except requests.exceptions.ConnectTimeout as e:
        note="A timeout error occurred creating a connection to the server (possible slow server or slow internet connection)."
    except requests.exceptions.ReadTimeout as e:
        note="A timeout error occurred when waiting for a read response from the server."
    except requests.exceptions.InvalidURL as e:
        note="The URL is not valid."
    except requests.exceptions.URLRequired as e:
        note="The URL is not valid."
    except requests.exceptions.TooManyRedirects as e:
        note="A connection error occurred from too many server redirects."
    except Exception as e:
        note=f"Unrecognized error accessing URL: {sys.exc_info()[1]}"

    if (code == 404):
        note="The URL returned a 404 File Not Found response (no such page on the server)."
    elif (code == 403):
        note="The URL returned a 403 Forbidden error (the server refuses to authorize the URL request)."
    elif (code == 400):
        note="The URL returned a 400 Bad Request error (the URL may not be intended to visited using a standard browser)."

    return [filenum, pagenum, url, code, note]

def printerrex(msg):
    sys.stdout.write(msg + "\n")

    if os.name == 'nt':
        x=input("Press Enter to exit.")

    sys.exit(-1)


if __name__ == "__main__":

    urlignore = [] # List of URLs to ignore, may be empty

    if (len(sys.argv) == 1):
        print("Validate URLs in the notes and slides of one or more PowerPoint pptx files. (version 2.1)")
        print("Check GitHub for updates: http://github.com/joswr1ght/pptxurlcheck\n")
        if os.name == 'nt':
            print("Usage: pptxurlcheck.exe [pptx file(s)]")
        else:
            print("Usage: pptxurlcheck.py [pptx file(s)]")
        sys.exit(1)

    signal.signal(signal.SIGINT, signal_exit)

    # Check each file supplied to make sure it has the .pptx extension
    for filename in sys.argv[1:]:
        if (os.path.splitext(filename)[1] == ".txt"):
            # Treat this file as a list of URLs to ignore
            with open(filename, 'r') as urllistfd:
                urlignore = [line.strip() for line in urllistfd.readlines()]
            sys.argv.remove(filename) # Remove .txt file from array
        elif (os.path.splitext(filename)[1] != ".pptx"):
            printerrex(f"Invalid PPTX file supplied: {filename}")

    # Build dictionary of URLs
    urls = parsepptx(sys.argv[1:])

    # Remove URLs in exclusion list
    for url in urlignore:
        if url in urls:
            urls.pop(url)

    urlchkres = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=CONNECTIONS) as executor:
        # The urls data object is url:[filenum, pagenum]
        futureurl = (executor.submit(testurl, url, urls[url][0], urls[url][1]) for url in urls)
        for future in concurrent.futures.as_completed(futureurl):
            try:
                data = future.result()
            except Exception as exc:
                data = [sys.exc_info()]
            finally:
                urlchkres.append(data)
                print(str(len(urlchkres)),end="\r")

    #print(urlchkres)

    # Sort list by file num, page num
    urlchkres=sorted(urlchkres, key=lambda x: (x[0], x[1]))

    # If SKIP200=0 in the env then we create the CSV with all URLs, not just problem URLs
    skip200=int(os.getenv('SKIP200', 1))

    # Get the report filename for the CSV, placing it in the dir where the PPTX files are
    basedir,fpptxfile = os.path.split(sys.argv[1])
    if basedir == '': # Relative directory
        reportfilename="pptxurlreport.csv"
    else:
        reportfilename=f"{basedir + os.sep}pptxurlreport.csv"

    # Open and generate the CSV report
    with open(reportfilename, mode='w') as csv_report:
        csvwriter = csv.writer(csv_report, lineterminator='\n')
        csvwriter.writerow(["File#","Page","Response","URL","Note"])
        # Loop through results to make CSV report
        for urldata in urlchkres:
            filenum=urldata[0]
            pagenum=urldata[1]
            url=urldata[2]
            response=urldata[3]
            note=urldata[4]
            if (skip200==1 and response==200):
                continue
            csvwriter.writerow([filenum,pagenum,response,url,note])

    print(f"URL validation report created at {reportfilename}")

    if os.name == 'nt':
        x=input("Press Enter to exit.")
