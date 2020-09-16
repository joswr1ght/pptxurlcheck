#!/usr/bin/env python3
# -*- coding: utf-8 -*-
#
# With code by Eric Jang ericjang2004@gmail.com
SKIP200=1

import sys
import re
import os
import shutil
import glob
import tempfile
import urllib3
#import pdb

# Attempt to add SNI support
try:
    import urllib3.contrib.pyopenssl
    urllib3.contrib.pyopenssl.inject_into_urllib3()
except ImportError:
    pass

import signal
from zipfile import ZipFile
from xml.dom.minidom import parse

# Remove trailing unwanted characters from the end of URL's
# This is a recursive function. Did I do it well? I don't know.
def striptrailingchar(s):
    # The valid URL charset is A-Za-z0-9-._~:/?#[]@!$&'()*+,;= and & followed by hex character
    # I don't have a better way to parse URL's from the cruft that I get from XML content, so I
    # also remove .),;'? too.  Note that this is only the end of the URL (making ? OK to remove)
    if s[-1] not in "ABCDEFGHIJKLMNOPQRSTUVWXYZZabcdefghijklmnopqrstuvwxyzz0123456789-_~:#[]@!$&(*+=/":
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


# Return a hash of links in the urls object indexed by page number
# Read from slide notes and slide text boxes and other text elements
def parseslidenotes(pptxfile):
    global paragraphtext

    # This may be the most insane regex I've ever seen.  It's very comprehensive, but it's too aggressive for
    # what I want.  It matches arp:remote in ettercap -TqM arp:remote // //, so I'm using something simpler
    #urlmatchre = re.compile(r"""((?:[a-z][\w-]+:(?:/{1,3}|[a-z0-9%])|www\d{0,3}[.]|[a-z0-9.\-]+[.‌​][a-z]{2,4}/)(?:[^\s()<>]+|(([^\s()<>]+|(([^\s()<>]+)))*))+(?:(([^\s()<>]+|(‌​([^\s()<>]+)))*)|[^\s`!()[]{};:'".,<>?«»“”‘’]))""", re.DOTALL)
    urlmatchre = re.compile(r'((https?://[^\s<>"]+|www\.[^\s<>"]+))',re.DOTALL)
    urls = {}

    tmpd = tempfile.mkdtemp()
    ZipFile(pptxfile).extractall(path=tmpd, pwd=None)

    # Parse slide content first
    path = tmpd + os.sep + 'ppt' + os.sep + 'slides' + os.sep
    for infile in glob.glob(os.path.join(path, '*.xml')):
        #parse each XML notes file from the notes folder.
        slideText = ''
        slideNumber = re.sub(r'\D', "", infile.split(os.sep)[-1])
        dom = parse(infile)

        # In slides, content is grouped by paragraph using <a:p>
        # Within the paragraph, there are multiple text blocks denoted as <a:t>
        # For each paragraph, concatenate all of the text blocks without whitespace,
        # then concatenate each paragraph delimited by a space.
        paragraphs = dom.getElementsByTagName('a:p')
        for paragraph in paragraphs:
            textblocks = paragraph.getElementsByTagName('a:t')
            for textblock in textblocks:
                slideText += textblock.toxml().replace('<a:t>','').replace('</a:t>','')
            slideText += " "

        # Parse URL content from notes text for the current paragraph
        urlmatches = re.findall(urlmatchre, slideText)
        for urlmatch in urlmatches:  # Now it's a tuple
            # Remove regex artifacts at the end of the URL: www.sans.org,
            url = striptrailingchar(urlmatch[0])

            # Add default URI for www.anything
            if url[0:3] == "www":
                url = "http://" + url

            # Add this URL to the hash
            slideNumber = int(slideNumber)
            if (slideNumber in urls):
                urls[slideNumber].append(url)
            else:
                urls[slideNumber] = [url]


    # Process notes content in slides
    path = tmpd + os.sep + 'ppt' + os.sep + 'notesSlides' + os.sep
    for infile in glob.glob(os.path.join(path, '*.xml')):
        # parse each XML notes file from the notes folder.

        # Get the slide number
        slideNumber = re.search("notesSlide(\d+)\.xml", infile).group(1)

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

                # Remove regex artifacts at the end of the URL: www.sans.org,
                url = striptrailingchar(urlmatch[0])

                # Add default URI for www.anything
                if url[0:3] == "www":
                    url = "http://" + url

                # Add this URL to the hash
                slideNumber = int(slideNumber)
                if (slideNumber in urls):
                    urls[slideNumber].append(url)
                else:
                    urls[slideNumber] = [url]

    # Remove all the files created with unzip
    shutil.rmtree(tmpd)
    return urls

def signal_exit(signal, frame):
    sys.exit(0)

if __name__ == "__main__":
    if (len(sys.argv) != 2):
        print("Validate URLs in the notes and slides of a PowerPoint pptx file. (version 1.2)")
        print("Check GitHub for updates: http://github.com/joswr1ght/pptxsanity\n")
        if os.name == 'nt':
            print("Usage: pptxsanity.exe [pptx file]")
        else:
            print("Usage: pptxsanity.py [pptx file]")
        sys.exit(1)

    signal.signal(signal.SIGINT, signal_exit)

    # Disable urllib3 InsecureRequestWarning
    try:
        urllib3.disable_warnings()
    except AttributeError:
        sys.stdout.write("You need to upgrade your version of the urllib3 library to the latest available.\n");
        sys.stdout.write("Try running the following command to upgrade urllib3: sudo pip install urllib3 --upgrade\n");
        sys.exit(1)


    SKIP200=int(os.getenv('SKIP200', 1))

    urls = parseslidenotes(sys.argv[1])

    # Deduplicate URLs on a single page (but not across pages)
    for key in urls:
        urls[key] = list(dict.fromkeys(urls[key]))

    for page in sorted(urls.keys()):
        for url in urls[page]:

            url = url.encode('ascii', 'ignore').decode('utf-8')

            # Add default URI for www.anything
            if url[0:3] == "www": url="http://"+url
    
            # Some authors include URLs in the form http://www.josh.net.[1], http://www.josh.net[1]. or http://www.josh.net[1] 
            # Remove the footnote and/or leading or trailing dot.
            footnote=re.compile(r"(\.\[\d+\]|\[\d+\]\.|\[\d+\])")
            if re.search(footnote, url):
                url=re.sub(footnote, "", url)

            # Remove a trailing period
            if url[-1] == ".":
                url = url[:-1]

            # Skip private IP addresses and localhost
            privateaddr = re.compile(r'(\S+127\.)|(\S+192\.168\.)|(\S+10\.)|(\S+172\.1[6-9]\.)|(\S+172\.2[0-9]\.)|(\S+172\.3[0-1]\.)|(\S+::1)')
            if re.match(privateaddr, url): continue
            if ("://localhost" in url): continue

            # Uncomment this debug line to print the URL before testing status to identify sites causing "Bus Error" fault on OSX
            #print "DEBUG: %s"%url
            headers = { 'User-Agent' : 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.102 Safari/537.36' }
            retries=urllib3.Retry(redirect=False, total=1, connect=0, read=0)
            http = urllib3.PoolManager(timeout=10, retries=retries)
            try:
                #req=http.request('HEAD', url, headers=headers)
                req=http.urlopen('GET', url, headers=headers, redirect=False)
                code=req.status
            except Exception as e:
                print(f"ERR : {url}, Page {page}")
                continue

            if (code == 302 or code == 301):
                # Do I still get a redirect if I add a trailing / ?
                redircode=None
                try:
                    req=http.request('GET', url + "/", headers=headers)
                    redircode=req.status
                    if (redircode == 200 and SKIP200 == 1):
                        # Adding a / to the end of the URL eliminated the redirect; skip this valid URL
                        continue
                except Exception as e:
                    print(f"ERR : {url}, Page {page}")
                    continue
            if code == 200 and SKIP200 == 1:
                continue
            print(f"{code} : {url}, Page {page}")

    if os.name == 'nt':
        x=input("Press Enter to exit.")
