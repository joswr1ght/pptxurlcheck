#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
# With code by Eric Jang ericjang2004@gmail.com
TIMEOUT=6 # URL request timeout in seconds
SKIP200=1

from pptx import Presentation
import sys
import re
import os
import urlparse
import shutil
import glob
import tempfile
import urllib2
import signal
from zipfile import ZipFile
from xml.dom.minidom import parse
import platform
from selenium import webdriver
from PIL import Image
from PIL import ImageFont
from PIL import ImageDraw
import traceback


def unique_file(file_name):
    counter = 1
    file_name_parts = os.path.splitext(file_name) # returns ('/path/file', '.ext')
    while 1:
        if (not os.path.isfile(file_name)):
            return file_name
        else:
            file_name = file_name_parts[0] + '_' + str(counter) + file_name_parts[1]
            counter += 1


# Remove trailing unwanted characters from the end of URL's
# This is a recursive function. Did I do it well? I don't know.
def striptrailingchar(s):
    # The valid URL charset is A-Za-z0-9-._~:/?#[]@!$&'()*+,;= and & followed by hex character
    # I don't have a better way to parse URL's from the cruft that I get from XML content, so I
    # also remove .),;'? too.  Note that this is only the end of the URL (making ? OK to remove)
    if s[-1] not in "ABCDEFGHIJKLMNOPQRSTUVWXYZZabcdefghijklmnopqrstuvwxyzz0123456789-_~:/#[]@!$&(*+=":
        s = striptrailingchar(s[0:-1])
    elif s[-5:] == "&quot":
        s = striptrailingchar(s[0:-5])
    else:
        pass
    return s

def parseslidenotes(pptxfile):
    urls = []
    tmpd = tempfile.mkdtemp()

    ZipFile(pptxfile).extractall(path=tmpd, pwd=None)
    path = tmpd + '/ppt/notesSlides/'

    for infile in glob.glob(os.path.join(path, '*.xml')):
        #parse each XML notes file from the notes folder.
        dom = parse(infile)
        noteslist = dom.getElementsByTagName('a:t')
        #separate last element of noteslist for use as the slide marking.
        slideNumber = noteslist.pop()
        slideNumber = slideNumber.toxml().replace('<a:t>', '').replace('</a:t>', '')
        #start with this empty string to build the presenter note itself
        text = ''

        for node in noteslist:
            xmlTag = node.toxml()
            xmlData = xmlTag.replace('<a:t>', '').replace('</a:t>', '')
            #concatenate the xmlData to the text for the particular slideNumber index.
            text += " " + xmlData

        # Convert to ascii to simplify
        text = text.encode('ascii', 'ignore')
        urlmatches = re.findall(urlmatchre,text)
	if len(urlmatches) > 0:
            for match in urlmatches: # Now it's a tuple
                 for urlmatch in match:
                      if urlmatch != '':
                          urls.append(striptrailingchar(urlmatch))

    # Remove all the files created with unzip
    shutil.rmtree(tmpd)
    return urls

# Parse the text on slides using the python-pptx module, return URLs
def parseslidetext(prs):
    urls = []
    nexttitle = False
    for slide in prs.slides:
        text_runs = []
        for shape in slide.shapes:
            try:
                if not shape.has_text_frame:
                    continue
            except AttributeError:
                sys.stderr.write("Error: Please upgrade your version of python-pptx: pip uninstall python-pptx ; pip install python-pptx\n")
                sys.exit(-1)
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    text_runs.append(run.text)

            for text in text_runs:
                if text == None : continue
                try:
                    m = re.match(urlmatchre,text)
                except IndexError,TypeError:
                    continue
                if m != None:
                    url = striptrailingchar(m.groups()[0])
                    if url not in urls:
                        urls.append(url)
    return urls

def signal_exit(signal, frame):
    sys.exit(0)

if __name__ == "__main__":
    opt_pptxfile = None
    opt_renderdir = None

    if (len(sys.argv) != 2 and len(sys.argv) != 4):
        print "Validate URLs in the notes and slides of a PowerPoint pptx file."
        print "Check GitHub for updates: http://github.com/joswr1ght/pptxsanity\n"
        # There are instructions on how to write the Usage output:
        # http://courses.cms.caltech.edu/cs11/material/general/usage.html
        if (platform.system() == 'Windows'):
            print "Usage: pptxsanity.exe pptxfile"
            print "  pptxfile: The PowerPoint PPTX file"
        else:
            print "Usage: pptxsanity.py [-r dir] pptxfile"
            print "  pptxfile: The PowerPoint PPTX file"
            print "  -r dir  : Render each retrieved web page as a PNG in the specified directory"
        sys.exit(1)

    signal.signal(signal.SIGINT, signal_exit)

    if (len(sys.argv) == 2):
        opt_pptxfile = sys.argv[1]
    elif sys.argv[1] == '-r':
        opt_renderdir=sys.argv[2]
        opt_pptxfile = sys.argv[3]
        if (not os.path.exists(opt_renderdir)):
            sys.stderr.write("Output directory " + opt_renderdir + " does not exist.\n")
            sys.exit(-1)
    try:
        prs = Presentation(opt_pptxfile)
    except Exception:
        sys.stderr.write("Invalid PPTX file: " + opt_pptxfile + "\n")
        sys.exit(-1)
    
    # This may be the most insane regex I've ever seen.  It's very comprehensive, but it's too aggressive for
    # what I want.  It matches arp:remote in ettercap -TqM arp:remote // //, so I'm using something simpler
    #urlmatchre = re.compile(r"""((?:[a-z][\w-]+:(?:/{1,3}|[a-z0-9%])|www\d{0,3}[.]|[a-z0-9.\-]+[.‌​][a-z]{2,4}/)(?:[^\s()<>]+|(([^\s()<>]+|(([^\s()<>]+)))*))+(?:(([^\s()<>]+|(‌​([^\s()<>]+)))*)|[^\s`!()[]{};:'".,<>?«»“”‘’]))""", re.DOTALL)
    urlmatchre = re.compile(r'((https?://[^\s<>"]+|www\.[^\s<>"]+))',re.DOTALL)
    privateaddr = re.compile(r'(\S+127\.)|(\S+192\.168\.)|(\S+10\.)|(\S+172\.1[6-9]\.)|(\S+172\.2[0-9]\.)|(\S+172\.3[0-1]\.)|(\S+::1)')

    SKIP200=int(os.getenv('SKIP200', 1))
    
    urls = []
    urls += parseslidetext(prs)
    urls += parseslidenotes(opt_pptxfile)

    # De-duplicate URL's
    urls = list(set(urls))

    for url in urls:
        url = url.encode('ascii', 'ignore')

        # Add default URI for www.anything
        if url[0:3] == "www": url="http://"+url

        # Skip private IP addresses
        if re.match(privateaddr,url): continue

        headers = { 'User-Agent' : 'Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:35.0) Gecko/20100101 Firefox/35.0' }
        try:
            #ul=urllib2.urlopen(url, timeout=TIMEOUT)
            req=urllib2.Request(url, None, headers)
            ul=urllib2.urlopen(req, timeout=TIMEOUT)
            code=ul.getcode()
            if opt_renderdir:
                try:
                    driver = webdriver.PhantomJS()
                    driver.set_window_size(1024,800)
                    driver.get(url)
                    imagefile=unique_file(opt_renderdir + "/" + urlparse.urlparse(url).netloc + ".png")
                    driver.get_screenshot_as_file(imagefile)
                    img=Image.open(imagefile)
                    draw=ImageDraw.Draw(img)
                    raw=ImageDraw.Draw(img)
                    font=ImageFont.truetype("MesloLGLDZ-Regular.ttf",12)
                    draw.text((1,1), url,(255,255,255),font=font)
                    draw.text((0,0), url,(0,0,0),font=font)
                    img2 = img.crop((0,0,1024,800))
                    img2.save(imagefile)
                except Exception, err:
                    print(traceback.format_exc())
                    #sys.stderr.write("Cannot render content for URL " + url + "\n")

            if code == 200 and SKIP200 == 1:
                continue
            print str(code) + " : " + url

        except Exception, e:
            try:
                print str(e.code) + " : " + url
            except Exception:
                print "ERR : " + url
