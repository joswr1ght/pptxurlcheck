# pptxurlcheck

Parse a PowerPoint PPTX file, extracting all URLs from notes and slides, and
test for validity returning ERR or the non-OK HTTP status code.

## Usage

```
$ pptxurlcheck.py
Validate URLs in the notes and slides of one or more PowerPoint pptx files. (version 2.0)
Check GitHub for updates: http://github.com/joswr1ght/pptxurlcheck

Usage: pptxurlcheck.py [pptx file(s)]
$ pptxurlcheck.py SEC555/*.pptx
URL validation report created at SEC555/pptxurlreport.csv.
$ head -4 SEC555/pptxurlreport.csv
File#,Page,Response,URL,Note
1,5,ERR,https://intel.criticalstack.com,Maximum retry failure exceeded (possible bad server name)
1,5,ERR,https://sec555.com/4p,Maximum retry failure exceeded (possible bad server name)
2,54,404,http://schemas.microsoft.com/win/2004/08/events/event,
2,157,404,https://www.elastic.co/elasticon/2015/sf/scaling-elasticsearch-for-production-at-verizon,
```

Pptxurlcheck searches all slide bullets and notes pages for URLs, and attempts
to retrieve the URL. By default, URLs that are valid (e.g. that return a 200
OK message) are not displayed; all other URLs are displayed along with the
return code. `ERR` indicates that the server could not be reached. If you
want to see each URL that is tested, set the environment variable `SKIP200` to
`0`:

```
$ SKIP200=0 ~/Dev/pptxurlcheck/pptxurlcheck.py SEC555/SEC555_1_G01_01_JH.pptx
URL validation report created at SEC555/pptxurlreport.csv.
$ head -4 SEC555/pptxurlreport.csv
File#,Page,Response,URL,Note
1,7,200,https://content.fireeye.com/m-trends,
1,7,200,https://sec555.com/2g,
1,9,500,https://sec555.com/2i,
```

Windows users can set an environment variable before running the `set` command:

```
C:\>set SKIP200=0
C:\>pptxurlcheck SEC561.pptx
...
```

## Platforms

Tested on Windows 10, macOS 11.2, and Debian-based Linux. Windows binary
included in the `bin/` directory, built with `pyinstaller --onefile --hidden-import urllib3
pptxurlcheck.py` using Python 3.9.4.

## Questions, Comments, Concerns?

Open a ticket, or drop me a note: jwright@hasborg.com.
