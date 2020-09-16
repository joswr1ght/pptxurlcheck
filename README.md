# pptxsanity

Parse a PowerPoint PPTX file, extracting all URLs from notes and slides, and
test for validity returning ERR or the non-OK HTTP status code.

## Usage

```
$ python pptxsanity.py
Validate URLs in the notes and slides of a PowerPoint pptx file. (version 1.2)
Check GitHub for updates: http://github.com/joswr1ght/pptxsanity

Usage: pptxsanity.py [pptx file]
$ python pptxsanity.py SEC561.pptx
$ ./pptxsanity.py ~/Dropbox\ \(SANS\)/SEC504/SEC504-F02/SEC504_2_F01_01.pptx
ERR : http://www.[target_company].com, Page 41
404 : http://bit.ly/14GZzcT, Page 87
```

Pptxsanity searches all slide bullets and notes pages for URLs, and attempts to retrieve the URL.
By default, URLs that are valid (e.g. that return a 200 OK message) are not displayed; all other URLs are displayed along
with the return code.  `ERR` indicates that the server could not be reached.  If you want to see each URL that is tested,
set the environment variable `SKIP200` to `0`:

```
$ SKIP200=0 python pptxsanity.py SEC561.pptx
200 : http://w3af.org
200 : http://blogs.msdn.com/b/tzink/archive/2012/08/29/how-rainbow-tables-work.aspx
ERR : http://www.ipbackupanalyzer.com
200 : http://www.cclgroupltd.com/Buy-Software/other-software-a-scripts.html
```

Windows users will set an environment variable before running the command:
```
C:\>set SKIP200=0
C:\>pptxsanity SEC561.pptx
200 : http://w3af.org
200 : http://blogs.msdn.com/b/tzink/archive/2012/08/29/how-rainbow-tables-work.aspx
ERR : http://www.ipbackupanalyzer.com
200 : http://www.cclgroupltd.com/Buy-Software/other-software-a-scripts.html
```

## Platforms

Tested on Windows 10, OS X 10.14, and Debian-based Linux. Windows binary
included in the `bin/` directory, built with `C:\Python38\scripts\pyinstaller
--onefile pptxsanity.py`.

## Questions, Comments, Concerns?

Open a ticket, or drop me a note: jwright@hasborg.com.

-Josh
