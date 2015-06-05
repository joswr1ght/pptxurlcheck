# pptxsanity
Parse a PowerPoint PPTX file, extracting all URL's from notes and slides, and test for validity

## Usage
```
$ python pptxsanity.py
Validate URLs in the notes and slides of a PowerPoint pptx file.
Check GitHub for updates: http://github.com/joswr1ght/pptxsanity

Usage: pptxsanity.py [pptx file]
$ python pptxsanity.py SEC561.pptx
ERR : http://www.ipbackupanalyzer.com
403 : http://java.decompiler.free.fr/?q=jdgui
ERR : http://host/u.php?id=u1
404 : http://securityweekly.com/2011/11/safely-dumping-hashes-from-liv.html
ERR : https://localhost:8834
404 : http://www.willhackforsushi.com/ios-key-recovery.pdf
```

Pptxsanity searches all slide bullets and notes pages for URL's, and attempts to retrieve the URL.
By default, URL's that are valid (e.g. that return a 200 OK message) are not displayed; all other URL's are displayed along
with the return code.  `ERR` indicates that the server could not be reached.  If you want to see each URL that is tested,
set the environment variable `SKIP200` to `0`:

```
$ SKIP200=0 python pptxsanity.py SEC561.pptx
200 : http://w3af.org
200 : http://blogs.msdn.com/b/tzink/archive/2012/08/29/how-rainbow-tables-work.aspx
ERR : http://www.ipbackupanalyzer.com
200 : http://www.cclgroupltd.com/Buy-Software/other-software-a-scripts.html
```

Windows users will set an envionment variable before running the command:
```
C:\>set SKIP200=0
C:\>pptxsanity SEC561.pptx
200 : http://w3af.org
200 : http://blogs.msdn.com/b/tzink/archive/2012/08/29/how-rainbow-tables-work.aspx
ERR : http://www.ipbackupanalyzer.com
200 : http://www.cclgroupltd.com/Buy-Software/other-software-a-scripts.html
```

## Platforms

Tested on Windows 8.1, OS X 10.10, and Debian-based Linux.  Windows binary included in the `bin/` directory, built with `C:\Python27\scripts\pyinstaller --onefile bitfit.py`.

## Questions, Comments, Concerns?

Open a ticket, or drop me a note: jwright@hasborg.com.

-Josh

