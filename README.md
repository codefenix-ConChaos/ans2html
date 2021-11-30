**Description:**

  This VBScript is for converting an ANSI file to a HTML file.  Useful for
  displaying BBS door game scores on a website.

  It reads each character from a standard ANSI source file and generates
  a file containing HTML5 markup.  It interprets most ANSI escape codes
  and translates all 255 codepage 437 characters to the best matching
  equivalent HTML entity.
  See: https://en.wikipedia.org/wiki/Code_page_437

  After reading the ANSI source data, the script will first "flatten" it,
  eliminating all cursor movement sequences so that it need only convert 
  the "m" escape sequences for in-line text coloring.

  The "Source Code Pro" font is optional but highly recommended for best
  results. Download it from https://github.com/adobe-fonts/source-code-pro
  and install it as a web font for your site.  If this font is not present,
  web browsers will default to whatever monospace font is configured,
  leading to mixed results for box and line drawing characters,
  especially on mobile browsers.
  
**Example score bulletin output:**
  - Legend of the Red Dragon Player Rankings: https://conchaos.synchro.net/doors/lord_1scores.html
  - Operation Overkill Top 10: https://conchaos.synchro.net/doors/ooii-a_1scores.html
  - Rockin Radio Top 10: https://conchaos.synchro.net/doors/rradio_1scores.html

**Other examples:**
  - TradeWars 2002 Cineplex: https://conchaos.synchro.net/doors/CINEPLEX.html
  - TradeWars 2002 Derelict Spacecraft: https://conchaos.synchro.net/doors/ALN1.html

**Known issues:**

  Blinking text is achieved using CSS (built into the "htmlOutput" string)
  but it needs work to allow alternating between foreground and background
  colors.  Currently it just blinks them simultaneously.

**Usage:**

  `cscript ans2html.vbs path_to_ansi.ans path_to_html.html [page_title]`

Probably goes without saying, but paths containing spaces must be wrapped
in double-quotes.


**TODO:**
- Improve CSS blink.


