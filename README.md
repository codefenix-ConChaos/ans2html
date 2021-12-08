**Description:**

  This VBScript is for converting an ANSI file to a HTML file.  Useful for
  displaying BBS door game scores on a website.

  It reads each character from a standard ANSI source file, and generates
  a file containing HTML markup.  It interprets most ANSI escape codes
  and translates all 255 codepage 437 characters to the best matching
  equivalent HTML entity.
  See: https://en.wikipedia.org/wiki/Code_page_437

  After reading the ANSI source data, the script will first "flatten" it,
  eliminating all cursor movement sequences so that it then need only convert 
  the "m" escape sequences for in-line text coloring.

  The "Source Code Pro" font is optional but highly recommended for best
  results. Download it from https://github.com/adobe-fonts/source-code-pro
  and install it as a web font for your site.  If this font is not present,
  web browsers will default to whatever monospace font is configured,
  leading to mixed results for box and line drawing characters,
  especially on mobile browsers.

  Blinking text is achieved using keyframes, setting the color:hsla propery
  in CSS. The fade effect is deliberate, but can be replaced by a steady 
  blink by changing `linear` to `step-end` in the CSS animation properties.

  The `<head>` tags, `<body>` tags, and outer `<html>` tags are all 
  intentionally left out of the resulting HTML, since they're not needed 
  for my specific purposes. One could easily add them if wanted.

**Usage:**

  `cscript ans2html.vbs path_to_ansi.ans path_to_html.html [page_title]`

Probably goes without saying, but paths containing spaces must be wrapped
in double-quotes.
  
**Example score bulletin output:**
  - Legend of the Red Dragon Player Rankings: https://conchaos.synchro.net/doors/lord_1scores.html
  - Operation Overkill Top 10: https://conchaos.synchro.net/doors/ooii-a_1scores.html
  - Rockin Radio Top 10: https://conchaos.synchro.net/doors/rradio_1scores.html

**Other examples:**
  - TradeWars 2002 Cineplex (with blinking demo): https://conchaos.synchro.net/doors/CINEPLEX.html
  - TradeWars 2002 Derelict Spacecraft: https://conchaos.synchro.net/doors/ALN1.html

Enjoy!
