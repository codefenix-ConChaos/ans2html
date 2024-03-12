# ANSI-To-HTML Converter (ans2html.vbs)

## Description:

   This VBScript is for converting an ANSI file to a HTML file.  Useful for
   displaying BBS door game scores on a website.

   Also optionally supports other common coloring schemes:
   - Pipes |
   - Tildes ~
   - RTSoft ` codes (LoRD, LoRD2, TEOS, and others)
   - Yankee Trader Galactic Newspaper

   It reads each character from a standard ANSI source file and generates
   a file containing HTML5 markup.  It interprets most ANSI escape codes
   and translates all 255 codepage 437 characters to the best matching
   equivalent HTML entity.
   See: https://en.wikipedia.org/wiki/Code_page_437

   After reading the ANSI source data, the script will first "flatten" it,
   eliminating all cursor movement sequences so that it need only convert
   the "m" escape sequences for in-line text coloring. This flattening does
   not occur if the using the pipe, tilde, RTSoft, or Yankee Trader 
   conversion modes.

   The IBM VGA font from the Ultimate Oldschool PC Font Pack is the optimal
   font to use for displaying CP437 characters in browsers. Download it from 
   https://int10h.org/oldschool-pc-fonts/download and set it up as a webfont
   on your site. If this font is not present, web browsers will default to 
   whatever default monospace font is configured, leading to mixed results 
   for box and line drawing characters, especially on mobile browsers.

   The "Source Code Pro" font is another good monospace font that gives nice
   results. Download it from https://github.com/adobe-fonts/source-code-pro.

   Blinking text is achieved using keyframes, setting the color:hsla property
   in CSS. Use either "linear" to "step-end" in the CSS animation properties
   for a gentle fade or sharp flash.

## Usage:

  `cscript ans2html.vbs path_to_ansi.ans path_to_html.html [page_title] [opts]`

The "opts" can be any or all processing modes:

 - P: Pipe codes
 - T: Tilde codes
 - L: RTSoft "LoRD" codes
 - Y: Yankee Trader Galactic Newspaper bulletin prefixes

You must specify a page title if using one of the optional processing modes.

Example:

  `cscript ans2html.vbs c:\lord\LOGNOW.TXT c:\web\lord_news.html "LorD News" L`



Probably goes without saying, but paths containing spaces must be wrapped
in double-quotes.


# Enjoy!
 I hope people find this script useful.
