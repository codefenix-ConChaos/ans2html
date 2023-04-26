'! ANSI-To-HTML Converter
'! (ans2html.vbs)
'! Author:           Craig Hendricks (aka Codefenix)
'! Initial Version:  3/15/2021
'! Updated:          4/26/2023
'! ============================================================================
'!
'! Description:
'!
'!   This VBScript is for converting an ANSI file to a HTML file.  Useful for
'!   displaying BBS door game scores on a website.
'!
'!   Also optionally supports other common coloring schemes:
'!   - Pipes |
'!   - Tildes ~
'!   - RTSoft ` codes (LoRD, LoRD2, TEOS, and others)
'!   - Yankee Trader Galactic Newspaper
'!
'!   It reads each character from a standard ANSI source file and generates
'!   a file containing HTML5 markup.  It interprets most ANSI escape codes
'!   and translates all 255 codepage 437 characters to the best matching
'!   equivalent HTML entity.
'!   See: https://en.wikipedia.org/wiki/Code_page_437
'!
'!   After reading the ANSI source data, the script will first "flatten" it,
'!   eliminating all cursor movement sequences so that it need only convert
'!   the "m" escape sequences for in-line text coloring. This flattening does
'!   not occur if the using the pipe, tilde, RTSoft, or Yankee Trader 
'!   conversion modes.
'!
'!   The IBM VGA font from the Ultimate Oldschool PC Font Pack is the optimal
'!   font to use for displaying CP437 characters in browsers. Download it from 
'!   https://int10h.org/oldschool-pc-fonts/download and set it up as a webfont
'!   on your site. If this font is not present, web browsers will default to 
'!   whatever default monospace font is configured, leading to mixed results 
'!   for box and line drawing characters, especially on mobile browsers.
'!
'!   The "Source Code Pro" font is another good monospace font that gives nice
'!   results. Download it from https://github.com/adobe-fonts/source-code-pro.
'!
'!   Blinking text is achieved using keyframes, setting the color:hsla property
'!   in CSS. Use either "linear" to "step-end" in the CSS animation properties
'!   for a gentle fade or sharp flash.
'!
'! Usage:
'!
'!   cscript ans2html.vbs path_to_ansi.ans path_to_html.html [page_title] [opts]
'!
'! The "opts" can be any or all processing modes:
'!
'!   P: Pipe codes
'!   T: Tilde codes
'!   L: RTSoft "LoRD" codes
'!   Y: Yankee Trader Galactic Newspaper bulletin prefixes
'! 
'! You must specify a page title if using one of the optional processing modes.
'! 
'! Example:
'!
'!   cscript ans2html.vbs c:\lord\LOGNOW.TXT c:\web\lord_news.html "LorD News" L
'!
'! Probably goes without saying, but paths containing spaces must be wrapped
'! in double-quotes.
'!
'! Enjoy!
'!

Option Explicit
On Error Resume Next

' Constants -------------------------------------------------------------------
Const FOR_READING = 1, FOR_WRITING = 2, FOR_APPEND = 8

' HTML hex values for the 16 ANSI text mode colors.
'! @see https://en.wikipedia.org/wiki/ANSI_escape_code#Colors
Const BLACK = "#000"
Const RED = "#A00"
Const GREEN = "#0A0"
Const BROWN = "#A50" ' low-intensity yellow
Const BLUE = "#00A"
Const MAGENTA = "#A0A"
Const CYAN = "#0AA"
Const GRAY = "#AAA" ' low-intensity white
Const DARKGRAY = "#555" ' high-intensity black
Const LIGHTRED = "#F55"
Const LIGHTGREEN = "#5F5"
Const YELLOW = "#FF5"
Const LIGHTBLUE = "#55F"
Const LIGHTMAGENTA = "#F5F"
Const LIGHTCYAN = "#5FF"
Const WHITE = "#FFF"

' Best font for displaying CP437 ANSI.
Const FONT_FAMILY = "Web437 IBM VGA"
Const FONT_PATH = "../styles/web437-ibm-vga.css"

' Variables         Description -----------------------------------------------
Dim CSI             ' See: https://en.wikipedia.org/wiki/ANSI_escape_code#CSI_sequences
Dim args            ' Incoming arguments
Dim fso             ' FileSystemObject
Dim fAnsiSource     ' ANSI source file object
Dim sourceFile      ' ANSI source filename
Dim targetFile      ' HTML output filename
Dim title           ' Title of HTML output file
Dim ansiData        ' Contents of ANSI source file
Dim charAtI         ' Character At I
Dim charCode        ' ASCII code of character at I
Dim escapeSequence  ' ANSI escape code
Dim csiFinalByte    ' Function of the ANSI escape sequence
Dim csiParams       ' Parameters of the ANSI escape sequence
Dim sgrParam        ' SGR (select graphic rendition) parameter
Dim spanTag         ' <span> tag to be written
Dim blink           ' CSS blink class
Dim bgColor         ' HTML Background color
Dim fgColor         ' HTML Foreground color
Dim fgIntensity     ' Foreground intensity
Dim swapColor       ' Swap foreground and background colors
Dim holdColor       ' Variable for holding a color if swapping
Dim startPos        ' Reading start position. For now just 1.
Dim colPos          ' Output column number. Used for auto-wrapping the output.
Dim ignoreLF        ' Ignore line feeds
Dim i               ' Main loop counter
Dim oStream         ' Output stream
Dim convertOptions  ' Argument for conversion options
Dim convertPipes    ' P - Convert pipe codes? Boolean
Dim convertTildes   ' T - Convert tilde codes? Boolean
Dim convertLord     ' L - Convert LoRD codes? Boolean
Dim convertYT       ' Y - Colorize Yankee Trader bulletins? Boolean
Dim cp437html       ' Array for holding HTML entities for the CP437 ANSI chars

' Initialize Start-------------------------------------------------------------
cp437html = Array ("", _
   "&#x263A;","&#x263B;","&#x2665;","&#x2666;","&#x2663;","&#x2660;",_
   "&#x2022;","&#x25D8;","&#x25CB;","&#x25D9;","&#x2642;","&#x2640;", _
   "&#x266A;","&#x266B;","&#x263C;","&#x25BA;","&#x25C0;","&#x2195;", _
   "&#x203C;","&#x00B6;","&#x00A7;","&#x25AC;","&#x21A8;","&#x2191;", _
   "&#x2193;","&#x2192;","&#x2190;","&#x221F;","&#x2194;","&#x25B2;", _
   "&#x25BC;","&#32;","&#33;","&#34;","&#35;","&#36;","&#37;","&#38;", _
   "&#39;","&#40;","&#41;","&#42;","&#43;","&#44;","&#45;","&#46;","&#47;", _
   "&#48;","&#49;","&#50;","&#51;","&#52;","&#53;","&#54;","&#55;","&#56;", _
   "&#57;","&#58;","&#59;","&#60;","&#61;","&#62;","&#63;","&#64;","&#65;", _
   "&#66;","&#67;","&#68;","&#69;","&#70;","&#71;","&#72;","&#73;","&#74;", _
   "&#75;","&#76;","&#77;","&#78;","&#79;","&#80;","&#81;","&#82;","&#83;", _
   "&#84;","&#85;","&#86;","&#87;","&#88;","&#89;","&#90;","&#91;","&#92;", _
   "&#93;","&#94;","&#95;","&#96;","&#97;","&#98;","&#99;","&#100;","&#101;", _
   "&#102;","&#103;","&#104;","&#105;","&#106;","&#107;","&#108;","&#109;", _
   "&#110;","&#111;","&#112;","&#113;","&#114;","&#115;","&#116;","&#117;", _
   "&#118;","&#119;","&#120;","&#121;","&#122;","&#123;","&#124;","&#125;", _
   "&#126;","&#x2302;","&#x00C7;","&#x00FC;","&#x00E9;","&#x00E2;","&#x00E4;", _
   "&#x00E0;","&#x00E5;","&#x00E7;","&#x00EA;","&#x00EB;","&#x00E8;", _
   "&#x00EF;","&#x00EE;","&#x00EC;","&#x00C4;","&#x00C5;","&#x00C9;", _
   "&#x00E6;","&#x00C6;","&#x00F4;","&#x00F6;","&#x00F2;","&#x00FB;", _
   "&#x00F9;","&#x00FF;","&#x00D6;","&#x00DC;","&#x00A2;","&#x00A3;", _
   "&#x00A5;","&#x20A7;","&#x0192;","&#x00E1;","&#x00ED;","&#x00F3;", _
   "&#x00FA;","&#x00F1;","&#x00D1;","&#x00AA;","&#x00BA;","&#x00BF;", _
   "&#x2310;","&#x00AC;","&#x00BD;","&#x00BC;","&#x00A1;","&#x00AB;", _
   "&#x00BB;","&#x2591;","&#x2592;","&#x2593;","&#x2502;","&#x2524;", _
   "&#x2561;","&#x2562;","&#x2556;","&#x2555;","&#x2563;","&#x2551;", _
   "&#x2557;","&#x255D;","&#x255C;","&#x255B;","&#x2510;","&#x2514;", _
   "&#x2534;","&#x252C;","&#x251C;","&#x2500;","&#x253C;","&#x255E;", _
   "&#x255F;","&#x255A;","&#x2554;","&#x2569;","&#x2566;","&#x2560;", _
   "&#x2550;","&#x256C;","&#x2567;","&#x2568;","&#x2564;","&#x2565;", _
   "&#x2559;","&#x2558;","&#x2552;","&#x2553;","&#x256B;","&#x256A;", _
   "&#x2518;","&#x250C;","&#x2588;","&#x2584;","&#x258C;","&#x2590;", _
   "&#x2580;","&#x03B1;","&#x00DF;","&#x0393;","&#x03C0;","&#x03A3;", _
   "&#x03C3;","&#x00B5;","&#x03C4;","&#x03A6;","&#x0398;","&#x03A9;", _
   "&#x03B4;","&#x221E;","&#x03C6;","&#x03B5;","&#x2229;","&#x2261;", _
   "&#x00B1;","&#x2265;","&#x2264;","&#x2320;","&#x2321;","&#x00F7;", _
   "&#x2248;","&#x00B0;","&#x2219;","&#x00B7;","&#x221A;","&#x207F;", _
   "&#x00B2;","&#x25A0;","")
convertPipes = False
convertTildes = False
convertLord = False
convertYT = False
CSI = chr(27) & "["
startPos = 1
colPos = 0
fgIntensity = 0
fgColor = GRAY
bgColor = BLACK
ignoreLF = False

Set args = Wscript.Arguments
sourceFile = args(0)
targetFile = args(1)
title = args(2)
convertOptions = UCase(args(3))

If InStr(convertOptions, "P") >= 1 Then
   convertPipes = True
   Wscript.echo "Pipe color conversion enabled"
End If
If InStr(convertOptions, "T") >= 1 Then
   convertTildes = True
   Wscript.echo "Tilde color conversion enabled"
End If
If InStr(convertOptions, "L") >= 1 Then
   convertLord = True
   Wscript.echo "LoRD color conversion enabled"
End If
If InStr(convertOptions, "Y") >= 1 Then
   convertYT = True
   Wscript.echo "Yankee Trader colorizing enabled"
End If

Set fso = CreateObject("Scripting.FileSystemObject")
' Initialize End --------------------------------------------------------------

If fso.FileExists(sourceFile) Then
   Wscript.echo title
   Set oStream = CreateObject("ADODB.Stream")
   oStream.charSet = "ASCII"
   oStream.Open
   oStream.WriteText "<!DOCTYPE html>" & vbCrlf & "<html lang='en'>" & vbCrlf & _
                "<head>" & vbCrlf & "<meta charset='UTF-8'>" & vbCrlf & "<title>" & title & "</title>" & vbCrlf & _
                "<style>" & vbCrlf & ".blink{animation:blinker 0.8s infinite step-end;}" & vbCrlf & "@keyframes blinker{50%{color:hsla(0,0%,0%,0.0);}}" & vbCrlf & "</style>" & vbCrlf & _
                "<link rel='stylesheet' type='text/css' href='" & FONT_PATH & "'>" & vbCrlf & "</head>" & vbCrlf & _
                "<body style='color:" & fgColor & ";background-color:" & bgColor & ";'>" & _
                "<!-- " & title & " file generated on " & Now & " -->" & vbCrlf & _
                "<pre style='font-family:""" & FONT_FAMILY & """,monospace;'>" & vbCrLf

   ' Open the ANSI source file.
   Set fAnsiSource = fso.OpenTextFile(sourceFile, FOR_READING, False, 0)
   If convertOptions = "" Then
      ansiData = FlattenAnsi(fAnsiSource.ReadAll)
   Else
      ansiData = fAnsiSource.ReadAll
   End If
   fAnsiSource.Close

   ' This simply inserts some pipe codes for Yankee Trader's galactic newspaper
   ' bulletin, and then those pipe codes get converted.
   If convertYT Then
      ansiData = Replace(ansiData, vbCrLf & " *** ", vbCrLf & "|09 *** ") ' attacks
      ansiData = Replace(ansiData, vbCrLf & " +++ ", vbCrLf & "|12 +++ ") ' xannor defeats
      ansiData = Replace(ansiData, vbCrLf & "  -  ", vbCrLf & "|14  -  ") ' planetary headlines
      ansiData = Replace(ansiData, vbCrLf & "-=*=-", vbCrLf & "|15-=*=-") ' Logons
      ansiData = Replace(ansiData, vbCrLf, vbCrLf & "|02")                ' standard newspaper color
      convertPipes = True
   End If

   ' Begin reading the ANSI contents
   For i = startPos To Len(ansiData)
      charAtI = Mid(ansiData, i, 1)
      charCode = Asc(charAtI)
      ' Wrap at 80 columns...
      ' If this is the 80th OUTPUT character, and the next INPUT character is
      ' NOT a CR or LF, we need to add a <br/> here now.
      If colPos = 80 Then
         colPos = 0
         If charCode <> 13 and charCode <> 10 Then
            oStream.WriteText vbCrlf
         End If
      End If
      If charCode = 13 Then
         ignoreLF = True
         oStream.WriteText vbCrlf
         colPos = 0
      ElseIf charCode = 10 Then
         If ignoreLF = False Then
            oStream.WriteText vbCrlf
            colPos = 0
         End If
      ElseIf charCode = 32 Then
         oStream.WriteText " "
         colPos = colPos + 1
      ElseIf charCode = 0 Then
         oStream.WriteText " "
         colPos = colPos + 1
      ElseIf charAtI = "<" Then
         oStream.WriteText "&lt;"
         colPos = colPos + 1
      ElseIf charAtI = ">" Then
         oStream.WriteText "&gt;"
         colPos = colPos + 1
      ElseIf charAtI = "&" Then
         oStream.WriteText "&amp;"
         colPos = colPos + 1
      ElseIf charAtI = "'" Then
         oStream.WriteText "&apos;"
         colPos = colPos + 1
      ElseIf charAtI = """" Then
         oStream.WriteText "&quot;"
         colPos = colPos + 1
      ElseIf Mid(ansiData, i, 2) = CSI Then
         ' Terminate the previous span tag if one was started.
         If spanTag <> "" Then
            oStream.WriteText "</span>"
         End If
         ' Locate the next alpha after this point
         escapeSequence = Mid(ansiData, i, InStrNextAlpha(i, ansiData, csiFinalByte) - i)
         csiParams = Mid(escapeSequence, 3)        
         i = i + Len(escapeSequence) ' Advance the parser.
         Select Case csiFinalByte
            ' See: http://ascii-table.com/ansi-escape-sequences.php
            '      https://en.wikipedia.org/wiki/ANSI_escape_code#Terminal_output_sequences
            '
            ' Only "m" is implemented here. The rest are all implemented in
            ' the FlattenAnsi function.
            Case "m"
               ' Set Graphics Mode
               ' See: https://en.wikipedia.org/wiki/ANSI_escape_code#SGR
               For Each sgrParam In Split(csiParams, ";")
                  Select Case sgrParam
                     Case 0
                        fgIntensity = 0
                        blink = ""
                        swapColor = 0
                        bgColor = BLACK
                        fgColor = GRAY
                     Case 1
                        fgIntensity = 1
                     Case 3
                        ' Italic. Sometimes treated as inverse (Swap colors?).
                        swapColor = 1
                     Case 5
                        blink = "class='blink' "
                     Case 7
                        swapColor = 1
                     Case 30
                        fgColor = BLACK
                     Case 31
                        fgColor = RED
                     Case 32
                        fgColor = GREEN
                     Case 33
                        fgColor = BROWN
                     Case 34
                        fgColor = BLUE
                     Case 35
                        fgColor = MAGENTA
                     Case 36
                        fgColor = CYAN
                     Case 37
                        fgColor = GRAY
                     Case 40
                        bgColor = BLACK
                     Case 41
                        bgColor = RED
                     Case 42
                        bgColor = GREEN
                     Case 43
                        bgColor = BROWN
                     Case 44
                        bgColor = BLUE
                     Case 45
                        bgColor = MAGENTA
                     Case 46
                        bgColor = CYAN
                     Case 47
                        bgColor = GRAY
                  End Select
               Next
               If swapColor = 1 Then
                  holdColor = fgColor
                  fgColor = bgColor
                  bgColor = holdColor
               End If
               spanTag = "<span " & blink & "style='color:" & SetColorIntensity(fgColor, fgIntensity) & ";background-color:" & bgColor & ";'>"
         End Select
         oStream.WriteText spanTag
      ElseIf charAtI = "|" And convertPipes Then
         ' Terminate the previous span tag if one was started.
         If spanTag <> "" Then
            oStream.WriteText "</span>"
         End If
         ' The following pulled from "How to use Color" secion in Jezebel's INSTRUCT.DOC file:
         ' =====================================================================
         ' Renegade's colors....
         ' You take a pipe, | (it's that thing above your \ key) and a number between
         ' 01 and 15, MAKE SURE it's 2 digits, not |1, it has to be |01.. ;)
         ' 01 is Blue 02 is green 03 is cyan 04 is red 05 is magenta 06 is brown
         ' 07 is grey 08 lt black 09 lt blue 10 lt green 11 lt cyan 12 lt red
         ' 13 lt magenta 14 yellow 15 white
         ' |03This would be Cyan
         Select Case Mid(ansiData, i + 1, 2)
            Case "00"
               fgColor = BLACK
               fgIntensity = 0
               blink = ""
            Case "01"
               fgColor = BLUE
               fgIntensity = 0
               blink = ""
            Case "02"
               fgColor = GREEN
               fgIntensity = 0
               blink = ""
            Case "03"
               fgColor = CYAN
               fgIntensity = 0
               blink = ""
            Case "04"
               fgColor = RED
               fgIntensity = 0
               blink = ""
            Case "05"
               fgColor = MAGENTA
               fgIntensity = 0
               blink = ""
            Case "06"
               fgColor = BROWN
               fgIntensity = 0
               blink = ""
            Case "07"
               fgColor = GRAY
               fgIntensity = 0
               blink = ""
            Case "08"
               fgColor = BLACK
               fgIntensity = 1
               blink = ""
            Case "09"
               fgColor = BLUE
               fgIntensity = 1
               blink = ""
            Case "10"
               fgColor = GREEN
               fgIntensity = 1
               blink = ""
            Case "11"
               fgColor = CYAN
               fgIntensity = 1
               blink = ""
            Case "12"
               fgColor = RED
               fgIntensity = 1
               blink = ""
            Case "13"
               fgColor = MAGENTA
               fgIntensity = 1
               blink = ""
            Case "14"
               fgColor = BROWN
               fgIntensity = 1
               blink = ""
            Case "15"
               fgColor = GRAY
               fgIntensity = 1
               blink = ""
            Case "16"
               bgColor = BLACK
            Case "17"
               bgColor = BLUE
            Case "18"
               bgColor = GREEN
            Case "19"
               bgColor = CYAN
            Case "20"
               bgColor = RED
            Case "21"
               bgColor = MAGENTA
            Case "22"
               bgColor = BROWN
            Case "23"
               bgColor = GRAY
            Case "24"
               fgColor = BLACK
               fgIntensity = 1
               blink = "class='blink' "
            Case "25"
               fgColor = BLUE
               fgIntensity = 1
               blink = "class='blink' "
            Case "26"
               fgColor = GREEN
               fgIntensity = 1
               blink = "class='blink' "
            Case "27"
               fgColor = CYAN
               fgIntensity = 1
               blink = "class='blink' "
            Case "28"
               fgColor = RED
               fgIntensity = 1
               blink = "class='blink' "
            Case "29"
               fgColor = MAGENTA
               fgIntensity = 1
               blink = "class='blink' "
            Case "30"
               fgColor = BROWN
               fgIntensity = 1
               blink = "class='blink' "
            Case "31"
               fgColor = GRAY
               fgIntensity = 1
               blink = "class='blink' "           
            Case "AL" ' Begin DARKNESS colors 
               fgColor = RED
               fgIntensity = 1
               blink = ""
            Case "DE" '
               fgColor = BLACK
               fgIntensity = 1
               blink = ""
            Case "DI" '
               fgColor = BLACK
               fgIntensity = 1
               blink = ""
            Case "DT" '
               fgColor = GRAY
               fgIntensity = 0
               blink = ""
            Case "LT" '
               fgColor = GRAY
               fgIntensity = 1
               blink = ""
            Case "H1"
               fgColor = BROWN
               fgIntensity = 1
               blink = ""
            Case "H2" '
               fgColor = CYAN
               fgIntensity = 0
               blink = ""
            Case "TI"
               fgColor = GRAY
               fgIntensity = 1
               blink = ""
            Case "IN"
               fgColor = GREEN
               fgIntensity = 1
               blink = ""
         End Select        
         i = i + 2 ' Advance the parser.
         spanTag = "<span " & blink & "style='color:" & SetColorIntensity(fgColor, fgIntensity) & ";background-color:" & bgColor & ";'>"
         oStream.WriteText spanTag
      ElseIf charAtI = "~" And convertTildes Then
         ' Terminate the previous span tag if one was started.
         If spanTag <> "" Then
            oStream.WriteText "</span>"
         End If
         ' The following pulled from "Special Control Codes" secion in SYSOP.DOC
         ' of Death Masters:
         ' =====================================================================
         ' + Special Control Codes

         '   There are MANY control codes in this file.  Most of them are forbidden to
         '   you.  The only ones you can fool around with are:
         '   ~1 Change text to GREEN until a new ~# sequence is found
         '   ~2 Change text to BLUE until a new ~# sequence is found
         '   ~3 Change text to CYAN until a new ~# sequence is found
         '   ~4 Change text to RED until a new ~# sequence is found
         '   ~5 Change text to MAGENTA until a new ~# sequence is found
         '   ~6 Change text to BROWN until a new ~# sequence is found
         '   ~7 Change text to LIGHT GREY until a new ~# sequence is found
         '   ~8 Change text to DARK GREY until a new ~# sequence is found
         '   ~9 Change text to BRIGHT BLUE until a new ~# sequence is found
         '   ~a Change text to BRIGHT GREEN until a new ~# sequence is found
         '   ~b Change text to BRIGHT CYAN until a new ~# sequence is found
         '   ~c Change text to BRIGHT RED a new ~# sequence is found
         '   ~d Change text to BRIGHT MAGENTA until a new ~# sequence is found
         '   ~e Change text to YELLOW a new ~# sequence is found
         '   ~f Change text to BRIGHT WHITE until a new ~# sequence is found
         Select Case Mid(ansiData, i + 1, 1)
            Case "1"
               fgColor = GREEN
               fgIntensity = 0
            Case "2"
               fgColor = BLUE
               fgIntensity = 0
            Case "3"
               fgColor = CYAN
               fgIntensity = 0
            Case "4"
               fgColor = RED
               fgIntensity = 0
            Case "5"
               fgColor = MAGENTA
               fgIntensity = 0
            Case "6"
               fgColor = BROWN
               fgIntensity = 0
            Case "7"
               fgColor = GRAY
               fgIntensity = 0
            Case "8"
               fgColor = BLACK
               fgIntensity = 1
            Case "9"
               fgColor = BLUE
               fgIntensity = 1
            Case "a"
               fgColor = GREEN
               fgIntensity = 1
            Case "b"
               fgColor = CYAN
               fgIntensity = 1
            Case "c"
               fgColor = RED
               fgIntensity = 1
            Case "d"
               fgColor = MAGENTA
               fgIntensity = 1
            Case "e"
               fgColor = BROWN
               fgIntensity = 1
            Case "f"
               fgColor = GRAY
               fgIntensity = 1
         End Select        
         i = i + 1 ' Advance the parser.
         spanTag = "<span " & blink & "style='color:" & SetColorIntensity(fgColor, fgIntensity) & ";background-color:" & bgColor & ";'>"
         oStream.WriteText spanTag
      ElseIf charAtI = "`" And convertLord Then
         ' Terminate the previous span tag if one was started.
         If spanTag <> "" Then
            oStream.WriteText "</span>"
         End If
         ' The following pulled from "Screen Commands" secion in LADY.DOC:
         ' =====================================================================
         ' foreground color -
         ' `1 dark blue     `6 brownish      `! light cyan     and seldom used
         ' `2 dark green    `7 grey          `@ light red      `^ black
         ' `3 dark cyan     `8 dark grey     `# light violet
         ' `4 dark red      `9 light blue    `$ yellow
         ' `5 dark violet   `0 light green   `% white

         ' ** Note.. The black foreground here is only available here. Lady authors
         ' are expected to use it wisely..

         ' background color -
         ' `r0 black               `r4 dark red
         ' `r1 dark blue           `r5 dark violet
         ' `r2 dark green          `r6 brownish
         ' `r3 dark cyan           `r7 grey
         Select Case Mid(ansiData, i + 1, 1)
            Case "."
               ' Apparently an undocumented reset.
               fgIntensity = 0
               bgColor = BLACK
               fgColor = GRAY
            Case "1"
               fgColor = BLUE
               fgIntensity = 0
            Case "2"
               fgColor = GREEN
               fgIntensity = 0
            Case "3"
               fgColor = CYAN
               fgIntensity = 0
            Case "4"
               fgColor = RED
               fgIntensity = 0
            Case "5"
               fgColor = MAGENTA
               fgIntensity = 0
            Case "6"
               fgColor = BROWN
               fgIntensity = 0
            Case "7"
               fgColor = GRAY
               fgIntensity = 0
            Case "8"
               fgColor = BLACK
               fgIntensity = 1
            Case "9"
               fgColor = BLUE
               fgIntensity = 1
            Case "0"
               fgColor = GREEN
               fgIntensity = 1
            Case "!"
               fgColor = CYAN
               fgIntensity = 1
            Case "@"
               fgColor = RED
               fgIntensity = 1
            Case "#"
               fgColor = MAGENTA
               fgIntensity = 1
            Case "$"
               fgColor = BROWN
               fgIntensity = 1
            Case "%"
               fgColor = GRAY
               fgIntensity = 1
            Case "^"
               fgColor = BLACK
               fgIntensity = 0
            Case "r"
               ' Select the NEXT character for the background color.
               Select Case Mid(ansiData, i + 2, 1)
                  Case "0"
                     bgColor = BLACK
                  Case "1"
                     bgColor = BLUE
                  Case "2"
                     bgColor = GREEN
                  Case "3"
                     bgColor = CYAN
                  Case "4"
                     bgColor = RED
                  Case "5"
                     bgColor = MAGENTA
                  Case "6"
                     bgColor = BROWN
                  Case "7"
                     bgColor = GRAY
               End Select              
               i = i + 1 ' Advance the parser again.
         End Select         
         i = i + 1 ' Advance the parser.
         spanTag = "<span " & blink & "style='color:" & SetColorIntensity(fgColor, fgIntensity) & ";background-color:" & bgColor & ";'>"
         oStream.WriteText spanTag
      ElseIf (charCode >= 1 And charCode <= 31) Or (charCode >= 127 And charCode <= 254) Then
         oStream.WriteText cp437html(charCode)
         colPos = colPos + 1
      Else
         oStream.WriteText charAtI
         colPos = colPos + 1
      End If
   Next

   ' Terminate the last span tag if needed.
   If spanTag <> "" Then
      oStream.WriteText "</span>"
   End If

   oStream.WriteText vbCrLf & "</pre>"
   oStream.WriteText vbCrlf & "</body>" & vbCrlf & "</html>"
   oStream.SaveToFile targetFile, 2
   oStream.Close
End If

' *** END ***

' *** FUNCTIONS ***

'! Removes cursor movement sequences.
'! Result should be the final ANSI image displayed, rather than the
'! cursor-by-cursor movement or animation that generated it.
'!
'! @param  ansiData   The raw ANSI data from the file.
'! @return            The rearranged ANSI data containing only "m" sequences.
Function FlattenAnsi(ansiData)
   Const MAX_COLS = 80 ' Standard 80 column width
   Const STARTING_ROWS = 5
   Dim returnAnsi
   Dim j
   Dim charAtJ
   Dim chrCode
   Dim escSeq
   Dim csiLastByte
   Dim csiArgs
   Dim row
   Dim col
   Dim rowSav
   Dim colSav
   Dim newEscSeq
   Dim args
   Dim cBuf
   Dim rBuf
   Dim adding
   Dim a

   ReDim screenBuffer(MAX_COLS, STARTING_ROWS)
   For rBuf = 0 To UBound(screenBuffer, 2)
      For cBuf = 0 To UBound(screenBuffer, 1)
         screenBuffer(cBuf, rBuf) = " "
      Next
   Next

   row = 0
   col = 1
   newEscSeq = ""

   For j = 1 To Len(ansiData)
      charAtJ = Mid(ansiData, j, 1)
      chrCode = Asc(charAtJ)

      If Mid(ansiData, j, 2) = CSI Then

         escSeq = Mid(ansiData, j, InStrNextAlpha(j, ansiData, csiLastByte) - j)
         csiArgs = Mid(escSeq, 3)        
         j = j + Len(escSeq) ' Advance the parser.

         Select Case csiLastByte
            Case "H", "f" ' Cursor position
               IF InStr(csiArgs, ";") > 0 Then
                  args = Split(csiArgs, ";")
                  row = CInt(args(0)) ' n
                  col = 1
                  If Ubound(args) > 0 Then
                     col = CInt(args(1)) ' m
                  End If
               ElseIf csiArgs <> "" Then
                  row = CInt(csiArgs) ' n
                  col = 1
               Else
                  row = 1
                  col = 1
               End If
            Case "A"   ' Cursor Up
               newEscSeq = ""
               If csiArgs = "" Then
                  csiArgs = 1
               End If
               row = row - CInt(csiArgs)
            Case "B"   ' Cursor Down
               newEscSeq = ""
               If csiArgs = "" Then
                  csiArgs = 1
               End If
               row = row + CInt(csiArgs)
            Case "C"   ' Cursor Forward
               ' Cancel the last SGR sequence before moving the "cursor",
               ' otherwise it will drag the sequence with it, leading
               ' to an incorrect background and/or foreground color.
               if col <= MAX_COLS And row <= UBound(screenBuffer, 2) Then
                  screenBuffer(col, row) = CSI & "40m" & screenBuffer(col, row)
               End If
               newEscSeq = ""
               If csiArgs = "" Then
                  csiArgs = 1
               End If
               col = col + CInt(csiArgs)
            Case "D"   ' Cursor Backward
               If col <= MAX_COLS And row <= UBound(screenBuffer, 2) Then
                  screenBuffer(col, row) = CSI & "40m" & screenBuffer(col, row)
               End If
               newEscSeq = ""
               If csiArgs = "" Then
                  csiArgs = 1
               End If
               col = col - CInt(csiArgs)
            Case "s"   ' Save cursor position
               rowSav = CInt(row)
               colSav = CInt(col)
            Case "u"   ' Restore cursor position
               row = CInt(rowSav)
               col = CInt(colSav)
            'Probably won't bother implementing these:
            ' Case "2J"  ' Erase display (?)
            ' Case "K"   ' Erase line (?) -- Problematic..?
            ' Case "h"   ' Set mode (screen width/height). 
            ' Case "l"   ' Reset mode
            ' Case "p"   ' Set keyboard strings
            Case Else ' Store the escape sequence to travel with the next characters
               newEscSeq = newEscSeq & escSeq & csiLastByte

         End Select

         If row < 1 Then
            row = 1
         End If
         If col < 1 Then
            col = 1
         End If

      Else

         If chrCode = 13 Then
            ignoreLF = True
            if row <= UBound(screenBuffer, 2) then
               screenBuffer(col, row) = newEscSeq & screenBuffer(col, row)
            end if
            row = row + 1
            col = 1
            If j + 1 <= Len(ansiData) Then
               If Asc(Mid(ansiData, j + 1, 1)) = 10 Then
                  j = j + 1 ' Advance the parser past the LF.
               End If
            End If
         ElseIf chrCode = 10 Then
            if ignoreLF = False then
               if row <= UBound(screenBuffer, 2) then
                  screenBuffer(col, row) = newEscSeq & screenBuffer(col, row)
               end if
               row = row + 1
               col = 1
            end if
         Else
            ' Append a line if it goes beyond the current max.
            If row > UBound(screenBuffer, 2) Then
               adding = row - UBound(screenBuffer, 2)
               ReDim Preserve screenBuffer(MAX_COLS, row)

               ' Initialize the new row with all blanks/CR-LFs.
               For a = 1 to adding Step 1
                  For cBuf = 1 To UBound(screenBuffer, 1)
                     screenBuffer(cBuf, row-(a-1)) = " "
                  Next
               Next
            End If

            screenBuffer(col, row) = newEscSeq & charAtJ

            ' Clear the newEscSeq after using it, don't need it again.
            If newEscSeq <> "" Then
               newEscSeq = ""
            End If

            col = col + 1

            If col > MAX_COLS Then
               col = 1
               row = row + 1
            End If
         End If
      End If

   Next

   ' Now form new ansiData out of the screen buffer contents.
   For rBuf = 0 To UBound(screenBuffer, 2)
      For cBuf = 1 To UBound(screenBuffer, 1)
         ' TODO - replace this with an ADO Stream object...
         '        No more concatenation.
         returnAnsi = returnAnsi & screenBuffer(cBuf, rBuf)
      Next
   Next

   FlattenAnsi = returnAnsi
End Function

'! Searches a string for the position of the next alpha character.
'!
'! @param  startIndex      The index of the string to start the search.
'! @param  stringToSearch  The string to search.
'! @param  alphaFound      Holds the alpha character that was found.
'! @return                 The index of the alpha character found.
Function InStrNextAlpha(startIndex, stringToSearch, ByRef alphaFound)
   Dim j
   Dim thisAsc
   Dim returnIndex
   Dim cha
   alphaFound = ""
   For j = startIndex To Len(stringToSearch)
      cha = Mid(stringToSearch, j, 1)
      thisAsc = Asc(cha)
      If ((thisAsc >= 65 And thisAsc <= 90) Or (thisAsc >= 97 And thisAsc <= 122)) and alphaFound = "" Then
         returnIndex = j
         alphaFound = cha
         Exit For
      End If
   Next
   InStrNextAlpha = returnIndex
End Function

'! Determines whether to intensify a hex color value.
'! (I'm certain there's a more elegant way to intensify the color values
'! programmatically, but this way works fine.)
'!
'! @param  color      Any one of the 8 hex color value constants.
'! @param  intensity  If 1, return the intensified color, 0 if normal.
'! @return            Hex value for either the normal or intensified color.
Function SetColorIntensity(color, intensity)
   If intensity = 1 Then
      Select Case color
         Case BLACK
            SetColorIntensity = DARKGRAY
         Case RED
            SetColorIntensity = LIGHTRED
         Case GREEN
            SetColorIntensity = LIGHTGREEN
         Case BROWN
            SetColorIntensity = YELLOW
         Case BLUE
            SetColorIntensity = LIGHTBLUE
         Case MAGENTA
            SetColorIntensity = LIGHTMAGENTA
         Case CYAN
            SetColorIntensity = LIGHTCYAN
         Case GRAY
            SetColorIntensity = WHITE
         Case Else
            SetColorIntensity = color
      End Select
   Else
      SetColorIntensity = color
   End If
End Function
