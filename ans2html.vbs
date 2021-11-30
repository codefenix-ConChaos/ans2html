'! ANSI-To-HTML Converter
'! (ans2html.vbs)
'! Author:           Craig Hendricks (aka Codefenix)
'! Initial Version:  3/15/2021
'!
'! ============================================================================
'!
'! Description:
'!
'!   This VBScript is for converting an ANSI file to a HTML file.  Useful for
'!   displaying BBS door game scores on a website.
'!
'!   It reads each character from a standard ANSI source file and generates
'!   a file containing HTML5 markup.  It interprets most ANSI escape codes
'!   and translates all 255 codepage 437 characters to the best matching
'!   equivalent HTML entity.
'!   See: https://en.wikipedia.org/wiki/Code_page_437
'!
'!   After reading the ANSI source data, the script will first "flatten" it,
'!   eliminating all cursor movement sequences so that it need only convert 
'!   the "m" escape sequences for in-line text coloring.
'!
'!   The "Source Code Pro" font is optional but highly recommended for best
'!   results. Download it from https://github.com/adobe-fonts/source-code-pro
'!   and install it as a web font for your site.  If this font is not present,
'!   web browsers will default to whatever monospace font is configured,
'!   leading to mixed results for box and line drawing characters,
'!   especially on mobile browsers.
'!
'! Known issues:
'!
'!   Blinking text is achieved using CSS (built into the "htmlOutput" string)
'!   but it needs work to allow alternating between foreground and background
'!   colors.  Currently it just blinks them simultaneously.
'!
'! Usage:
'!
'!   cscript ans2html.vbs path_to_ansi.ans path_to_html.html [page_title]
'!
'! Probably goes without saying, but paths containing spaces must be wrapped
'! in double-quotes.
'!
'!
'! TODO:
'! - Improve CSS blink.
'!
'!

Option Explicit
On Error Resume Next

' Constants
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

' Best font for box drawing.
Const FONT_FAMILY = """Source Code Pro"",monospace"
Const FONT_SIZE = "13px" ' Set this to whatever size you like best.

' Variables         Description:
Dim CSI             ' See: https://en.wikipedia.org/wiki/ANSI_escape_code#CSI_sequences
Dim args            ' Incoming arguments
Dim fso             ' FileSystemObject
Dim fAnsiSource     ' ANSI source file object
Dim sourceFile      ' ANSI source filename
Dim targetFile      ' HTML output filename
Dim title           ' Title of HTML output file
Dim ansiData        ' Contents of ANSI source file
Dim htmlOutput      ' HTML markup to be written
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
Dim i
Dim oStream

' Initialize

Set oStream = CreateObject("ADODB.Stream")
oStream.charSet = "ASCII"
oStream.Open

Set args = Wscript.Arguments
sourceFile = args(0)
targetFile = args(1)
title = args(2)
Wscript.echo title

CSI = chr(27) & "["
startPos = 1
colPos = 0
fgIntensity = 0
fgColor = GRAY
bgColor = BLACK

oStream.WriteText "<div style='font-family:" & FONT_FAMILY & ";white-space:nowrap;padding:0;color:" & fgColor & ";background-color:" & bgColor & ";'>" & vbCrlf & _
             "<style>" & vbCrlf & ".blink{animation:blinker 1s linear infinite;}" & vbCrlf & "@keyframes blinker{50%{opacity:0;}}" & vbCrlf & "</style>" & vbCrlf & _
             "<!-- " & title & " file generated on " & Now & " -->" & vbCrlf & _
             "<pre style='font-family:" & FONT_FAMILY & ";'>" & vbCrLf

Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(sourceFile) Then

   ' Open the ANSI source file.
   Set fAnsiSource = fso.OpenTextFile(sourceFile, FOR_READING)
   ansiData = FlattenAnsi(fAnsiSource.ReadAll)
   fAnsiSource.Close
   
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
            'oStream.WriteText "<br/>" & vbCrlf
            oStream.WriteText vbCrlf
         End If
      End If

      If charCode = 13 Then
         'oStream.WriteText "<br/>" & vbCrlf
         oStream.WriteText vbCrlf
         colPos = 0
         If Asc(Mid(ansiData, i + 1, 1)) = 10 Then
            i = i + 1 ' Advance the parser past the LF.
         End If
      ElseIf charCode = 10 Then
         'oStream.WriteText "<br/>" & vbCrlf
         oStream.WriteText vbCrlf
         colPos = 0
      ElseIf charCode = 32 Then
         'oStream.WriteText "&nbsp;"
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
         oStream.WriteText "&#39;"
         colPos = colPos + 1
      ElseIf charAtI = """" Then
         oStream.WriteText "&quot;"
         colPos = colPos + 1
         
      ElseIf Mid(ansiData, i, 2) = CSI Then
         ' Start of ANSI escape sequence...
      
         'Wscript.echo "CSI at " & i

         ' Terminate the previous span tag if one was started.
         If spanTag <> "" Then
            oStream.WriteText "</span>"
         End If

         ' Locate the next alpha after this point
         escapeSequence = Mid(ansiData, i, InStrNextAlpha(i, ansiData, csiFinalByte) - i)
         csiParams = Mid(escapeSequence, 3)
         'Wscript.echo csiParams

         ' Advance the parser.
         i = i + Len(escapeSequence)

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
         'Wscript.echo spanTag

         oStream.WriteText spanTag
      ElseIf (charCode >= 1 And charCode <= 31) Or (charCode >= 127 And charCode <= 254) Then
         oStream.WriteText ToHtmlEntity(charCode)
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

   oStream.WriteText vbCrLf & "</pre><br/><br/><span style='color:" & DARKGRAY & ";background-color:" & BLACK & ";'>(updated at " & (fso.GetFile(sourceFile)).DateLastModified & ")</span>" & vbCrLf
   oStream.WriteText vbCrlf & "</div>"
   
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
'!
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
   'Dim prevCol
   'Dim prevRow
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

         ' Locate the next alpha after this point
         escSeq = Mid(ansiData, j, InStrNextAlpha(j, ansiData, csiLastByte) - j)
         csiArgs = Mid(escSeq, 3)

         ' Advance the parser.
         j = j + Len(escSeq)

         Select Case csiLastByte

            Case "H" ' Cursor position
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
            Case "f" ' Cursor position, same as "H"
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
            'Case "2J"  ' Erase display (?)
            'Case "K"   ' Erase line (?) -- Problematic..?
            '  'wscript.echo "Clearing row " & row
            '   Dim k
            '   If Ubound(screenBuffer, 2) > 1 Then
            '      For k = 1 to Ubound(screenBuffer, 2)
            '        'wscript.echo "k " & k
            '         screenBuffer(k, row) = ""
            '      Next
            '   End If
            'Case "h"   ' Set mode (screen width/height). Probably won't bother implementing.
            'Case "l"   ' Reset mode
            'Case "p"   ' Set keyboard strings (most likely won't be implemented)
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
      
      ' Store the previous row and column before incrementing them. (not needed anymore?)
         'prevRow = row  
         'prevCol = col        

         If chrCode = 13 Then       
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
         ElseIf charCode = 10 Then  
            if row <= UBound(screenBuffer, 2) then
               screenBuffer(col, row) = newEscSeq & screenBuffer(col, row)
            end if
            row = row + 1
            col = 1
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
            
            ' Reached the end of the screen.
            If col > MAX_COLS Then
               col = 1
               row = row + 1
            End If
         End If
      End If      
      
   Next

   ' Now form new ansiData out of the screen buffer contents.
   For rBuf = 1 To UBound(screenBuffer, 2)
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
'!
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
'!
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

'! Translates an ANSI character value from code page 437 to its modern HTML
'! equivalent.
'! 
'! Could use an array for these instead, but with the huge gap between 31 
'! and 127 there would be a lot of wasted elements. Probably not much 
'! gained anyway.
'!
'! @param  ansiCharCode  The character to translate.
'! @return               The HTML symbol entity equivalent of the ANSI
'!                       character.
'!
'! @see https://en.wikipedia.org/wiki/Code_page_437
Function ToHtmlEntity(ansiCharCode)
   Select Case ansiCharCode
      Case 1
         ToHtmlEntity = "&#x263A;" ' Smiley
      Case 2
         ToHtmlEntity = "&#x263B;" ' Inverted smiley
      Case 3
         ToHtmlEntity = "&#x2665;" ' Heart
      Case 4
         ToHtmlEntity = "&#x2666;" ' Diamond
      Case 5
         ToHtmlEntity = "&#x2663;" ' Club
      Case 6
         ToHtmlEntity = "&#x2660;" ' Spade
      Case 7
         ToHtmlEntity = "&#x2022;" ' Bullet
      Case 8
         ToHtmlEntity = "&#x25D8;" ' Inverted bullet
      Case 9
         ToHtmlEntity = "&#x25CB;" ' Circle
      Case 10
         ' Also a line feed.
         ToHtmlEntity = "&#x25D9;" ' Inverted circle
      Case 11
         ToHtmlEntity = "&#x2642;" ' Male
      Case 12
         ToHtmlEntity = "&#x2640;" ' Female
      Case 13
         ' Also a carriage return.
         ToHtmlEntity = "&#x266A;" ' Eighth note
      Case 14
         ToHtmlEntity = "&#x266B;" ' Beamed eighth note
      Case 15
         ToHtmlEntity = "&#x263C;" ' Solar
      Case 16
         ToHtmlEntity = "&#x25BA;" ' Right triangle
      Case 17
         ToHtmlEntity = "&#x25C0;" ' Left triangle
      Case 18
         ToHtmlEntity = "&#x2195;" ' Up and down arrow
      Case 19
         ToHtmlEntity = "&#x203C;" ' Double bang
      Case 20
         ToHtmlEntity = "&#x00B6;" ' Paragraph
      Case 21
         ToHtmlEntity = "&#x00A7;" ' Section
      Case 22
         ToHtmlEntity = "&#x25AC;" ' Horizontal line
      Case 23
         ToHtmlEntity = "&#x21A8;" ' Up and down arrow with underscore
      Case 24
         ToHtmlEntity = "&#x2191;" ' Up arrow
      Case 25
         ToHtmlEntity = "&#x2193;" ' Down arrow
      Case 26
         ToHtmlEntity = "&#x2192;" ' Right arrow
      Case 27
         ToHtmlEntity = "&#x2190;" ' Left arrow
      Case 28
         ToHtmlEntity = "&#x221F;" ' Right angle
      Case 29
         ToHtmlEntity = "&#x2194;" ' Left and Right arrow
      Case 30
         ToHtmlEntity = "&#x25B2;" ' Up triangle
      Case 31
         ToHtmlEntity = "&#x25BC;" ' Down triangle
      Case 127
         ToHtmlEntity = "&#x2302;" ' House
      Case 128
         ToHtmlEntity = "&#x00C7;" ' Latin letter cedilla
      Case 129
         ToHtmlEntity = "&#x00FC;" ' u-umlaut
      Case 130
         ToHtmlEntity = "&#x00E9;" ' e-acute
      Case 131
         ToHtmlEntity = "&#x00E2;" ' a-circumflex
      Case 132
         ToHtmlEntity = "&#x00E4;" ' a-umlaut
      Case 133
         ToHtmlEntity = "&#x00E0;" ' a-grave
      Case 134
         ToHtmlEntity = "&#x00E5;" ' a-ring
      Case 135
         ToHtmlEntity = "&#x00E7;" ' Latin letter cedilla, lowercase
      Case 136
         ToHtmlEntity = "&#x00EA;" ' e-circumflex
      Case 137
         ToHtmlEntity = "&#x00EB;" ' e-umlaut
      Case 138
         ToHtmlEntity = "&#x00E8;" ' e-grave
      Case 139
         ToHtmlEntity = "&#x00EF;" ' i-umlaut
      Case 140
         ToHtmlEntity = "&#x00EE;" ' i-circumflex
      Case 141
         ToHtmlEntity = "&#x00EC;" ' i-grave
      Case 142
         ToHtmlEntity = "&#x00C4;" ' A-umlaut
      Case 143
         ToHtmlEntity = "&#x00C5;" ' A-ring
      Case 144
         ToHtmlEntity = "&#x00C9;" ' E-acute
      Case 145
         ToHtmlEntity = "&#x00E6;" ' lowercase aesc
      Case 146
         ToHtmlEntity = "&#x00C6;" ' uppercase AEsc
      Case 147
         ToHtmlEntity = "&#x00F4;" ' o-circumflex
      Case 148
         ToHtmlEntity = "&#x00F6;" ' o-umlaut
      Case 149
         ToHtmlEntity = "&#x00F2;" ' o-grave
      Case 150
         ToHtmlEntity = "&#x00FB;" ' u-circumflex
      Case 151
         ToHtmlEntity = "&#x00F9;" ' u-grave
      Case 152
         ToHtmlEntity = "&#x00FF;" ' y-umlaut
      Case 153
         ToHtmlEntity = "&#x00D6;" ' O-umlaut
      Case 154
         ToHtmlEntity = "&#x00DC;" ' U-umlaut
      Case 155
         ToHtmlEntity = "&#x00A2;" ' cents
      Case 156
         ToHtmlEntity = "&#x00A3;" ' British pound
      Case 157
         ToHtmlEntity = "&#x00A5;" ' yen
      Case 158
         ToHtmlEntity = "&#x20A7;" ' peseta
      Case 159
         ToHtmlEntity = "&#x0192;" ' f with hook
      Case 160
         ToHtmlEntity = "&#x00E1;" ' a-acute
      Case 161
         ToHtmlEntity = "&#x00ED;" ' i-acute
      Case 162
         ToHtmlEntity = "&#x00F3;" ' o-acute
      Case 163
         ToHtmlEntity = "&#x00FA;" ' u-acute
      Case 164
         ToHtmlEntity = "&#x00F1;" ' Spanish n (enye)
      Case 165
         ToHtmlEntity = "&#x00D1;" ' Spanish N (eNye)
      Case 166
         ToHtmlEntity = "&#x00AA;" ' ordinal a
      Case 167
         ToHtmlEntity = "&#x00BA;" ' ordinal o
      Case 168
         ToHtmlEntity = "&#x00BF;" ' inverted ?
      Case 169
         ToHtmlEntity = "&#x2310;" ' negation (left)
      Case 170
         ToHtmlEntity = "&#x00AC;" ' negation (right)
      Case 171
         ToHtmlEntity = "&#x00BD;" ' 1 half
      Case 172
         ToHtmlEntity = "&#x00BC;" ' 1 fourth
      Case 173
         ToHtmlEntity = "&#x00A1;" ' inverted !
      Case 174
         ToHtmlEntity = "&#x00AB;" ' left guillemets
      Case 175
         ToHtmlEntity = "&#x00BB;" ' right guillemets
      Case 176
         ToHtmlEntity = "&#x2591;" ' shaded block, light
      Case 177
         ToHtmlEntity = "&#x2592;" ' shaded block, medium
      Case 178
         ToHtmlEntity = "&#x2593;" ' shaded block, dark
      Case 179
         ToHtmlEntity = "&#x2502;" ' thin line, vertical
      Case 180
         ToHtmlEntity = "&#x2524;" ' thin right intersect
      Case 181
         ToHtmlEntity = "&#x2561;" ' thin double right intersect
      Case 182
         ToHtmlEntity = "&#x2562;" ' double thin right intersect
      Case 183
         ToHtmlEntity = "&#x2556;" ' thin double NE corner
      Case 184
         ToHtmlEntity = "&#x2555;" ' double thin corner
      Case 185
         ToHtmlEntity = "&#x2563;" ' double right intersect
      Case 186
         ToHtmlEntity = "&#x2551;" ' double vertical
      Case 187
         ToHtmlEntity = "&#x2557;" ' double NE corner
      Case 188
         ToHtmlEntity = "&#x255D;" ' double SE corner
      Case 189
         ToHtmlEntity = "&#x255C;" ' thin double SE corner
      Case 190
         ToHtmlEntity = "&#x255B;" ' double thin SE corner
      Case 191
         ToHtmlEntity = "&#x2510;" ' thin NE corner
      Case 192
         ToHtmlEntity = "&#x2514;" ' thin SW corner
      Case 193
         ToHtmlEntity = "&#x2534;" ' thin bottom intersect
      Case 194
         ToHtmlEntity = "&#x252C;" ' thin top intersect
      Case 195
         ToHtmlEntity = "&#x251C;" ' thin left intersect
      Case 196
         ToHtmlEntity = "&#x2500;" ' thin line horizontal
      Case 197
         ToHtmlEntity = "&#x253C;" ' thin center intersect
      Case 198
         ToHtmlEntity = "&#x255E;" ' thin double left intersect
      Case 199
         ToHtmlEntity = "&#x255F;" ' double thin left intersect
      Case 200
         ToHtmlEntity = "&#x255A;" ' double SW corner
      Case 201
         ToHtmlEntity = "&#x2554;" ' double NW corner
      Case 202
         ToHtmlEntity = "&#x2569;" ' double bottom intersect
      Case 203
         ToHtmlEntity = "&#x2566;" ' double top intersect
      Case 204
         ToHtmlEntity = "&#x2560;" ' double left intersect
      Case 205
         ToHtmlEntity = "&#x2550;" ' double line horizontal
      Case 206
         ToHtmlEntity = "&#x256C;" ' double center intersect
      Case 207
         ToHtmlEntity = "&#x2567;" ' thin double bottom intersect
      Case 208
         ToHtmlEntity = "&#x2568;" ' double thin bottom intersect
      Case 209
         ToHtmlEntity = "&#x2564;" ' double thin top intersect
      Case 210
         ToHtmlEntity = "&#x2565;" ' thin double top intersect
      Case 211
         ToHtmlEntity = "&#x2559;" ' double thin SW corner
      Case 212
         ToHtmlEntity = "&#x2558;" ' thin double SW corner
      Case 213
         ToHtmlEntity = "&#x2552;" ' thin double NW corner
      Case 214
         ToHtmlEntity = "&#x2553;" ' double thin NW corner
      Case 215
         ToHtmlEntity = "&#x256B;" ' thin double center intersect
      Case 216
         ToHtmlEntity = "&#x256A;" ' double thin center intersect
      Case 217
         ToHtmlEntity = "&#x2518;" ' thin SE corner
      Case 218
         ToHtmlEntity = "&#x250C;" ' thin NW corner
      Case 219
         ToHtmlEntity = "&#x2588;" ' solid block
      Case 220
         ToHtmlEntity = "&#x2584;" ' bottom half block
      Case 221
         ToHtmlEntity = "&#x258C;" ' left half block
      Case 222
         ToHtmlEntity = "&#x2590;" ' right half block
      Case 223
         ToHtmlEntity = "&#x2580;" ' top half block
      Case 224
         ToHtmlEntity = "&#x03B1;" ' alpha
      Case 225
         'ToHtmlEntity = "&#x03B2;" ' Beta
         ToHtmlEntity = "&#x00DF;" ' Eszett
      Case 226
         ToHtmlEntity = "&#x0393;" ' gamma
      Case 227
         ToHtmlEntity = "&#x03C0;" ' pi
      Case 228
         ToHtmlEntity = "&#x03A3;" ' sigma uppercase
      Case 229
         ToHtmlEntity = "&#x03C3;" ' sigma lowercase
      Case 230
         'ToHtmlEntity = "&#x03BC;"
         ToHtmlEntity = "&#x00B5;" ' mu
      Case 231
         ToHtmlEntity = "&#x03C4;" ' tau
      Case 232
         'ToHtmlEntity = "&#x0424;"
         ToHtmlEntity = "&#x03A6;" ' phi
      Case 233
         ToHtmlEntity = "&#x0398;" ' theta
      Case 234
         ToHtmlEntity = "&#x03A9;" ' Omega
      Case 235
         ToHtmlEntity = "&#x03B4;" ' Delta
      Case 236
         ToHtmlEntity = "&#x221E;" ' infinity
      Case 237
         'ToHtmlEntity = "&#x0444;"
         ToHtmlEntity = "&#x03C6;" ' Phi
      Case 238
         'ToHtmlEntity = "&#x0152;"
         ToHtmlEntity = "&#x03B5;" ' Epsilon
      Case 239
         'ToHtmlEntity = "&#x22C2;"
         ToHtmlEntity = "&#x2229;" ' intersection
      Case 240
         ToHtmlEntity = "&#x2261;" ' triple bar
      Case 241
         'ToHtmlEntity = "&#x2213;"
         ToHtmlEntity = "&#x00B1;" ' plus minus
      Case 242
         ToHtmlEntity = "&#x2265;" ' greater or equal to
      Case 243
         ToHtmlEntity = "&#x2264;" ' less or equal to
      Case 244
         'ToHtmlEntity = "&#x256D;"
         ToHtmlEntity = "&#x2320;" ' top integral
      Case 245
         'ToHtmlEntity = "&#x256F;"
         ToHtmlEntity = "&#x2321;" ' bottom integral
      Case 246
         ToHtmlEntity = "&#x00F7;" ' obelus (division)
      Case 247
         ToHtmlEntity = "&#x2248;" ' approximation
      Case 248
         ToHtmlEntity = "&#x00B0;" ' degree
      Case 249
         'ToHtmlEntity = "&#x2022;"
         ToHtmlEntity = "&#x2219;" ' bullet
      Case 250
         ToHtmlEntity = "&#x00B7;" ' interpunct
      Case 251
         ToHtmlEntity = "&#x221A;" ' square root / check mark
      Case 252
         ToHtmlEntity = "&#x207F;" ' ordinal n
      Case 253
         ToHtmlEntity = "&#x00B2;" ' squared (raised 2)
      Case 254
         ToHtmlEntity = "&#x25A0;" ' small block
   End Select
End Function
