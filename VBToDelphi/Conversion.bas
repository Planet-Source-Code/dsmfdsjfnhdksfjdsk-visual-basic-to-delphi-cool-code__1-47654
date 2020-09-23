Attribute VB_Name = "Conversion"
Rem ***********************************
Rem Conversion program to change Visual Basic code modules
Rem to Delphi - by Tim Gathercole 8/8/96
Rem ***********************************
Global inpflag%, commentflag%, dblwrdflag%, delfag%, mainflag%
Global filename
Global varcase$, varcaseflag%

Function Exists%(F$)
On Error Resume Next
X& = FileLen(F$)
If X& Then Exists% = True
End Function

Sub findCASEvar(l$)
 Nindex% = InStr(l$, "SELECT CASE")
 l$ = Trim(Right$(l$, Len(l$) - Nindex% - 11))
End Sub



Sub looksee(bl$)
Rem pre parse check & breakdown
Dim tempc As Integer
Dim semic As String

On Error Resume Next

Rem fix for THEN followed by command
 If InStr(bl$, "THEN") Then
   temp$ = Right$(bl$, (Len(bl$) - InStr(bl$, "THEN") - 3))
   If Len(Trim(temp$)) > 1 And InStr(temp$, "'") = 0 And InStr(temp$, "REM") = 0 Then
     temp2$ = Left$(bl$, InStr(bl$, "THEN") + 3)
     bl$ = temp2$ + Chr$(13) + Chr$(10) + temp$ + Chr$(13) + Chr$(10) + "End; // IF"
   End If
 End If

Rem *** WHILE  Preporsesor for the change in Psionit
  If InStr(bl$, "WHILE") Then
    bl$ = bl$ + " Do" + Chr$(13) + Chr$(10) + "Begin"
  End If

Rem *** PRINT # - needs an APPEND. Preporsesor for the change in Psionit
 ' If InStr(bl$, "PRINT #") Then
 '   bl$ = bl$ + Chr$(13) + Chr$(10) + "APPEND"
 ' End If

Rem *** INPUT # - needs a NEXT. Preporsesor for the change in Psionit
 ' If InStr(bl$, "INPUT #") Then
 '   bl$ = bl$ + Chr$(13) + Chr$(10) + "NEXT"
 ' End If

Rem is there a quote
 If InStr(bl$, Chr$(34)) Then
    quoteflag% = 1
    inpflag% = 0 ' reset
    Call NoofQuote(bl$, c%)
    If c% / 2 <> Int(c% / 2) Then
       ' rem it something wrong with line
       bl$ = "REM No. quotes>>> " + bl$
       GoTo finish
    End If
    Call quote(bl$, lstr$, lend$, lmid$)
    Call psionit(lstr$)
    Call psionit(lend$)
    bl$ = lstr$ + lmid$ + lend$
    GoTo finish
  End If

'  if instr(bl$,"'") then commentflag%=1
  If InStr(bl$, "EXIT DO") Then dblwrdflag% = 1
  If InStr(bl$, "EXIT SUB") Then dblwrdflag% = 1

  Call psionit(bl$)

Rem **** Write the translated line ****
finish:
Rem *** Look for The other end of a needed bracket ***
    braketplus = "(" + Chr$(34)
  If InStr(bl$, braketplus) Then
    Itsat = InStr(InStr(bl$, Chr$(34)) + 1, bl$, Chr$(34))
    p1$ = Left(bl$, Itsat)
    p2$ = Right(bl$, Len(l$) - Itsat)
    bl$ = p1$ + ");" + p2$
  End If
 
 If delflag% = 1 Then
   delflag% = 0
   Exit Sub
 End If
 
 Rem *** look for " & put in ' ***
 For tempc = 1 To Len(bl$)
   If Mid$(bl$, tempc, 1) = Chr$(34) Then Mid(bl$, tempc) = "'"
 Next tempc
 
 If InStr(UCase$(bl$), "IF ") Or InStr(UCase$(bl$), "ELSE ") Or InStr(UCase$(bl$), "BEGIN") Then
   semic = ""
 ElseIf Trim(bl$) = "" Or Right(Trim(bl$), 1) = ":" Or Right(Trim(bl$), 1) = ";" Then
   semic = ""
 ElseIf InStr(UCase$(bl$), "Case ") And InStr(UCase$(bl$), " Of") Then
   semic = ""
 Else
    semic = ";"
 End If
 
 If InStr(UCase$(bl$), " //") Then
   
 End If
 
 Rem *** OK Write it ***********
 Print #2, bl$ + semic

End Sub

Sub NoofQuote(bl$, c%)
  n% = Len(bl$)
  For i% = 1 To n%
    temp$ = Mid$(bl$, i%, 1)
    If temp$ = Chr$(34) Then c% = c% + 1
  Next i%
End Sub

Sub not_necessary()
Rem *** ABS Function OK ***
Rem *** AND Operator ***
Rem *** ASC Function ***
Rem *** CHR$ Function ***
Rem *** CLS Statement ***
Rem *** COS Function ***
Rem --- EOF Function ---
Rem *** ELSE Keyword ***
Rem *** ELSEIF Keyword ***
Rem *** ERR Function ***
Rem *** GOTO Statement ***
Rem *** HEX$ Function ***
Rem *** INT Function ***
Rem *** LEFT$ Function ***
Rem *** LEN Function ***
Rem *** LPRINT Statement ***
Rem *** MID$ Function ***
Rem *** MKDIR Statement ***
Rem *** NOT Operator ***
Rem *** OCT$ Function ***
Rem *** OR Operator ***
Rem *** PRINT Statement ***
Rem *** REM Statement ***
Rem *** RIGHT$ Function ***
Rem *** RMDIR Statement ***
Rem *** RND Function ***
Rem *** SIN Function ***
Rem *** TAN Function ***
Rem *** VAL Function ***

End Sub

Sub notimplimented()
Rem *** TG 12/9/96 ***

Rem --- ABSOLUTE Keyword X ---
  Keyword$ = "ABSOLUTE": newkey$ = "REM * ABSOLUTE"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- ACCESS Keyword ---
  Keyword$ = "ACCESS": newkey$ = "REM * ACCESS"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- ANY Keyword ---
  Keyword$ = "ANY": newkey$ = "REM * ANY"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- APPEND Keyword ---
  Keyword$ = "APPEND": newkey$ = "REM * APPEND"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- AS Keyword ---
  Keyword$ = "AS": newkey$ = "REM * AS"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- BASE Keyword ---
  Keyword$ = "BASE": newkey$ = "REM * BASE"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- BINARY Keyword ---
  Keyword$ = "BINARY": newkey$ = "REM * BINARY"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- BLOAD Statement ---
  Keyword$ = "BLOAD": newkey$ = "REM * BLOAD"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- BSAVE Statement ---
  Keyword$ = "BSAVE": newkey$ = "REM * BSAVE"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- CALL ABSOLUTE Statement ---
  Keyword$ = "CALL": newkey$ = "REM * CALL"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- CDBL Function ---
  Keyword$ = "CDBL": newkey$ = "REM * CDBL"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- CHAIN Statement ---
  Keyword$ = "CHAIN": newkey$ = "REM * CHAIN"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- CHDIR Statement ---
  Keyword$ = "CHDIR": newkey$ = "REM * CHDIR"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- CINT Function ---
  Keyword$ = "CINT": newkey$ = "REM * CINT"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- CIRCLE Statement ---
  Keyword$ = "CIRCLE": newkey$ = "REM * CIRCLE"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- CLEAR Statement ---
  Keyword$ = "CLEAR": newkey$ = "REM * CLEAR"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- CLNG Function ---
  Keyword$ = "CLNG": newkey$ = "REM * CLNG"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- CLOSE Statement ---
  Keyword$ = "CLOSE": newkey$ = "REM * CLOSE"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- COLOR Statement ---
  Keyword$ = "COLOR": newkey$ = "REM * COLOR"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- COM Statement ---
  Keyword$ = "COM": newkey$ = "REM * COM"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- CONST Statement ---
  Keyword$ = "CONST": newkey$ = "REM * CONST"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- CSNG Function ---
  Keyword$ = "CSNG": newkey$ = "REM * CSNG"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- CSRLIN Function ---
  Keyword$ = "CSRLIN": newkey$ = "REM * CSRLIN"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- CVD Function ---
  Keyword$ = "CVD": newkey$ = "REM * CVD"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- CVDMBF Function ---
  Keyword$ = "CVDMBF": newkey$ = "REM * CVDMBF"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- CVI Function ---
  Keyword$ = "CVI": newkey$ = "REM * CVI"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- CVL Function ---
  Keyword$ = "CVL": newkey$ = "REM * CVL"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- CVS Function ---
  Keyword$ = "CVS": newkey$ = "REM * CVS"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- CVSMBF Function ---
  Keyword$ = "CVSMBF": newkey$ = "REM * CVSMBF"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- Data Type Keywords ---
  Keyword$ = "Data": newkey$ = "REM * Data"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- DATE$ Function ---
  Keyword$ = "DATE$": newkey$ = "REM * DATE$"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- DATE$ Statement ---
  Keyword$ = "DATE$": newkey$ = "REM * DATE$"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- DEF FN Statement ---
  Keyword$ = "DEF": newkey$ = "REM * DEF"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- DEF SEG Statement ---
  Keyword$ = "DEF": newkey$ = "REM * DEF"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- DEFDBL Statement ---
  Keyword$ = "DEFDBL": newkey$ = "REM * DEFDBL"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- DEFINT Statement ---
  Keyword$ = "DEFINT": newkey$ = "REM * DEFINT"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- DEFLNG Statement ---
  Keyword$ = "DEFLNG": newkey$ = "REM * DEFLNG"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- DEFSNG Statement ---
  Keyword$ = "DEFSNG": newkey$ = "REM * DEFSNG"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- DEFSTR Statement ---
  Keyword$ = "DEFSTR": newkey$ = "REM * DEFSTR"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- DOUBLE Keyword ---
  Keyword$ = "DOUBLE": newkey$ = "REM * DOUBLE"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- DRAW Statement ---
  Keyword$ = "DRAW": newkey$ = "REM * DRAW"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- $DYNAMIC Metacommand ---
  Keyword$ = "$DYNAMIC": newkey$ = "REM * $DYNAMIC"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- ENVIRON Statement ---
  Keyword$ = "ENVIRON": newkey$ = "REM * ENVIRON"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- ENVIRON$ Function ---
  Keyword$ = "ENVIRON$": newkey$ = "REM * ENVIRON$"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- EQV Operator ---
  Keyword$ = "EQV": newkey$ = "REM * EQV"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- ERASE Statement ---
  Keyword$ = "ERASE": newkey$ = "REM * ERASE"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- ERDEV Function ---
  Keyword$ = "ERDEV": newkey$ = "REM * ERDEV"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- ERDEV$ Function ---
  Keyword$ = "ERDE": newkey$ = "REM * ERDE"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- ERL Function ---
  Keyword$ = "ERL": newkey$ = "REM * ERL"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- ERROR Statement ---
  Keyword$ = "ERROR": newkey$ = "REM * ERROR"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- EXP Function ---
  Keyword$ = "EXP": newkey$ = "REM * EXP"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- FIELD Statement ---
  Keyword$ = "FIELD": newkey$ = "REM * FIELD"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- FILEATTR Function ---
  Keyword$ = "FILEATTR": newkey$ = "REM * FILEATTR"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- FILES Statement ---
  Keyword$ = "FILES": newkey$ = "REM * FILES"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- FIX Function ---
  Keyword$ = "FIX": newkey$ = "REM * FIX"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- FREEFILE Function ---
  Keyword$ = "FREEFILE": newkey$ = "REM * FREEFILE"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- FUNCTION Statement ---
  Keyword$ = "FUNCTION": newkey$ = "REM * FUNCTION"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- GET (File I/O) Statement ---
  Keyword$ = "GET": newkey$ = "REM * GET"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- GET (Graphics) Statement ---
  Keyword$ = "GET": newkey$ = "REM * GET"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- IMP Operator ---
  Keyword$ = "IMP": newkey$ = "REM * IMP"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- INP Function ---
  Keyword$ = "INP": newkey$ = "REM * INP"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- INPUT$ Function ---
  Keyword$ = "INPUT$": newkey$ = "REM * INPUT$"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- INTEGER Keyword ---
  Keyword$ = "INTEGER": newkey$ = "REM * INTEGER"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- IOCTL Statement ---
  Keyword$ = "IOCTL": newkey$ = "REM * IOCTL"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- IOCTL$ Function ---
  Keyword$ = "IOCTL$": newkey$ = "REM * IOCTL$"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- IS Keyword ---
  Keyword$ = "IS": newkey$ = "REM * IS"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- KEY (Assignment) Statement ---
  Keyword$ = "KEY": newkey$ = "REM * KEY"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- KEY (Event Trapping) Statement ---
  Keyword$ = "KEY": newkey$ = "REM * KEY"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- LBOUND Function ---
  Keyword$ = "LBOUND": newkey$ = "REM * LBOUND"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- LINE (Graphics) Statement ---
  Keyword$ = "LINE": newkey$ = "REM * LINE"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- LIST ---
  Keyword$ = "LIST": newkey$ = "REM * LIST"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- LOC Function ---
  Keyword$ = "LOC": newkey$ = "REM * LOC"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- LOCK...UNLOCK Statements ---
  Keyword$ = "LOC": newkey$ = "REM * LOC"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- LOF Function ---
  Keyword$ = "LOF": newkey$ = "REM * LOF"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- LOG Function ---
  Keyword$ = "LOG": newkey$ = "REM * LOG"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- LONG Keyword ---
  Keyword$ = "LONG": newkey$ = "REM * LONG"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- LPOS Function ---
  Keyword$ = "LPOS": newkey$ = "REM * LPOS"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- LPRINT USING Statement ---
  Keyword$ = "LPRINT": newkey$ = "REM * LPRINT"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- LSET Statement ---
  Keyword$ = "LSET": newkey$ = "REM * LSET"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- MKD$ Function ---
  Keyword$ = "MKD$": newkey$ = "REM * MKD$"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- MKDMBF$ Function ---
  Keyword$ = "MKDMBF$": newkey$ = "REM * MKDMBF$"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- MKI$ Function ---
  Keyword$ = "MKI$": newkey$ = "REM * MKI$"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- MKL$ Function ---
  Keyword$ = "MKL$": newkey$ = "REM * MKL$"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- MKS$ Function ---
  Keyword$ = "MKS$": newkey$ = "REM * MKS$"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- MKSMBF$ Function ---
  Keyword$ = "MKSMBF$": newkey$ = "REM * MKSMBF$"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- MOD Operator ---
  Keyword$ = "MOD": newkey$ = "REM * MOD"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- NAME Statement ---
  Keyword$ = "NAME": newkey$ = "REM * NAME"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- OFF Keyword ---
  Keyword$ = "OFF": newkey$ = "REM * OFF"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- ON COM Statement ---
  Keyword$ = "ON": newkey$ = "REM * ON"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- ON Keyword ---
  Keyword$ = "ON": newkey$ = "REM * ON"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- ON KEY Statement ---
  Keyword$ = "ON": newkey$ = "REM * ON"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- ON PEN Statement ---
  Keyword$ = "ON": newkey$ = "REM * ON"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- ON PLAY Statement ---
  Keyword$ = "ON": newkey$ = "REM * ON"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- ON STRIG Statement ---
  Keyword$ = "ON": newkey$ = "REM * ON"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- ON TIMER Statement ---
  Keyword$ = "ON": newkey$ = "REM * ON"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- ON...GOSUB Statement ---
  Keyword$ = "O": newkey$ = "REM * O"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- ON...GOTO Statement ---
  Keyword$ = "O": newkey$ = "REM * O"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- OPEN COM Statement ---
  Keyword$ = "OPEN": newkey$ = "REM * OPEN"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- OPTION BASE Statement ---
  Keyword$ = "OPTION": newkey$ = "REM * OPTION"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- OUT Statement ---
  Keyword$ = "OUT": newkey$ = "REM * OUT"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- OUTPUT Keyword ---
  Keyword$ = "OUTPUT": newkey$ = "REM * OUTPUT"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- PAINT Statement ---
  Keyword$ = "PAINT": newkey$ = "REM * PAINT"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- PALETTE Statements ---
  Keyword$ = "PALETTE": newkey$ = "REM * PALETTE"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- PCOPY Statement ---
  Keyword$ = "PCOPY": newkey$ = "REM * PCOPY"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- PEEK Function ---
  Keyword$ = "PEEK": newkey$ = "REM * PEEK"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- PEN Function ---
  Keyword$ = "PEN": newkey$ = "REM * PEN"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- PEN Statement ---
  Keyword$ = "PEN": newkey$ = "REM * PEN"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- PLAY Function ---
  Keyword$ = "PLAY": newkey$ = "REM * PLAY"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- PLAY (Music) Statement ---
  Keyword$ = "PLAY": newkey$ = "REM * PLAY"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- PLAY (Event Trapping) Statements ---
  Keyword$ = "PLAY": newkey$ = "REM * PLAY"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- PMAP Function ---
  Keyword$ = "PMAP": newkey$ = "REM * PMAP"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- POINT Function ---
  Keyword$ = "POINT": newkey$ = "REM * POINT"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- POKE Statement ---
  Keyword$ = "POKE": newkey$ = "REM * POKE"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- POS Function ---
  Keyword$ = "POS": newkey$ = "REM * POS"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- PRESET Statement ---
  Keyword$ = "PRESET": newkey$ = "REM * PRESET"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- PRINT USING Statement ---
  Keyword$ = "PRINT": newkey$ = "REM * PRINT"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- PSET Statement ---
  Keyword$ = "PSET": newkey$ = "REM * PSET"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- PUT (File I/O) Statement ---
  Keyword$ = "PUT": newkey$ = "REM * PUT"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- PUT (Graphics) Statement ---
  Keyword$ = "PUT": newkey$ = "REM * PUT"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- RANDOM Keyword ---
  Keyword$ = "RANDOM": newkey$ = "REM * RANDOM"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- READ Statement ---
  Keyword$ = "READ": newkey$ = "REM * READ"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- REDIM Statement ---
  Keyword$ = "REDIM": newkey$ = "REM * REDIM"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- RESET Statement ---
  Keyword$ = "RESET": newkey$ = "REM * RESET"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- RESTORE Statement ---
  Keyword$ = "RESTORE": newkey$ = "REM * RESTORE"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- RESUME Statement ---
  Keyword$ = "RESUME": newkey$ = "REM * RESUME"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- RETURN Statement ---
  Keyword$ = "RETURN": newkey$ = "REM * RETURN"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- RSET Statement ---
  Keyword$ = "RSET": newkey$ = "REM * RSET"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- RUN Statement ---
  Keyword$ = "RUN": newkey$ = "REM * RUN"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- SCREEN Function ---
  Keyword$ = "SCREEN": newkey$ = "REM * SCREEN"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- SCREEN Statement ---
  Keyword$ = "SCREEN": newkey$ = "REM * SCREEN"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- SEEK Function ---
  Keyword$ = "SEEK": newkey$ = "REM * SEEK"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- SEEK Statement ---
  Keyword$ = "SEEK": newkey$ = "REM * SEEK"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- SGN Function ---
  Keyword$ = "SGN": newkey$ = "REM * SGN"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- SHARED Statement ---
  Keyword$ = "SHARED": newkey$ = "REM * SHARED"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- SHELL Statement ---
  Keyword$ = "SHELL": newkey$ = "REM * SHELL"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- SINGLE Keyword ---
  Keyword$ = "SINGLE": newkey$ = "REM * SINGLE"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- SPACE$ Function ---
  Keyword$ = "SPACE$": newkey$ = "REM * SPACE$"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- SPC Function ---
  Keyword$ = "SPC": newkey$ = "REM * SPC"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- SQR Function ---
  Keyword$ = "SQR": newkey$ = "REM * SQR"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- STATIC Statement ---
  Keyword$ = "STATIC": newkey$ = "REM * STATIC"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- $STATIC Metacommand ---
  Keyword$ = "$STATIC": newkey$ = "REM * $STATIC"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- STICK Function ---
  Keyword$ = "STICK": newkey$ = "REM * STICK"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- STOP Statement ---
  Keyword$ = "STOP": newkey$ = "REM * STOP"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- STR$ Function ---
  Keyword$ = "STR$": newkey$ = "REM * STR$"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- STRIG Function ---
  Keyword$ = "STRIG": newkey$ = "REM * STRIG"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- STRIG Statements ---
  Keyword$ = "STRIG": newkey$ = "REM * STRIG"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- STRING Keyword ---
  Keyword$ = "STRING": newkey$ = "REM * STRING"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- STRING$ Function ---
  Keyword$ = "STRING$": newkey$ = "REM * STRING$"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- SWAP Statement ---
  Keyword$ = "SWAP": newkey$ = "REM * SWAP"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- TAB Function ---
  Keyword$ = "TAB": newkey$ = "REM * TAB"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- TIME$ Function ---
  Keyword$ = "TIME$": newkey$ = "REM * TIME$"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- TIME$ Statement ---
  Keyword$ = "TIME$": newkey$ = "REM * TIME$"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- TIMER Function ---
  Keyword$ = "TIMER": newkey$ = "REM * TIMER"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- TIMER Statements ---
  Keyword$ = "TIMER": newkey$ = "REM * TIMER"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- TO Keyword ---
  Keyword$ = "TO": newkey$ = "REM * TO"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- TROFF Statement ---
  Keyword$ = "TROFF": newkey$ = "REM * TROFF"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- TRON Statement ---
  Keyword$ = "TRON": newkey$ = "REM * TRON"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- TYPE Statement ---
  Keyword$ = "TYPE": newkey$ = "REM * TYPE"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- UBOUND Function ---
  Keyword$ = "UBOUND": newkey$ = "REM * UBOUND"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- UNLOCK Statement ---
  Keyword$ = "UNLOCK": newkey$ = "REM * UNLOCK"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- USING Keyword ---
  Keyword$ = "USING": newkey$ = "REM * USING"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- VARPTR Function ---
  Keyword$ = "VARPTR": newkey$ = "REM * VARPTR"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- VARPTR$ Function ---
  Keyword$ = "VARPTR$": newkey$ = "REM * VARPTR$"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- VARSEG Function ---
  Keyword$ = "VARSEG": newkey$ = "REM * VARSEG"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- VIEW Statement ---
  Keyword$ = "VIEW": newkey$ = "REM * VIEW"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- VIEW PRINT Statement ---
  Keyword$ = "VIEW": newkey$ = "REM * VIEW"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- WAIT Statement ---
  Keyword$ = "WAIT": newkey$ = "REM * WAIT"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- WIDTH Statements ---
  Keyword$ = "WIDTH": newkey$ = "REM * WIDTH"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- WINDOW Statement ---
  Keyword$ = "WINDOW": newkey$ = "REM * WINDOW"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- WRITE Statement ---
  Keyword$ = "WRITE": newkey$ = "REM * WRITE"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- XOR Operator ---
  Keyword$ = "XOR": newkey$ = "REM * XOR"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

End Sub

Sub parse(l$, Keyword$, newkey$)
Rem parse line
  llen% = Len(l$)
    position% = InStr(l$, Keyword$)
  lstart$ = Left$(l$, position% - 1)
  lend$ = Right$(l$, llen% - (position% + Len(Keyword$) - 1))

  l$ = lstart$ + newkey$ + lend$

End Sub

Sub psionit(l$)
  On Error Resume Next

  l$ = RTrim$(l$) ' delete trailing spaces
    
'Rem ** deal with labels **
'  If Right$(l$, 1) = ":" Then
'    l$ = l$ + ":"
'    GoTo writeit ' labels are on their own line
'  End If

'powrlook:
'Rem ** ^ (Power sign) **
'  Keyword$ = "^": newkey$ = "**"
'  If InStr(l$, Keyword$) > 0 Then
'    Call parse(l$, Keyword$, newkey$)
'    GoTo powrlook ' could be multiple ^ in a formula
'  End If

Rem -- ' (comment) **
  Keyword$ = "'": newkey$ = " // "
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem -- REM if line starts with REM then : not needed **
  Keyword$ = ":REM": newkey$ = "// "
  If Left$(LTrim(l$), 4) = "//" Then Call parse(l$, Keyword$, newkey$)
  If Left$(LTrim(l$), 3) = "//" Then GoTo writeit ' not needed

Rem --- REM (comment) ---
  Keyword$ = "REM": newkey$ = " // "
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)
  
Rem --- = ---
  Keyword$ = "=": newkey$ = ":="
  If InStr(l$, Keyword$) > 0 And InStr(l$, "IF") = 0 Then Call parse(l$, Keyword$, newkey$)
    Keyword$ = ">:=": newkey$ = ">=" ' easier just to fix
    If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)
    Keyword$ = "<:=": newkey$ = "<=" ' easier just to fix
    If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- AS Function ---
  Keyword$ = " AS ": newkey$ = " : "
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- ATN Function ---
  Keyword$ = "ATN": newkey$ = "ArcTan"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem *** CALL Statement ***
  Keyword$ = "CALL": newkey$ = ""
  If InStr(l$, Keyword$) > 0 Then
    Call parse(l$, Keyword$, newkey$)
     ' If InStr(l$, "(") Then
     '   temp$ = Left$(l$, (InStr(l$, "(")) - 1)
     '   l$ = RTrim$(temp$) + ":" + Chr$(13) + Chr$(10) + "REM " + Right$(l$, (Len(l$) - InStr(l$, "(") + 1))
     ' Else
     '   l$ = l$ + ":"
     ' End If
  End If

Rem --- CLOSE Function ---
  Keyword$ = "CLOSE": newkey$ = "CloseFile("
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem *** COMMON Statement ***
  Keyword$ = "COMMON SHARED": newkey$ = "// COMMON SHARED"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- DATA Statement ---
  Keyword$ = "DATA": newkey$ = "// * DATA"
  If InStr(l$, Keyword$) > 0 Then
    Call parse(l$, Keyword$, newkey$)
    GoTo writeit ' no further req
  End If

Rem *** DIM Statement ***
  Keyword$ = "DIM": newkey$ = "Var >> "
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem *** DECLARE Statement ***
  Keyword$ = "DECLARE": newkey$ = "// DECLARE"
  If InStr(l$, Keyword$) > 0 Then
    delfag% = 1
    Call parse(l$, Keyword$, newkey$)
    GoTo writeit ' dont need other checks on line
  End If

Rem *** DO Statement ***
  Keyword$ = "DO UNTIL": newkey$ = "DO :REM * "
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

  Keyword$ = "DO WHILE": newkey$ = "WHILE"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem ** ELSEIF **
  Keyword$ = "ELSEIF": newkey$ = "End;" + Chr$(13) + Chr$(10) + "IF "
  If InStr(l$, Keyword$) > 0 Then
    Call parse(l$, Keyword$, newkey$)
    GoTo writeit ' dont need other checks on line
  End If

Rem ** END IF **
  Keyword$ = "END IF": newkey$ = "End; // If"
  If InStr(l$, Keyword$) > 0 Then
    Call parse(l$, Keyword$, newkey$)
    GoTo writeit ' dont need other checks on line
  End If

Rem *** WEND Keyword ***
  Keyword$ = "WEND": newkey$ = "End; // While"
  If InStr(l$, Keyword$) > 0 Then
    Call parse(l$, Keyword$, newkey$)
    GoTo writeit ' dont need other checks on line
  End If

Rem ** END SUB ** ++++++
  Keyword$ = "END SUB": newkey$ = Chr$(13) + Chr$(10) + "End; // Proc"
  If InStr(l$, Keyword$) > 0 Then
    Call parse(l$, Keyword$, newkey$)
    GoTo writeit ' dont need other checks on line
  End If

Rem ** END SELECT ++++++
  Keyword$ = "END SELECT": newkey$ = "End; // Case"
  If InStr(l$, Keyword$) > 0 Then
    Call parse(l$, Keyword$, newkey$)
    GoTo writeit ' dont need other checks on line
  End If

Rem ** CASE ELSE ++++++
  Keyword$ = "CASE ELSE": newkey$ = "Else"
  If InStr(l$, Keyword$) > 0 Then
    Call parse(l$, Keyword$, newkey$)
    GoTo writeit ' dont need other checks on line
  End If

Rem *** END Statement ***
Rem (Many commands have END in them do this one last) **
  Keyword$ = "END": newkey$ = "Application.Terminate;"
  If InStr(l$, Keyword$) And InStr(l$, "APPEND") = 0 Then ' no mistakes
    Call parse(l$, Keyword$, newkey$)
  End If

Rem *** EXIT Statement *** ++++++
  Keyword$ = "EXIT DO": newkey$ = "Exit"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem *** FileLen( Statement ***
Keyword$ = "FileLen(": newkey$ = "FileSize( // "
If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem *** FOR... Statement ***
  If InStr(l$, "FOR") > 0 And InStr(l$, "=") > 0 Then
     l$ = l$ + " Do Begin"
  End If

Rem *** FRE Function ***
  Keyword$ = "FRE": newkey$ = "SPACE"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem *** IF...THEN...ELSE Statement *** ++++++
  Keyword$ = "THEN": newkey$ = "Then Begin"
  If InStr(l$, Keyword$) > 0 Then ' single line commands dealt with in looksee
    Call parse(l$, Keyword$, newkey$)
  End If

Rem *** INKEY$ Function ***
  Keyword$ = "INKEY$": newkey$ = "KEY$"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- KILL ---
  Keyword$ = "KILL": newkey$ = "DeleteFile"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem *** LINE INPUT Statement ***
  Keyword$ = "LINE INPUT": newkey$ = "INPUT"
  If InStr(l$, Keyword$) > 0 Then
    Call parse(l$, Keyword$, newkey$)
    GoTo writeit ' dont need other checks on line
  End If

Rem *** PRINT # - seems linked to below *** ????????
  Keyword$ = "PRINT #": newkey$ = "Writeln(" ' partially dealt with in looksee
  If InStr(l$, Keyword$) > 0 Then
    Call parse(l$, Keyword$, newkey$)
    GoTo writeit ' dont need other checks on line
  End If

Rem *** INPUT # - seems linked to below *** ????????
  Keyword$ = "INPUT #": newkey$ = "Read(" ' partially dealt with in looksee
  If InStr(l$, Keyword$) > 0 Then
    Call parse(l$, Keyword$, newkey$)
    GoTo writeit ' dont need other checks on line
  End If

Rem *** INPUT Statement *** ????
  Keyword$ = "INPUT": newkey$ = "PRINT "
  If InStr(l$, Keyword$) And InStr(l$, "FOR INPUT ") = 0 Then
    Call parse(l$, Keyword$, newkey$)
    inpflag% = 1
  End If

Rem *** INPUTBOX Statement ***
  Keyword$ = "InputBox$": newkey$ = "InputBox"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem *** INSTR Function ***
  Keyword$ = "INSTR": newkey$ = "Pos"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem *** KILL Statement ***
  Keyword$ = "KILL": newkey$ = "DeleteFile("
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem *** LCASE$ Function ***
  Keyword$ = "LCASE$": newkey$ = "LowerCase"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem *** LCASE$ Function ***
  Keyword$ = "LEN": newkey$ = "Length"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem *** LOCATE Statement ***
  Keyword$ = "LOCATE": newkey$ = "AT"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem *** LOOP + Keyword ***
  Keyword$ = "LOOP WHILE": newkey$ = "END; // WHILE "
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

  Keyword$ = "LOOP UNTIL": newkey$ = "END; // While/Do"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- LTRIM$ Function ---
  Keyword$ = "LTRIM$": newkey$ = "TrimLeft"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem *** MsgBox Function ***
  Keyword$ = "MsgBox": newkey$ = "ShowMessage("
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- MID$ Statement ---
  Keyword$ = "MID$": newkey$ = "Copy"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)
  Keyword$ = "Mid": newkey$ = "Copy"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem *** NEXT Keyword ***
   Keyword$ = "NEXT": newkey$ = "End; // For"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem *** ON ERROR Statement ***
  Keyword$ = "ON ERROR": newkey$ = "ONERR"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem *** OPEN Statement ***
  Keyword$ = "OPEN": newkey$ = "FileOpen("
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)
   ' llen% = Len(l$)
   '   position% = InStr(l$, "FOR")
   ' If position% > 1 Then
   '   lstart$ = Left$(l$, position% - 1)
   ' Else
   '   lstart$ = l$
   ' End If
   ' l$ = "REM >> " + lstart$ + ",A,??? fill fields in by hand"

'      If InStr(l$, "OUTPUT") Then ' do we need to create?
'        Keyword$ = "OPEN": newkey$ = "CREATE"
'        Call parse(l$, Keyword$, newkey$)
'      End If
   ' GoTo writeit ' dont need other checks on line
 ' End If

Rem *** Help for the above ***
 ' If InStr(l$, "FOR OUTPUT AS") Or InStr(l$, "FOR APPEND AS") Then
 '   ' problem with the line see reM - must be after 1st write.
 '   l$ = l$ + Chr$(13) + Chr$(10) + "REM IF NOT EXIST(filename)" + Chr$(13) + Chr$(10) + "REM   CREATE filename, A, var$" + Chr$(13) + Chr$(10) + "REM Else" + Chr$(13) + Chr$(10) + "REM   OPEN filename,A,var$ " + Chr$(13) + Chr$(10) + " REM ENDIF"
 '   GoTo writeit ' dont need other checks on line
 ' End If

Rem --- PRINT Statement ---
  Keyword$ = "PRINT": newkey$ = "Write("
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)
  
Rem --- RIGHT$ Function ---
  Keyword$ = "RIGHT$": newkey$ = "Copy"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)
  Keyword$ = "Right": newkey$ = "Copy"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- RTRIM$ Function ---
  Keyword$ = "RTRIM$": newkey$ = "TrimRight"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem *** SELECT CASE Statement ***
  Rem >>>>> Needs work
  Keyword$ = "SELECT CASE"
  If InStr(l$, Keyword$) > 0 Then
    Call findCASEvar(l$)
      l$ = "Case " + l$ + " Of"
    GoTo writeit ' dont need other checks on line
  End If

Rem *** SLEEP Statement ***
  Keyword$ = "SLEEP": newkey$ = "PAUSE"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem *** STEP Keyword ***
  Keyword$ = "STEP": newkey$ = " // * STEP"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem ** EXIT SUB ** ++++++
  Keyword$ = "EXIT SUB": newkey$ = "Exit"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem *** GOSUB Statement ***
  Keyword$ = "GOSUB": newkey$ = ""
  If InStr(l$, Keyword$) > 0 Then
    Call parse(l$, Keyword$, newkey$)
    GoTo writeit
  End If

Rem --- SOUND Statement ---
  Keyword$ = "SOUND": newkey$ = "BEEP"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem --- Private Sub Statement ---
  Keyword$ = "Private Sub": newkey$ = "Procedure"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)
Rem --- Public Sub Statement ---
  Keyword$ = "Public Sub": newkey$ = "Procedure"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem *** SUB Statement ***
  Keyword$ = "SUB": newkey$ = "Procedure" + Chr$(13) + Chr$(10) + "Begin"
  If InStr(l$, Keyword$) > 0 Then
    Call parse(l$, Keyword$, newkey$)
      If InStr(l$, "(") Then
        temp$ = Left$(l$, (InStr(l$, "(")) - 1)
        l$ = RTrim$(temp$) + ";" + Chr$(13) + Chr$(10) + "// " + Right$(l$, (Len(l$) - InStr(l$, "(") + 1))
      Else
        l$ = RTrim$(l$) + ";"
      End If

    If mainflag% = 0 Then
      Print #2, "END; // Procedure"
      Print #2,
      mainflag% = 1
    End If
    Rem Call parse(l$, Keyword$, newkey$)
  End If

Rem *** STR$ Function ***
  Keyword$ = "STR$": newkey$ = "FloatToStr"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem *** SYSTEM Statement ***
  Keyword$ = "SYSTEM": newkey$ = "Application.Terminate"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem *** UCASE$ Function ***
  Keyword$ = "UCASE$": newkey$ = "UpperCase"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem *** VAL Function ***
  Keyword$ = "VAL": newkey$ = "StrToInt"
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)

Rem =======================================
Rem Sub statements
Rem *** CASE Keyword ***
  Keyword$ = "CASE": newkey$ = ""
  If InStr(l$, Keyword$) > 0 Then
    If InStr(l$, "IS") > 0 Then Call parse(l$, "IS", "")
    If InStr(l$, "TO") > 0 Then Call parse(l$, "TO", ".. ")
    Call parse(l$, Keyword$, newkey$)
    l$ = l$ + " :"
  End If
Rem =======================================



Rem /////////////////////////////////////////////////
Rem *** LET Statement ***
  Keyword$ = "LET": newkey$ = ""
  If InStr(l$, Keyword$) > 0 Then Call parse(l$, Keyword$, newkey$)
Rem /////////////////////////////////////////////////

Rem **** Write the translated line ****
writeit:

End Sub

Sub quote(l$, lstr$, lend$, lmid$)
  lpartb = Len(l$)
  quote1 = InStr(l$, Chr$(34))
    lstr$ = Left$(l$, (quote1 - 1)) ' bit before quote
  quote2 = InStr((quote1 + 1), l$, Chr$(34))
    lend$ = Right$(l$, (lpartb - quote2)) ' bit after quote
    lpart = lpartb - (quote2 - quote1)
      lmid$ = Mid$(l$, quote1, (lpartb - lpart + 1)) ' bit in quote
End Sub



