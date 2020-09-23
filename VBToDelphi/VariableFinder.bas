Attribute VB_Name = "VariableFinder"
Global Resword$(236)
Global Nw%
Global MyArray$()
Global Qsorttemp$
Global MaxRow
Global VarsArray$()

Static Sub BubbleSort()

   Limit = MaxRow
   Do
      theSwitch = False
      For Row = 1 To (Limit - 1)

         If MyArray$(Row) > MyArray$(Row + 1) Then
            temp$ = MyArray$(Row + 1)
            MyArray$(Row + 1) = MyArray$(Row)
            MyArray$(Row) = temp$
            theSwitch = Row
         End If
      Next Row

      Limit = theSwitch
   Loop While theSwitch

End Sub

Sub FindVars()



Call Reserved
subname$ = "main"
stream = FreeFile
Open frmMain.txtInput.Text For Input As stream


DoEvents

While Not EOF(stream)
nxtline:
  Line Input #stream, bl$
  bl$ = bl$ + " "
  place% = 0
    DoEvents


  If Right$(Trim(bl$), 1) = ":" Then
    word$ = ""
    GoTo nxtline
  End If
  
  For index% = 1 To Len(bl$)
  Call GetC(bl$, a$, ungetdat$, place%)
' MsgBox (word$ + "*" + a$ + "*")
  If IsAlpha%(a$) Then
     word$ = word$ + a$
  End If

  If Not IsAlpha%(a$) Then
    If word$ <> "" Then
     temp$ = Left$(word$, 1)
     c% = Asc(temp$)
     If (c% >= Asc("0")) And (c% <= Asc("9")) Then
       word$ = ""
     End If
    End If

    If Not IsReserved%(word$) Then
      If Trim(word$) <> "" Then
        n% = n% + 1
        ReDim Preserve MyArray$(n%)
        MyArray$(n%) = Trim(word$) + "," + Trim(subname$)
      End If
    End If
    
    If word$ = "REM" Or word$ = "'" Or word$ = "GOTO" Or word$ = "CALL" Or word$ = "GOSUB" Or word$ = "DECLARE" Or word$ = "EXIT" Or word$ = "DATA" Then ' ignor comments
      word$ = ""
      GoTo nxtline
    End If
    
    If a$ = Chr$(34) Then
      Do
        Call GetC(bl$, a$, ungetdat$, place%)
      Loop Until a$ = Chr$(34)
    End If
    
    If word$ = "SUB" And InStr(bl$, "END") = 0 Then
      subname$ = ""
      Do
        Call GetC(bl$, a$, ungetdat$, place%)
        subname$ = subname$ + a$
      Loop Until a$ = Chr$(13) Or a$ = "(" Or a$ = " "
    End If
    
    word$ = ""
    Nw% = Nw% + 1
    DoEvents
  End If

  Next index%
Wend
Close 1, 2

 DoEvents
 MaxRow = n%

Call BubbleSort
    For q1% = 1 To n%
      If q1% > 1 Then
        If MyArray(q1% - 1) <> "" Then Last$ = MyArray(q1% - 1)
        If MyArray(q1%) = Last$ Then
         MyArray$(q1%) = ""
        End If
      End If
    Next q1%
 
 Call BubbleSort
 
    For q1% = 1 To n%
      If q1% > 1 And InStr(MyArray$(q1% - 1), ",") Then
        cp% = InStr(MyArray$(q1% - 1), ",")
          Last$ = Left$(MyArray$(q1% - 1), cp% - 1)
        
        If Left$(MyArray(q1%), (InStr(MyArray$(q1%), ",")) - 1) = Last$ Then
         MyArray$(q1%) = Last$ + ",GLOBAL"
           If InStr(MyArray(q1% - 1), "GLOBAL") = 0 Then
             MyArray$(q1% - 1) = Last$ + ",GLOBAL"
           End If
        End If
      End If
    Next q1%

 For q1% = 1 To n%
  If InStr(MyArray$(q1%), ",") Then
    cp% = InStr(MyArray$(q1%), ",")
    temp$ = Left$(MyArray$(q1%), cp% - 1) ' var
    temp2$ = Right$(MyArray$(q1%), Len(MyArray$(q1%)) - cp%)
    MyArray$(q1%) = temp2$ + "," + temp$
  End If
 Next q1%

 Call BubbleSort
    For q1% = 1 To n%
      If q1% > 1 Then
        If MyArray(q1% - 1) <> "" Then Last$ = MyArray(q1% - 1)
        If MyArray(q1%) = Last$ Then
         MyArray$(q1%) = ""
        End If
      End If
    Next q1%

 Open "c:\vars.txt" For Output As 1
  Print #1, "****** Variables *******"
    For q2% = 1 To n%
      If MyArray$(q2%) <> "" Then
        Print #1, MyArray$(q2%)
        vn% = vn% + 1
        ReDim Preserve VarsArray$(vn%)
        VarsArray$(vn%) = MyArray$(q2%)
      End If
    Next q2%
 Close 1

 Call InsertVars(vn%)

thefin:

 DoEvents

MsgBox "Conversion Successful!", vbInformation, "Yay!"

End Sub

Sub GetC(bl$, datavar$, ungetdat$, place%)

  place% = place + 1
  datavar$ = Mid$(bl$, place%, 1) ' get a single character

End Sub

Sub InsertVars(vn%)
Rem read file back into array & add Variable declares

stream = FreeFile
Open frmMain.txtOutput.Text For Input As stream
ReDim MyArray$(2)
         
  While Not EOF(stream)
    n% = n% + 1
    Line Input #stream, theline$
    ReDim Preserve MyArray$(n%)
    MyArray$(n%) = Trim(theline$)
      If InStr(theline$, "PROC") Then
        pcnt% = pcnt% + 1
        For idx% = 1 To vn%
          cp% = InStr(VarsArray$(idx%), ",")
          temp$ = Left$(VarsArray$(idx%), cp% - 1)
          temp2$ = Right$(VarsArray$(idx%), Len(VarsArray$(idx%)) - cp%)
          
          If Len(temp2$) > 8 Then
            temp2$ = temp2$ + " :REM too long"
          End If
          
          If Right(temp2$, 1) = "$" Then
            temp2$ = temp2$ + "(1)"
          End If
          
          If pcnt% = 1 Then
            If temp$ = "GLOBAL" Then
              n% = n% + 1
              ReDim Preserve MyArray$(n%)
              MyArray$(n%) = "GLOBAL " + temp2$
            ElseIf temp$ = "main" Then
              n% = n% + 1
              ReDim Preserve MyArray$(n%)
              MyArray$(n%) = "LOCAL " + temp2$
            End If
          End If
          
          If InStr(theline$, temp$) And temp$ <> "main" Then
            n% = n% + 1
            ReDim Preserve MyArray$(n%)
            MyArray$(n%) = "LOCAL " + temp2$
          End If

        Next idx%
      End If
  Wend
  Close stream

  Call Indent(n%)

  stream = FreeFile
  Open frmMain.txtOutput.Text For Output As stream
  Print #stream, "// Converted by Visual Basic to Delphi Converter made by Danny J on " + Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 2) + vbCrLf + "// This code belongs to www.Planet-Source-Code.com"
  For idx% = 1 To n%
    Print #stream, MyArray$(idx%)
  Next idx%
  Close stream
End Sub

Function IsAlpha%(datavar$)

   If datavar$ = "" Then
      IsAlpha% = 0
   ElseIf Asc(datavar$) > 300 Then
      IsAlpha% = 0
   Else
      CharAsc = Asc(datavar$)
      IsAlpha% = (CharAsc >= Asc("A") And (CharAsc <= Asc("Z")) Or (CharAsc >= Asc("0")) And (CharAsc <= Asc("9")) Or (CharAsc = Asc("_")) Or (CharAsc = Asc("$")) Or (CharAsc = Asc("%")) Or (CharAsc >= Asc("a")) And (CharAsc <= Asc("z")))
   End If
End Function

Function IsReserved%(word$)

  IsReserved% = 0
  For cindex = 1 To 236
    If word$ = Resword$(cindex) Then
      IsReserved% = -1
      Exit Function
    End If
  Next cindex
End Function

Sub Reserved()
Resword$(1) = "ABS"
Resword$(2) = "ABSOLUTE"
Resword$(3) = "ACCESS"
Resword$(4) = "AND"
Resword$(5) = "ANY"
Resword$(6) = "APPEND"
Resword$(7) = "AS"
Resword$(8) = "ASC"
Resword$(9) = "ATN"
Resword$(10) = "BASE"
Resword$(11) = "BEEP"
Resword$(12) = "BINARY"
Resword$(13) = "BLOAD"
Resword$(14) = "BSAVE"
Resword$(15) = "CALL"
Resword$(16) = "CASE"
Resword$(17) = "CDBL"
Resword$(18) = "CHAIN"
Resword$(19) = "CHDIR"
Resword$(20) = "CHR$"
Resword$(21) = "CINT"
Resword$(22) = "CIRCLE"
Resword$(23) = "CLEAR"
Resword$(24) = "CLNG"
Resword$(25) = "CLOSE"
Resword$(26) = "CLS"
Resword$(27) = "COLOR"
Resword$(28) = "COM"
Resword$(29) = "COMMON"
Resword$(30) = "CONST"
Resword$(31) = "COS"
Resword$(32) = "CSNG"
Resword$(33) = "CSRLIN"
Resword$(34) = "CVD"
Resword$(35) = "CVDMBF"
Resword$(36) = "CVI"
Resword$(37) = "CVL"
Resword$(38) = "CVS"
Resword$(39) = "CVSMBF"
Resword$(40) = "DATA"
Resword$(41) = "DATE$"
Resword$(42) = "DATE$"
Resword$(43) = "DECLARE"
Resword$(44) = "DEF"
Resword$(45) = "FN"
Resword$(46) = "SEG"
Resword$(47) = "DEFDBL"
Resword$(48) = "DEFINT"
Resword$(49) = "DEFLNG"
Resword$(50) = "DEFSNG"
Resword$(51) = "DEFSTR"
Resword$(52) = "DIM"
Resword$(53) = "DO"
Resword$(54) = "LOOP"
Resword$(55) = "DOUBLE"
Resword$(56) = "DRAW"
Resword$(57) = "$DYNAMIC"
Resword$(58) = "ELSE"
Resword$(59) = "ELSEIF"
Resword$(60) = "END"
Resword$(61) = "ENVIRON"
Resword$(62) = "ENVIRON$"
Resword$(63) = "EOF"
Resword$(64) = "EQV"
Resword$(65) = "ERASE"
Resword$(66) = "ERDEV"
Resword$(67) = "ERDEV$"
Resword$(68) = "ERL"
Resword$(69) = "ERR"
Resword$(70) = "ERROR"
Resword$(71) = "EXIT"
Resword$(72) = "EXP"
Resword$(73) = "FIELD"
Resword$(74) = "FILEATTR"
Resword$(75) = "FILES"
Resword$(76) = "FIX"
Resword$(77) = "FOR"
Resword$(78) = "NEXT"
Resword$(79) = "FRE"
Resword$(80) = "FREEFILE"
Resword$(81) = "GET"
Resword$(82) = "GOSUB"
Resword$(83) = "GOTO"
Resword$(84) = "HEX$"
Resword$(85) = "IF"
Resword$(86) = "THEN"
Resword$(87) = "ELSE"
Resword$(88) = "IMP"
Resword$(89) = "INKEY$"
Resword$(90) = "INP"
Resword$(91) = "INPUT"
Resword$(92) = "INPUT$"
Resword$(93) = "INSTR"
Resword$(94) = "INT"
Resword$(95) = "INTEGER"
Resword$(96) = "IOCTL"
Resword$(97) = "IOCTL$"
Resword$(98) = "IS"
Resword$(99) = "KEY"
Resword$(100) = "KEY"
Resword$(101) = "KILL"
Resword$(102) = "LBOUND"
Resword$(103) = "LCASE$"
Resword$(104) = "LEFT$"
Resword$(105) = "LEN"
Resword$(106) = "LET"
Resword$(107) = "LINE"
Resword$(108) = "LIST"
Resword$(109) = "LOC"
Resword$(110) = "LOCATE"
Resword$(111) = "LOCK"
Resword$(112) = "UNLOCK"
Resword$(113) = "LOF"
Resword$(114) = "LOG"
Resword$(115) = "LONG"
Resword$(116) = "LOOP"
Resword$(117) = "LPOS"
Resword$(118) = "LPRINT"
Resword$(119) = "USING"
Resword$(120) = "LSET"
Resword$(121) = "LTRIM$"
Resword$(122) = "MID$"
Resword$(123) = "MID$"
Resword$(124) = "MKD$"
Resword$(125) = "MKDIR"
Resword$(126) = "MKDMBF$"
Resword$(127) = "MKI$"
Resword$(128) = "MKL$"
Resword$(129) = "MKS$"
Resword$(130) = "MKSMBF$"
Resword$(131) = "MOD"
Resword$(132) = "NAME"
Resword$(133) = "NEXT"
Resword$(134) = "NOT"
Resword$(135) = "OCT$"
Resword$(136) = "OFF"
Resword$(137) = "ON"
Resword$(138) = "COM"
Resword$(139) = "ERROR"
Resword$(140) = "KEY"
Resword$(141) = "PEN"
Resword$(142) = "PLAY"
Resword$(143) = "STRIG"
Resword$(144) = "TIMER"
Resword$(145) = "GOSUB"
Resword$(146) = "GOTO"
Resword$(147) = "OPEN"
Resword$(148) = "OPTION"
Resword$(149) = "BASE"
Resword$(150) = "OR"
Resword$(151) = "OUT"
Resword$(152) = "OUTPUT"
Resword$(153) = "PAINT"
Resword$(154) = "PALETTE"
Resword$(155) = "PCOPY"
Resword$(156) = "PEEK"
Resword$(157) = "PEN"
Resword$(158) = "PEN"
Resword$(159) = "PLAY"
Resword$(160) = "PMAP"
Resword$(161) = "POINT"
Resword$(162) = "POKE"
Resword$(163) = "POS"
Resword$(164) = "PRESET"
Resword$(165) = "PRINT"
Resword$(166) = "USING"
Resword$(167) = "PSET"
Resword$(168) = "PUT"
Resword$(169) = "RANDOM"
Resword$(170) = "RANDOMIZE"
Resword$(171) = "READ"
Resword$(172) = "REDIM"
Resword$(173) = "REM"
Resword$(174) = "RESET"
Resword$(175) = "RESTORE"
Resword$(176) = "RESUME"
Resword$(177) = "RETURN"
Resword$(178) = "RIGHT$"
Resword$(179) = "RMDIR"
Resword$(180) = "RND"
Resword$(181) = "RSET"
Resword$(182) = "RTRIM$"
Resword$(183) = "RUN"
Resword$(184) = "SCREEN"
Resword$(185) = "SEEK"
Resword$(186) = "SELECT"
Resword$(187) = "CASE"
Resword$(188) = "SGN"
Resword$(189) = "SHARED"
Resword$(190) = "SHELL"
Resword$(191) = "SIN"
Resword$(192) = "SINGLE"
Resword$(193) = "SLEEP"
Resword$(194) = "SOUND"
Resword$(195) = "SPACE$"
Resword$(196) = "SPC"
Resword$(197) = "SQR"
Resword$(198) = "STATIC"
Resword$(199) = "$STATIC"
Resword$(200) = "STEP"
Resword$(201) = "STICK"
Resword$(202) = "STOP"
Resword$(203) = "STR$"
Resword$(204) = "STRIG"
Resword$(205) = "STRING"
Resword$(206) = "STRING$"
Resword$(207) = "SUB"
Resword$(208) = "SWAP"
Resword$(209) = "SYSTEM"
Resword$(210) = "TAB"
Resword$(211) = "TAN"
Resword$(212) = "THEN"
Resword$(213) = "TIME$"
Resword$(214) = "TIMER"
Resword$(215) = "TO"
Resword$(216) = "TROFF"
Resword$(217) = "TRON"
Resword$(218) = "TYPE"
Resword$(219) = "UBOUND"
Resword$(220) = "UCASE$"
Resword$(221) = "UNLOCK"
Resword$(222) = "UNTIL"
Resword$(223) = "USING"
Resword$(224) = "VAL"
Resword$(225) = "VARPTR"
Resword$(226) = "VARPTR$"
Resword$(227) = "VARSEG"
Resword$(228) = "VIEW"
Resword$(229) = "WAIT"
Resword$(230) = "WEND"
Resword$(231) = "WHILE"
Resword$(232) = "WEND"
Resword$(233) = "WIDTH"
Resword$(234) = "WINDOW"
Resword$(235) = "WRITE"
Resword$(236) = "XOR"
End Sub

