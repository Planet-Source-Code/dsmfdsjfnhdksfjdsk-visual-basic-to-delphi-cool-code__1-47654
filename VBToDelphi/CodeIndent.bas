Attribute VB_Name = "CodeIndent"
Global Thefilename

Sub Indent(n%)
Rem pretty thing up
Rem **** change to LEFT commands - safer ***

DoEvents
  
  add$ = ""
  For j% = 1 To n%           'indent loops
    If Right$(Trim(MyArray$(j%)), 1) = ":" Or InStr(MyArray$(j%), "LOCAL ") > 0 Or InStr(MyArray$(j%), "GLOBAL ") > 0 Or Left$(MyArray$(j%), 4) = "REM " Then
      ' do nothing
    ElseIf InStr(MyArray$(j%), "Procedure ") Or InStr(UCase$(MyArray$(j%)), "END ") Then
      ' do nothing
    ElseIf InStr(MyArray$(j%), "ELSE") And Len(add$) > 1 Then
      temp$ = Left$(add$, Len(add$) - 2) ' deal with ELSE & ELSEIF
      MyArray$(j%) = temp$ + MyArray$(j%)
    Else
      MyArray$(j%) = add$ + MyArray$(j%)
    End If
    
    If InStr(UCase$(MyArray$(j%)), "WHILE ") Then Call Push(add$)
    If InStr(UCase$(MyArray$(j%)), "END") Then Call Pull(add$, j%)
    
    If InStr(UCase$(MyArray$(j%)), "DO ") Then Call Push(add$)
    If InStr(UCase$(MyArray$(j%)), "END") Then Call Pull(add$, j%)
    
    If InStr(UCase$(MyArray$(j%)), "IF ") Then Call Push(add$)
    If InStr(UCase$(MyArray$(j%)), "END") Then Call Pull(add$, j%)
  
  Next j%

End Sub

Sub Pull(add$, j%)
Rem deindent by two spaces
  If Len(add$) > 2 Then
    add$ = Left$(add$, Len(add$) - 2)
    MyArray$(j%) = Right$(MyArray$(j%), Len(MyArray$(j%)) - 2)
  End If
End Sub

Sub Push(add$)
Rem indent by two spaces
  add$ = add$ + "  "
End Sub

