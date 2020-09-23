VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visual Basic 6 to Delphi Source Converter"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   2475
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "Idle"
            TextSave        =   "Idle"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   1200
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CD2 
      Left            =   1920
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse1 
      Caption         =   "Browse"
      Height          =   285
      Left            =   6000
      TabIndex        =   7
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   285
      Left            =   6000
      TabIndex        =   4
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Convert"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Frame fmeMain 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.TextBox txtOutput 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   600
         Width           =   4575
      End
      Begin VB.TextBox txtInput 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label lblNote 
         Caption         =   "Sorry, I don't have enough time to put Delphi into VB code, maybe tomarrow I will before my school starts."
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   6495
      End
      Begin VB.Label lblOutput 
         Caption         =   "Output:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblInput 
         Caption         =   "Input:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()
    CD1.ShowOpen
    If Len(CD1.filename) = 0 Then
        If Not CD1.CancelError = True Then
            Exit Sub
        Else
            MsgBox "Please select an input file.", vbInformation, "Input Error"
        End If
        Else
            txtInput.Text = CD1.filename
    End If
    If Trim(txtOutput) = "" Then
        cmdOk.Enabled = False
    Else
        cmdOk.Enabled = True
    End If
End Sub

Private Sub cmdBrowse1_Click()
    CD2.ShowSave
    If Trim(CD2.filename) = "" Then
        If CD2.CancelError = True Then
            Exit Sub
        Else
            MsgBox "Please select an output file.", vbInformation, "Input Error"
        End If
        Else
            txtOutput.Text = CD2.filename
    End If
    If Trim(txtInput.Text) = "" Then
        cmdOk.Enabled = False
    Else
        cmdOk.Enabled = True
    End If
End Sub

Private Sub cmdOk_Click()
Dim File As String
Dim NewFile As String
Dim bl
Dim num

    StatusBar1.SimpleText = "Now Converting..."

    File = txtInput.Text
    NewFile = txtOutput.Text

    Open File$ For Input As 1
    Open NewFile$ For Output As 2
    Print #2, "Procedure main"
 
    While Not EOF(1)
        Line Input #1, bl
        num = num + 1
        StatusBar1.SimpleText = CStr(num) & " Lines Read"
        looksee (bl)
        DoEvents
    Wend

    If mainflag% = 0 Then
        Print #2, "END; // Main Procedure"
        mainflag = 1
    End If

    Close 1, 2
  
    Call FindVars
    
    StatusBar1.SimpleText = "Added Indenting, Idle"
End Sub

