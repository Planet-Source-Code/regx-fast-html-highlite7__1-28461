VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   Caption         =   "Fast HTML Highlight7"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10050
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   10050
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdregx 
      Caption         =   "Run Regx Only"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Un-Highlight"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Highlight All"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Highlight Selected"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   1575
   End
   Begin RichTextLib.RichTextBox rtf1 
      CausesValidation=   0   'False
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   6588
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.Label lbltagcount 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Top             =   3840
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Fast HTML Highlight
'--------------------------------------------------
'Copyright 2001 DGS http://www.2dgs.com
'Written by Gary Varnell
'You may use this code freely as long as the above
'copyright info remains intact
'==================================================
' Needs reference to Microsoft VBscript Regular Expressions.
' Get it at http://msdn.microsoft.com/downloads/default.asp?URL=/downloads/sample.asp?url=/msdn-files/027/001/733/msdncompositedoc.xml
Option Explicit
Dim apppath As String
Dim starttime As Date
Dim tmpchr As String * 1
Dim tmpint As Long
Dim color(3) As Variant

Function colorhtml()
'-----------------------------------------------
'Define Regularexpressions for colorize function
'-----------------------------------------------
'regx for Tags
    Dim TagregEx, Match, Matches   ' Create variable.
    Set TagregEx = New RegExp      ' Create a regular expression.
    TagregEx.Pattern = "<(.)[^> ]*( ){0,1}[^>]*>"   ' Set pattern.
    TagregEx.IgnoreCase = False    ' Set case insensitivity.
    TagregEx.Global = True         ' Set global applicability.

'regx for property="value" pairs
    Dim tagPNregEx, Match2, Matches2    ' Create variable.
    Set tagPNregEx = New RegExp         ' Create a regular expression.
    tagPNregEx.Pattern = "(\w+ *=) *(\d+|""[^""]+"")"   ' tag propertyname.

    tagPNregEx.IgnoreCase = False       ' Set case insensitivity.
    tagPNregEx.Global = True            ' Set global applicability.
'---------------------------------------------
Dim rtfstart As Long
rtfstart = rtf1.SelStart ' Remember startpos since user might have selected text
If rtf1.SelLength < 1 Then
    MsgBox "No text selected"
Exit Function
End If
'----------------------------------------------
    Set Matches = TagregEx.Execute(rtf1.SelText)    ' Execute search.
    For Each Match In Matches     ' Iterate Matches collection.
        If Match.Value <> "" Then 'used to stop empty string match return
            rtf1.SelStart = rtfstart + Match.FirstIndex
            rtf1.SelLength = Match.Length
            rtf1.SelColor = color(0)
            ' now run some short circuit logic
            If Match.SubMatches(0) = "!" Then ' looks like a comment
               rtf1.SelColor = color(3)
               GoTo nextmatch
            ElseIf Match.SubMatches(1) <> " " Then ' this tag doesn't have properties
                GoTo nextmatch
            End If
            Set Matches2 = tagPNregEx.Execute(Match.Value) ' Execute search.
            For Each Match2 In Matches2
                If Match2.Value <> "" Then 'used to stop empty string match return
                    'Debug.Print Match2.Value & Match2.Length
                    rtf1.SelStart = Match.FirstIndex + rtfstart + Match2.FirstIndex
                    rtf1.SelLength = Match2.Length
                    rtf1.SelColor = color(2)
                    rtf1.SelLength = Len(Match2.SubMatches(0))
                    rtf1.SelColor = color(1)
                End If
            Next
        End If
nextmatch:
    Next
    lbltagcount.Caption = Matches.Count & " Tags"
End Function
Function regxonly()
'-----------------------------------------------
'Define Regularexpressions for colorize function
'-----------------------------------------------
'regx for Tags
    Dim TagregEx, Match, Matches   ' Create variable.
    Set TagregEx = New RegExp      ' Create a regular expression.
    TagregEx.Pattern = "<(.)[^> ]*( ){0,1}[^>]*>"   ' Set pattern.
    TagregEx.IgnoreCase = False    ' Set case insensitivity.
    TagregEx.Global = True         ' Set global applicability.

'regx for property="value" pairs
    Dim tagPNregEx, Match2, Matches2    ' Create variable.
    Set tagPNregEx = New RegExp         ' Create a regular expression.
    tagPNregEx.Pattern = "(\w+ *=) *(\d+|""[^""]+"")"   ' tag propertyname.

    tagPNregEx.IgnoreCase = False       ' Set case insensitivity.
    tagPNregEx.Global = True            ' Set global applicability.
'---------------------------------------------
Dim rtfstart As Long
rtfstart = rtf1.SelStart ' Remember startpos since user might have selected text
If rtf1.SelLength < 1 Then
    MsgBox "No text selected"
Exit Function
End If
'----------------------------------------------
    Set Matches = TagregEx.Execute(rtf1.SelText)    ' Execute search.
    For Each Match In Matches     ' Iterate Matches collection.
        If Match.Value <> "" Then 'used to stop empty string match return
            ' now run some short circuit logic
            If Match.SubMatches(0) = "!" Then ' looks like a comment
               GoTo nextmatch
            ElseIf Match.SubMatches(1) <> " " Then ' this tag doesn't have properties
                GoTo nextmatch
            End If
            Set Matches2 = tagPNregEx.Execute(Match.Value) ' Execute search.
            For Each Match2 In Matches2
                If Match2.Value <> "" Then 'used to stop empty string match return
                End If
            Next
        End If
nextmatch:
    Next
    lbltagcount.Caption = Matches.Count & " Tags"
End Function

Private Sub cmdregx_Click()
starttime = Time
Me.MousePointer = vbHourglass
rtf1.Visible = False
rtf1.SelStart = 0
rtf1.SelLength = Len(rtf1.Text)
regxonly
rtf1.SelStart = 1
rtf1.Visible = True
Me.MousePointer = vbNormal
MsgBox Time - starttime
End Sub

Private Sub Form_Load()
apppath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\")
rtf1.LoadFile apppath & "test.html"
' Define highlight colors

color(0) = vbBlue  'Tag color
color(1) = vbRed    'Tag property color
color(2) = &H8000&  'Tag property value color
color(3) = vbMagenta   'comment Tag color
End Sub
Private Sub Command1_Click()
colorhtml
End Sub

Private Sub Command2_Click()
starttime = Time
Me.MousePointer = vbHourglass
rtf1.Visible = False ' comment out to watch colorize function in action
rtf1.SelStart = 0
rtf1.SelLength = Len(rtf1.Text)
colorhtml
rtf1.SelStart = 1
rtf1.Visible = True ' comment out to watch colorize function in action
Me.MousePointer = vbNormal
MsgBox Time - starttime
End Sub

Private Sub Command3_Click()
rtf1.TextRTF = rtf1.Text
Me.lbltagcount.Caption = ""
End Sub
Private Sub Form_Resize()
rtf1.Left = 0
rtf1.Width = Form1.ScaleWidth
rtf1.Height = Form1.ScaleHeight - 500
Command1.Top = Form1.ScaleHeight - 400
Command2.Top = Command1.Top
Command3.Top = Command1.Top
cmdregx.Top = Command1.Top
lbltagcount.Top = Command1.Top

End Sub


Private Sub rtf1_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = "<" Then
    rtf1.SelColor = color(0)
End If
If INtag = True Then
    If Chr(KeyAscii) = " " Then
        If INpropval Then
            rtf1.SelColor = color(2)
        Else
            rtf1.SelColor = color(1)
        End If
    ElseIf Chr(KeyAscii) = """" Then
            rtf1.SelColor = color(2)
    ElseIf Chr(KeyAscii) = ">" Then
            rtf1.SelColor = color(0)
    ElseIf Chr(KeyAscii) = "!" Then
            rtf1.SelColor = color(3)
    End If
End If
End Sub

Private Sub rtf1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode & Shift = "1901" Then ' user pressed >
    rtf1.SelColor = vbBlack
End If
End Sub

Private Function INtag() As Boolean
If InStrRev(rtf1.Text, "<", rtf1.SelStart, vbTextCompare) > InStrRev(rtf1.Text, ">", rtf1.SelStart, vbTextCompare) Then INtag = True
End Function
Private Function INpropval() As Boolean
Dim x, y As Long
x = InStrRev(rtf1.Text, """", rtf1.SelStart, vbTextCompare)
y = InStrRev(rtf1.Text, "=", rtf1.SelStart, vbTextCompare)
If x > y Then
If InStrRev(rtf1.Text, """", x - 1, vbTextCompare) < InStrRev(rtf1.Text, "=", x - 1, vbTextCompare) Then INpropval = True
End If
End Function

