VERSION 5.00
Begin VB.Form commentedForm1 
   BackColor       =   &H00808080&
   Caption         =   "Fast HTML Highlight3"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   ScaleHeight     =   2175
   ScaleWidth      =   4470
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label2 
      Caption         =   "This form just contains Heavily commented Code"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Look at the code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "commentedForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This is just a Heavily commented version of
' the colorhtml function in Form1
' This is meant to help understand regular expresions


' Needs reference to Microsoft VBscript Regular Expressions.
' Get it at http://msdn.microsoft.com/downloads/default.asp?URL=/downloads/sample.asp?url=/msdn-files/027/001/733/msdncompositedoc.xml
' You should download the documentaion as well

Function colorhtml()
'-----------------------------------------------
'Define Regularexpressions for colorize function
'-----------------------------------------------
'Regular expressions allow you to search and replace strings using complex patterns
'This is great for HTML since there are loose standards
'for example <font size="1"> and <font  size =  "1"> and <font size=1> are all valid
'This makes HTML difficult to search properly, but with regular expresion it is easy.

'regx for Tags
    Dim TagregEx, Match, Matches   ' Create variable.
    Set TagregEx = New RegExp      ' Create a regular expression.
    TagregEx.Pattern = "<(.)[^> ]*( ){0,1}[^>]*>"   ' Set pattern.
    ' ok you are probaly wondering what <(.)[^> ]+( )*[^>]+> is.
    ' this is my regular expression to find HTML tags
    ' It basically says match any group of characters begining with
    ' < followed immediatly by any char . which is in parenthesis telling the regx engine
    ' to remember this value for later. - I will use this to see if our tag is a comment or end tag
    ' then we look for [^> ] brackets define a character class, and the ^ means not
    ' this is followed by a * wich means match [^> ] 0 or more times
    ' so [^> ]* means a group of characters that doesn't contain a > or a space
    ' then we have ( ){0,1} match 0 or 1 spaces and store it for later
    ' are you getting the hang of this yet?
    ' last is [^>]*> match any number of characters that isn't a > followed by a >
    ' This seems complex but really it isnt.
    ' A simple tag regx could be written as "<[^>]+>" the + means match one or more times
    ' Mine is more complex, because remembering the first character and
    ' if the tag contained a space after the name
    TagregEx.IgnoreCase = False    ' Set case insensitivity. We aren't looking for any chars so no need.
    TagregEx.Global = True         ' Set global applicability. Global just means match the pattern as many times as possible

'regx for property="value" pairs
    Dim tagPNregEx, Match2, Matches2    ' Create variable.
    Set tagPNregEx = New RegExp         ' Create a regular expression.
    tagPNregEx.Pattern = "(\w+ *=) *(\d+|""[^""]+"")"   ' tag propertyname.
    ' again a funny looking line of jibberish.
    ' all that is new here is \w (any word character) and \d (any digit) and | wich means or
    ' so now that you know that you know this jibberish says
    ' match a group of word characters followed by 0 or more spaces followed by an = sign
    ' followed by 0 or more spaces followed by either a group of digits or
    ' an equal sign followed by anything that isnt an equal sign followed by another equal sign
    ' also we are remembering up to including the equal sign and every thing after the equal sign
    ' what a mouthfull I sure love the fact that I can write all that as "(\w+ *=) *(\d+|""[^""]+"")"
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
            ' now run some short circuit logic so we only run second regx if we have to.
            ' since our tag regx said remember the first char, we will have it in our submatches collection
            If Match.SubMatches(0) = "!" Then ' looks like a comment
               rtf1.SelColor = color(3)
               GoTo nextmatch
            ElseIf Match.SubMatches(1) <> " " Then ' this tag doesn't have properties
                GoTo nextmatch
            End If
            Set Matches2 = tagPNregEx.Execute(Match.Value) ' Execute search.
            ' This would also work as tagPNregEx.Execute (rtf1.SelText)
            ' But since our matches collection has this information it
            ' is faster and cleaner to use the matches collection information
            ' notice, that the only time we need to reference the rtf1.sel attributes
            ' is when we make a color change.
            ' This is what slows everything down, not the regular expressions.
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

