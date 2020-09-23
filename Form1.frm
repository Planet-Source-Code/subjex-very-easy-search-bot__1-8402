VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form1 
   Caption         =   "Search Sample"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox search1 
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Text            =   "Astalavista"
      Top             =   0
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   3360
      TabIndex        =   2
      Text            =   "Keywords to search"
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   315
      Left            =   5040
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin SHDocVwCtl.WebBrowser Web1 
      Height          =   3615
      Left            =   840
      TabIndex        =   0
      Top             =   1800
      Width           =   5895
      ExtentX         =   10398
      ExtentY         =   6376
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label Label1 
      Caption         =   "ENGINE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' hi this is just a sample of how to search the
' search engines off your app.
' I hope the code is easy enough to figure out on your own
'                                   -pugz- voiceone@hotmail.com
Private Sub Command1_Click()
Dim searchme As String
searchme = search1
Select Case searchme$
    Case "Yahoo"
        Web1.Navigate "http://search.yahoo.com/bin/search?p=" + ReplaceString(Text1, " ", "+")
    Case "Astalavista"
        Web1.Navigate "http://astalavista7.box.sk/cgi-bin/astalavista/robot?srch=" + ReplaceString(Text1, " ", "+") + "&submit=+search+"
    Case "Goto"
        Web1.Navigate "http://www.goto.com/d/search/;$sessionid$0IGHHSYACKNGRQFIEE5APUQ?type=home&Keywords=" + ReplaceString(Text1, " ", "+")
    Case Else
        MsgBox "Invalid Search Engine, Please enter a valid search engine", "Opps"
End Select
End Sub

Private Sub Form_Load()
search1.AddItem "Yahoo"
search1.AddItem "Astalavista"
search1.AddItem "Goto"
End Sub
Public Function ReplaceString(MyString As String, ToFind As String, ReplaceWith As String) As String
    ' this function is brought to you
    ' by the great DoS
    ' geez, I am always hearing good stuff about him
    ' I hope your not getting all big headed DoS - (Chad)
    Dim Spot As Long, NewSpot As Long, LeftString As String
    Dim RightString As String, newstring As String
    Spot& = InStr(LCase(MyString$), LCase(ToFind))
    NewSpot& = Spot&
    Do
        If NewSpot& > 0& Then
            LeftString$ = Left(MyString$, NewSpot& - 1)
            If Spot& + Len(ToFind$) <= Len(MyString$) Then
                RightString$ = Right(MyString$, Len(MyString$) - NewSpot& - Len(ToFind$) + 1)
            Else
                RightString = ""
            End If
            newstring$ = LeftString$ & ReplaceWith$ & RightString$
            MyString$ = newstring$
        Else
            newstring$ = MyString$
        End If
        Spot& = NewSpot& + Len(ReplaceWith$)
        If Spot& > 0 Then
            NewSpot& = InStr(Spot&, LCase(MyString$), LCase(ToFind$))
        End If
    Loop Until NewSpot& < 1
    ReplaceString$ = newstring$
End Function

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then ' if you hit enter after typing in your keywords
Dim searchme As String
searchme = search1
Select Case searchme$
    Case "Yahoo"
        Web1.Navigate "http://search.yahoo.com/bin/search?p=" + ReplaceString(Text1, " ", "+")
    Case "Astalavista"
        Web1.Navigate "http://astalavista7.box.sk/cgi-bin/astalavista/robot?srch=" + ReplaceString(Text1, " ", "+") + "&submit=+search+"
    Case "Goto"
        Web1.Navigate "http://www.goto.com/d/search/;$sessionid$0IGHHSYACKNGRQFIEE5APUQ?type=home&Keywords=" + ReplaceString(Text1, " ", "+")
    Case Else
        MsgBox "Invalid Search Engine, Please enter a valid search engine", "Opps"
End Select
search1 = "" 'clears search1
End If
End Sub
