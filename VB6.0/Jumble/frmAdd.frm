VERSION 5.00
Begin VB.Form frmAdd 
   Caption         =   "Add Jumbles"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6075
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtWordsFinal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   16
      Top             =   4200
      Width           =   5940
   End
   Begin VB.TextBox txtFinal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   4200
      TabIndex        =   14
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox txtFinal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   4200
      TabIndex        =   11
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox txtFinal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   4200
      TabIndex        =   8
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox txtFinal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   4200
      TabIndex        =   5
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox txtFinal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   4200
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtAnswer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   3000
      Width           =   5940
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   19
      Top             =   6120
      Width           =   735
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Jumble\jumble.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Word"
      Top             =   6720
      Visible         =   0   'False
      Width           =   4260
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&ADD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   18
      Top             =   6120
      Width           =   735
   End
   Begin VB.TextBox txtCaption 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   5280
      Width           =   5940
   End
   Begin VB.TextBox txtScram 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   2880
      MaxLength       =   7
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtScram 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   2880
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtScram 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   2880
      TabIndex        =   7
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtScram 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   2880
      TabIndex        =   10
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtScram 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   2880
      TabIndex        =   13
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtWord 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   840
      MaxLength       =   7
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtWord 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   840
      MaxLength       =   7
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtWord 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   840
      MaxLength       =   7
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtWord 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   840
      MaxLength       =   7
      TabIndex        =   9
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtWord 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   840
      MaxLength       =   7
      TabIndex        =   12
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label17 
      Caption         =   "Words in final answer:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "In Final Answer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4200
      TabIndex        =   35
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label15 
      Caption         =   "Answer:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   2760
      Width           =   705
   End
   Begin VB.Label Label14 
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   0
      TabIndex        =   33
      Top             =   6600
      Width           =   6135
   End
   Begin VB.Label Label13 
      Caption         =   "Caption:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   120
      TabIndex        =   32
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Scrambled"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   2880
      TabIndex        =   31
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "Regular"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   840
      TabIndex        =   30
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "Word 5:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   29
      Top             =   2280
      Width           =   705
   End
   Begin VB.Label Label9 
      Caption         =   "Word 4:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   28
      Top             =   1800
      Width           =   705
   End
   Begin VB.Label Label8 
      Caption         =   "Word 3:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   27
      Top             =   1320
      Width           =   705
   End
   Begin VB.Label Label7 
      Caption         =   "Word 2:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   26
      Top             =   840
      Width           =   705
   End
   Begin VB.Label Label6 
      Caption         =   "Word 1:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   25
      Top             =   360
      Width           =   705
   End
   Begin VB.Label Label5 
      Caption         =   "Word 5:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   2280
      Width           =   705
   End
   Begin VB.Label Label4 
      Caption         =   "Word 4:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   1800
      Width           =   705
   End
   Begin VB.Label Label3 
      Caption         =   "Word 3:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   1320
      Width           =   705
   End
   Begin VB.Label Label2 
      Caption         =   "Word 2:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   840
      Width           =   705
   End
   Begin VB.Label Label1 
      Caption         =   "Word 1:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   360
      Width           =   700
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim vbCaption As String
    Dim vbAnswer As String
    Dim vbWords As String
    Dim vbFinal As String
    Dim vbScram As String
    Dim Numbers2Puzzle As String
    Dim I As Integer
    Dim Lengths As String
    Dim Temp As String
    Dim Index As Integer
    
    vbCaption = UCase(txtCaption.Text)
    vbAnswer = UCase(txtAnswer.Text)
    vbFinal = UCase(txtWordsFinal.Text)
    If vbCaption = "" Then
        MsgBox "The caption must be entered.", vbOKOnly + vbExclamation, "Enter a caption for this jumble."
        Exit Sub
    End If
    If vbAnswer = "" Then
        MsgBox "The answer must be entered.", vbOKOnly + vbExclamation, "Enter the answer for this jumble."
        Exit Sub
    End If
    If vbFinal = "" Then
        MsgBox "The words in final answer must be entered.", vbOKOnly + vbExclamation, "Enter the answer for this jumble."
        Exit Sub
    End If
        
    For I = 1 To 4
        If txtWord(I).Text = "" Then
            MsgBox "The words must be entered.", vbOKOnly + vbExclamation, "Enter the answer for this jumble."
            Exit Sub
        End If
        If txtScram(I).Text = "" Then
            MsgBox "The scrambled words must be entered.", vbOKOnly + vbExclamation, "Enter the answer for this jumble."
            Exit Sub
        End If
        vbWords = vbWords + txtWord(I) + ","
        vbScram = vbScram + txtScram(I) + ","
        If txtFinal(I).Text = "" Then
            MsgBox "The in final answer must be entered.", vbOKOnly + vbExclamation, "Enter the answer for this jumble."
            Exit Sub
        End If
        Numbers2Puzzle = Numbers2Puzzle + txtFinal(I) + "-"
    Next I
    If txtWord(5).Text <> "" Then
        If txtScram(5).Text = "" Then
            MsgBox "The scrambled words must be entered.", vbOKOnly + vbExclamation, "Enter the answer for this jumble."
            Exit Sub
        End If
        If txtFinal(5).Text = "" Then
            MsgBox "The in final answer must be entered.", vbOKOnly + vbExclamation, "Enter the answer for this jumble."
            Exit Sub
        End If
        vbWords = vbWords + txtWord(5)
        vbScram = vbScram + txtScram(5)
        Numbers2Puzzle = Numbers2Puzzle + txtFinal(5)
    End If
    If Right$(vbWords, 1) = "," Then
        vbWords = Left$(vbWords, Len(vbWords) - 1)
    End If
    If Right$(vbScram, 1) = "," Then
        vbScram = Left$(vbScram, Len(vbScram) - 1)
    End If
    If Right$(Numbers2Puzzle, 1) = "-" Then
        Numbers2Puzzle = Left$(Numbers2Puzzle, Len(Numbers2Puzzle) - 1)
    End If
    vbWords = LTrim(Str(I - 1)) + "," + vbWords
    vbScram = LTrim(Str(I - 1)) + "," + vbScram
    Data1.Recordset.Edit
    Data1.Recordset.AddNew
    Data1.Recordset.Fields("Words") = vbWords
    Data1.Recordset.Fields("Scrambled") = vbScram
    Data1.Recordset.Fields("NumbersToPuzzle") = Numbers2Puzzle
    Temp = vbAnswer
    Do While I > 0
        I = InStr(Temp, " ")
        I = I - 1
        If I = -1 Then
            Lengths = Lengths + LTrim(Str(Len(Temp)))
        Else
            Lengths = Lengths + LTrim(Str(I)) + ","
            Temp = Mid$(Temp, I + 2, Len(vbAnswer))
        End If
        'I = InStr(Temp, " ")
    Loop
    Data1.Recordset.Fields("Lengths") = Lengths
    Data1.Recordset.Fields("Answer") = vbAnswer
    Data1.Recordset.Fields("NumberOfWords") = vbFinal
    Data1.Recordset.Fields("Caption") = vbCaption
    Data1.Recordset.Update
    Data1.Refresh
    For I = 1 To 5
        txtWord(I).Text = ""
        txtScram(I).Text = ""
        txtFinal(I).Text = ""
    Next I
    txtCaption.Text = ""
    txtAnswer.Text = ""
    txtWordsFinal.Text = ""
    
End Sub

Private Sub Command2_Click()
    Unload Me
    frmJumble.Show
    frmJumble.Data1.Refresh
End Sub

Private Sub Command3_Click()
Form1.Show

End Sub

Private Sub Form_Load()
    Data1.Refresh
End Sub


Private Sub txtScram_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim I As Integer
    Dim A As String
    Dim Temp As String
    
    If KeyAscii = 8 Then
        Exit Sub
    End If
    A = UCase(Chr(KeyAscii))
    Temp = txtWord(Index).Text
    For I = 1 To Len(txtWord(Index))
        If Mid$(Temp, I, 1) = A Then
            KeyAscii = Asc(A)
            Exit Sub
        End If
    Next I
    KeyAscii = 0
    
End Sub

Private Sub txtWord_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim A As String
    
    If KeyAscii = 8 Then
        Exit Sub
    End If
    A = UCase(Chr(KeyAscii))
    KeyAscii = Asc(A)
    
End Sub
