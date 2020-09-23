VERSION 5.00
Begin VB.Form frmJumble 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Jumble"
   ClientHeight    =   6105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8490
   FillColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Mask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   495
      TabIndex        =   88
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Jumble\jumble.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Word"
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox txt1 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txt1 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txt1 
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
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txt1 
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
      Left            =   2160
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txt1 
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
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txt1 
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
      Index           =   6
      Left            =   2880
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1800
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   35
      Top             =   4680
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txt1 
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
      Index           =   7
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtWord 
      Height          =   285
      Index           =   1
      Left            =   2520
      TabIndex        =   79
      Top             =   1560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtWord 
      Height          =   285
      Index           =   2
      Left            =   2520
      TabIndex        =   78
      Top             =   2160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtWord 
      Height          =   285
      Index           =   3
      Left            =   2520
      TabIndex        =   77
      Top             =   2760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtWord 
      Height          =   285
      Index           =   4
      Left            =   2520
      TabIndex        =   76
      Top             =   3360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtWord 
      Height          =   285
      Index           =   5
      Left            =   2520
      TabIndex        =   75
      Top             =   3960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txt2 
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
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txt2 
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
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   8
      Top             =   2400
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txt2 
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
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   9
      Top             =   2400
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txt2 
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
      Left            =   2160
      MaxLength       =   1
      TabIndex        =   10
      Top             =   2400
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txt2 
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
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   11
      Top             =   2400
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txt2 
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
      Index           =   6
      Left            =   2880
      MaxLength       =   1
      TabIndex        =   12
      Top             =   2400
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txt2 
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
      Index           =   7
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   13
      Top             =   2400
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txt3 
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
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   14
      Top             =   3000
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txt3 
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
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   15
      Top             =   3000
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txt3 
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
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   16
      Top             =   3000
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txt3 
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
      Left            =   2160
      MaxLength       =   1
      TabIndex        =   17
      Top             =   3000
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txt3 
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
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   18
      Top             =   3000
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txt3 
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
      Index           =   6
      Left            =   2880
      MaxLength       =   1
      TabIndex        =   19
      Top             =   3000
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txt3 
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
      Index           =   7
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   20
      Top             =   3000
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txt4 
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
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   21
      Top             =   3600
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txt4 
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
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   22
      Top             =   3600
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txt4 
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
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   23
      Top             =   3600
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txt4 
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
      Left            =   2160
      MaxLength       =   1
      TabIndex        =   24
      Top             =   3600
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txt4 
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
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   25
      Top             =   3600
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txt4 
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
      Index           =   6
      Left            =   2880
      MaxLength       =   1
      TabIndex        =   26
      Top             =   3600
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txt4 
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
      Index           =   7
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   27
      Top             =   3600
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txt5 
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
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   28
      Top             =   4200
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txt5 
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
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   29
      Top             =   4200
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txt5 
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
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   30
      Top             =   4200
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txt5 
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
      Left            =   2160
      MaxLength       =   1
      TabIndex        =   31
      Top             =   4200
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txt5 
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
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   32
      Top             =   4200
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txt5 
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
      Index           =   6
      Left            =   2880
      MaxLength       =   1
      TabIndex        =   33
      Top             =   4200
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txt5 
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
      Index           =   7
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   34
      Top             =   4200
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Left            =   1980
      MaxLength       =   1
      TabIndex        =   36
      Top             =   4680
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   37
      Top             =   4680
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Left            =   2580
      MaxLength       =   1
      TabIndex        =   38
      Top             =   4680
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      MaxLength       =   1
      TabIndex        =   39
      Top             =   4680
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Index           =   6
      Left            =   3180
      MaxLength       =   1
      TabIndex        =   40
      Top             =   4680
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Index           =   7
      Left            =   3480
      MaxLength       =   1
      TabIndex        =   41
      Top             =   4680
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Index           =   8
      Left            =   3765
      MaxLength       =   1
      TabIndex        =   42
      Top             =   4680
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Index           =   9
      Left            =   4065
      MaxLength       =   1
      TabIndex        =   43
      Top             =   4680
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Index           =   10
      Left            =   4365
      MaxLength       =   1
      TabIndex        =   44
      Top             =   4680
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Index           =   11
      Left            =   4665
      MaxLength       =   1
      TabIndex        =   45
      Top             =   4680
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Index           =   12
      Left            =   4965
      MaxLength       =   1
      TabIndex        =   46
      Top             =   4680
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Index           =   13
      Left            =   5265
      MaxLength       =   1
      TabIndex        =   47
      Top             =   4680
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Index           =   14
      Left            =   5565
      MaxLength       =   1
      TabIndex        =   48
      Top             =   4680
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Index           =   15
      Left            =   5865
      MaxLength       =   1
      TabIndex        =   49
      Top             =   4680
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Index           =   16
      Left            =   6165
      MaxLength       =   1
      TabIndex        =   50
      Top             =   4680
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Index           =   19
      Left            =   1140
      MaxLength       =   1
      TabIndex        =   53
      Top             =   5040
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Index           =   20
      Left            =   1425
      MaxLength       =   1
      TabIndex        =   54
      Top             =   5040
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Index           =   21
      Left            =   1725
      MaxLength       =   1
      TabIndex        =   55
      Top             =   5040
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Index           =   22
      Left            =   2020
      MaxLength       =   1
      TabIndex        =   56
      Top             =   5040
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Index           =   23
      Left            =   2320
      MaxLength       =   1
      TabIndex        =   57
      Top             =   5040
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Index           =   24
      Left            =   2620
      MaxLength       =   1
      TabIndex        =   58
      Top             =   5040
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Index           =   25
      Left            =   2920
      MaxLength       =   1
      TabIndex        =   59
      Top             =   5040
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Index           =   26
      Left            =   3220
      MaxLength       =   1
      TabIndex        =   60
      Top             =   5040
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Index           =   27
      Left            =   3520
      MaxLength       =   1
      TabIndex        =   61
      Top             =   5040
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Index           =   17
      Left            =   6465
      MaxLength       =   1
      TabIndex        =   51
      Top             =   4680
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Index           =   18
      Left            =   840
      MaxLength       =   1
      TabIndex        =   52
      Top             =   5040
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Index           =   28
      Left            =   3820
      MaxLength       =   1
      TabIndex        =   62
      Top             =   5040
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Index           =   29
      Left            =   4120
      MaxLength       =   1
      TabIndex        =   63
      Top             =   5040
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Index           =   30
      Left            =   4420
      MaxLength       =   1
      TabIndex        =   64
      Top             =   5040
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Index           =   31
      Left            =   4720
      MaxLength       =   1
      TabIndex        =   65
      Top             =   5040
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Index           =   32
      Left            =   5020
      MaxLength       =   1
      TabIndex        =   66
      Top             =   5040
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Index           =   33
      Left            =   5320
      MaxLength       =   1
      TabIndex        =   67
      Top             =   5040
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Index           =   34
      Left            =   5620
      MaxLength       =   1
      TabIndex        =   68
      Top             =   5040
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Index           =   35
      Left            =   5920
      MaxLength       =   1
      TabIndex        =   69
      Top             =   5040
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Index           =   36
      Left            =   6220
      MaxLength       =   1
      TabIndex        =   70
      Top             =   5040
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Index           =   37
      Left            =   6520
      MaxLength       =   1
      TabIndex        =   71
      Top             =   5040
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Index           =   38
      Left            =   960
      MaxLength       =   1
      TabIndex        =   72
      Top             =   5400
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Index           =   39
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   73
      Top             =   5400
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtAnswer1 
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
      Index           =   40
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   74
      Top             =   5400
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   6840
      TabIndex        =   95
      ToolTipText     =   "Exit"
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   6840
      TabIndex        =   94
      ToolTipText     =   "Exit"
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Command6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   6960
      TabIndex        =   93
      ToolTipText     =   "Exit"
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   6960
      TabIndex        =   92
      ToolTipText     =   "Exit"
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   6840
      TabIndex        =   91
      ToolTipText     =   "Exit"
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   6960
      TabIndex        =   90
      ToolTipText     =   "Exit"
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label MoveForm 
      Height          =   375
      Left            =   0
      TabIndex        =   89
      ToolTipText     =   "Right click for popup menu"
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Index           =   1
      Left            =   1080
      TabIndex        =   87
      Top             =   1540
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   960
      TabIndex        =   86
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Index           =   2
      Left            =   1080
      TabIndex        =   85
      Top             =   2140
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Index           =   3
      Left            =   1080
      TabIndex        =   84
      Top             =   2740
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Index           =   4
      Left            =   1080
      TabIndex        =   83
      Top             =   3340
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Index           =   5
      Left            =   1080
      TabIndex        =   82
      Top             =   3940
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   4200
      TabIndex        =   81
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Caption1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1335
      Left            =   4200
      TabIndex        =   80
      Top             =   1680
      Width           =   2295
   End
End
Attribute VB_Name = "frmJumble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Const RGN_OR = 2
Private lngRegion As Long
Public Prg As String, Sect As String ' for savesettings
Public SkinDir As String

Dim Words As String
Dim vbCaption As String
Dim NumLength(10) As Integer
Dim TotalLength As Integer
Dim vbLength As Integer
Dim WordLength As Integer
Function RegionFromBitmap(picSource As PictureBox, Optional lngTransColor As Long) As Long
    Dim lngRetr As Long, lngHeight As Long, lngWidth As Long
    Dim lngRgnFinal As Long, lngRgnTmp As Long
    Dim lngStart As Long, lngRow As Long
    Dim lngCol As Long
  
    If lngTransColor& < 1 Then
        lngTransColor& = GetPixel(picSource.hDC, 0, 0)
    End If
    
    lngHeight& = picSource.Height / Screen.TwipsPerPixelY
    lngWidth& = picSource.Width / Screen.TwipsPerPixelX
    lngRgnFinal& = CreateRectRgn(0, 0, 0, 0)
    For lngRow& = 0 To lngHeight& - 1
        lngCol& = 0
        Do While lngCol& < lngWidth&
            Do While lngCol& < lngWidth& And GetPixel(picSource.hDC, lngCol&, lngRow&) = lngTransColor&
                lngCol& = lngCol& + 1
            Loop
            If lngCol& < lngWidth& Then
                lngStart& = lngCol&
                Do While lngCol& < lngWidth& And GetPixel(picSource.hDC, lngCol&, lngRow&) <> lngTransColor&
                    lngCol& = lngCol& + 1
                Loop
                If lngCol& > lngWidth& Then lngCol& = lngWidth&
                    lngRgnTmp& = CreateRectRgn(lngStart&, lngRow&, lngCol&, lngRow& + 1)
                    lngRetr& = CombineRgn(lngRgnFinal&, lngRgnFinal&, lngRgnTmp&, RGN_OR)
                    DeleteObject (lngRgnTmp&)
                End If
        Loop
    Next
    RegionFromBitmap& = lngRgnFinal&
End Function

Sub ChangeMask()
    Dim lngRetr As Long
    
    On Error Resume Next ' In case of error

    lngRegion& = RegionFromBitmap(Mask)
    lngRetr& = SetWindowRgn(Me.hWnd, lngRegion&, True)
End Sub

Private Sub Command1_Click()
    
    DBCloseAll
    Unload Me
End Sub

Private Sub Command2_Click()
    Dim i As Integer
    Dim Temp As String
    Dim Answer As String
    
    Temp = ""
    For i = 1 To 40
        If txtAnswer1(i).Visible = True Then
            If txtAnswer1(i).Text = "" Then
                MsgBox "You must fill all boxes prior to solving", vbOKOnly + vbExclamation, "Fill all boxes"
                Exit Sub
            End If
        End If
    Next i
    
    For i = 1 To 40
        If txtAnswer1(i).Text = "" Then
            Temp = Temp + " "
        Else
            Temp = Temp + txtAnswer1(i).Text
        End If
    Next
    
    Answer = Data1.Recordset.Fields("Answer")
    If Left$(Temp, Len(Answer)) = Answer Then
        MsgBox "Correct. Great!!"
    Else
        MsgBox "Incorrect."
    End If
End Sub

Private Sub Command3_Click()
    Dim Answer As String
    Dim i As Integer
    Dim J As Integer
    Dim Words As String
    Dim vbSearch As Integer
    Dim Temp As String
    
    Answer = Data1.Recordset.Fields("Answer")
    Words = Data1.Recordset.Fields("Words")
    vbSearch = Left$(Words, 1)
    Words = Mid$(Words, 3, Len(Words) - 2)
    vbLength = InStr(Words, ",")
    If vbLength = 0 Then
        vbLength = Len(Words) + 1
    End If
    For J = 1 To vbLength - 1
        txt1(J).Text = Mid$(Words, J, 1)
    Next J
    If vbSearch > 1 Then
        Words = Mid$(Words, vbLength + 1, Len(Words) - vbLength)
        vbLength = InStr(Words, ",")
        If vbLength = 0 Then
            vbLength = Len(Words) + 1
        End If
        For J = 1 To vbLength - 1
            txt2(J).Text = Mid$(Words, J, 1)
        Next J
    End If
    If vbSearch > 2 Then
        Words = Mid$(Words, vbLength + 1, Len(Words) - vbLength)
        vbLength = InStr(Words, ",")
        If vbLength = 0 Then
            vbLength = Len(Words) + 1
        End If
        For J = 1 To vbLength - 1
            txt3(J).Text = Mid$(Words, J, 1)
            Temp = txt3(J).Text
        Next J
    End If
    If vbSearch > 3 Then
        Words = Mid$(Words, vbLength + 1, Len(Words) - vbLength)
        vbLength = InStr(Words, ",")
        If vbLength = 0 Then
            vbLength = Len(Words) + 1
        End If
        For J = 1 To vbLength - 1
            txt4(J).Text = Mid$(Words, J, 1)
        Next J
    End If
    If vbSearch > 4 Then
        Words = Mid$(Words, vbLength + 1, Len(Words) - vbLength)
        vbLength = InStr(Words, ",")
        If vbLength = 0 Then
            vbLength = Len(Words) + 1
        End If
        For J = 1 To vbLength - 1
            txt5(J).Text = Mid$(Words, J, 1)
        Next J
    End If
    For i = 1 To Len(Answer)
        txtAnswer1(i).Text = Mid$(Answer, i, 1)
    Next i
End Sub

Private Sub Command4_Click()

    Data1.Recordset.MoveNext
    If Data1.Recordset.EOF Then
        Data1.Recordset.MovePrevious
    End If
    GetStart
End Sub

Private Sub Command5_Click()
    Dim i As Integer
    Dim Temp As String
    Dim Answer As String
    
    Temp = txtWord(1).Text
    If txt1(1).Text = "" Then
        txt1(1).Text = Left$(Temp, 1)
        txt1(2).SetFocus
        Exit Sub
    End If

    Temp = txtWord(2).Text
    If txt2(1).Text = "" Then
        txt2(1).Text = Left$(Temp, 1)
        txt2(2).SetFocus
        Exit Sub
    End If
    
    Temp = txtWord(3).Text
    If txt3(1).Text = "" Then
        txt3(1).Text = Left$(Temp, 1)
        txt3(2).SetFocus
        Exit Sub
    End If
    
    Temp = txtWord(4).Text
    If txt4(1).Text = "" Then
        txt4(1).Text = Left$(Temp, 1)
        txt4(2).SetFocus
        Exit Sub
    End If
    
    Temp = txtWord(5).Text
    If Temp <> "" Then
        If txt5(1).Text = "" Then
            txt5(1).Text = Left$(Temp, 1)
            txt5(2).SetFocus
            Exit Sub
        End If
    End If
    
    Temp = txtWord(1).Text
    For i = 1 To Len(Temp)
        If txt1(i).Text = "" Then
            txt1(i).Text = Mid$(Temp, i, 1)
            If i < Len(txtWord(1).Text) Then
                txt1(i + 1).SetFocus
            Else
                txt2(1).SetFocus
            End If
            Exit Sub
        End If
    Next i
    
    Temp = txtWord(2).Text
    For i = 1 To Len(Temp)
        If txt2(i).Text = "" Then
            txt2(i).Text = Mid$(Temp, i, 1)
            If i < Len(txtWord(2).Text) Then
                txt2(i + 1).SetFocus
            Else
                txt3(1).SetFocus
            End If
            Exit Sub
        End If
    Next i
    
    Temp = txtWord(3).Text
    For i = 1 To Len(Temp)
        If txt3(i).Text = "" Then
            txt3(i).Text = Mid$(Temp, i, 1)
            If i < Len(txtWord(3).Text) Then
                txt3(i + 1).SetFocus
            Else
                txt4(1).SetFocus
            End If

            Exit Sub
        End If
    Next i
    
    Temp = txtWord(4).Text
    For i = 1 To Len(Temp)
        If txt4(i).Text = "" Then
            txt4(i).Text = Mid$(Temp, i, 1)
            If i < Len(txtWord(4).Text) Then
                txt4(i + 1).SetFocus
            Else
                If txtWord(5).Text <> "" Then
                    txt5(1).SetFocus
                Else
                    txt1(1).SetFocus
                End If
            End If

            Exit Sub
        End If
    Next i
    
    Temp = txtWord(5).Text
    If Temp <> "" Then
        For i = 1 To Len(Temp)
            If txt5(i).Text = "" Then
                txt5(i).Text = Mid$(Temp, i, 1)
                If i < Len(txtWord(5).Text) Then
                    txt5(i + 1).SetFocus
                Else
                    txt1(1).SetFocus
                End If
                    Exit Sub
                End If
        Next i
    End If
    
    Answer = Data1.Recordset.Fields("Answer")
    For i = 1 To Len(Answer)
        If txtAnswer1(i).Text = "" Then
            If Mid$(Answer, i, 1) <> " " Then
                txtAnswer1(i).Text = Mid$(Answer, i, 1)
                Exit Sub
            End If
        End If
    Next i
    
End Sub

Private Sub Command6_Click()
    Me.Hide
    frmAdd.Show
End Sub

Private Sub Form_Load()
    
    Data1.Refresh
    Data1.Recordset.MoveFirst
    
    On Error Resume Next ' In case of error
    Prg = "SkinDemo"
    Sect = "config" ' This is used for saving to registry
  
    SkinDir = App.Path + "\Jumble.skn"
    OpenSkin SkinDir ' opens the default skin

    MoveForm.Top = 0 ' Put lable top most
    MoveForm.Left = 0 ' Put lable at the left
    MoveForm.BackStyle = 0  ' Makes label transparent
    
    GetStart
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub

Sub GetLength()

    Select Case vbLength
        Case 1
            WordLength = 1
        Case 2
            WordLength = NumLength(1) + 1
        Case 3
            WordLength = NumLength(1) + 1 + NumLength(2) + 1
        Case 4
            WordLength = NumLength(1) + 1 + NumLength(2) + 1 + NumLength(3) + 1
        Case 5
            WordLength = NumLength(1) + 1 + NumLength(2) + 1 + NumLength(3) + 1 + NumLength(4) + 1
        Case 6
            WordLength = NumLength(1) + 1 + NumLength(2) + 1 + NumLength(3) + 1 + NumLength(4) + 1 + NumLength(5) + 1
        Case 7
            WordLength = NumLength(1) + 1 + NumLength(2) + 1 + NumLength(3) + 1 + NumLength(4) + 1 + NumLength(5) + 1 + NumLength(6) + 1
        Case 8
            WordLength = NumLength(1) + 1 + NumLength(2) + 1 + NumLength(3) + 1 + NumLength(4) + 1 + NumLength(5) + 1 + NumLength(6) + 1 + NumLength(7) + 1
        Case 9
            WordLength = NumLength(1) + 1 + NumLength(2) + 1 + NumLength(3) + 1 + NumLength(4) + 1 + NumLength(5) + 1 + NumLength(6) + 1 + NumLength(7) + 1 + NumLength(8) + 1
        Case 10
            WordLength = NumLength(1) + 1 + NumLength(2) + 1 + NumLength(3) + 1 + NumLength(4) + 1 + NumLength(5) + 1 + NumLength(6) + 1 + NumLength(7) + 1 + NumLength(8) + 1 + NumLength(9) + 1
    End Select
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim A As String
    Dim i As Integer
    
    A = UCase(Chr(KeyAscii))
    For i = 1 To Len(Label1(1).Caption)
        If A = Mid$(Label1(1).Caption, i, 1) Then
            KeyAscii = Asc(A)
            If Index < Len(txtWord(1).Text) Then
                txt1(Index + 1).SetFocus
            Else
                txt2(1).SetFocus
            End If
            Exit Sub
        End If
    Next i

    If KeyAscii = 8 Then
        Exit Sub
    End If
    
    KeyAscii = 0
End Sub

Private Sub txt2_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim A As String
    Dim i As Integer
    
    A = UCase(Chr(KeyAscii))
    For i = 1 To Len(Label1(2).Caption)
        If A = Mid$(Label1(2).Caption, i, 1) Then
            KeyAscii = Asc(A)
            If Index < Len(txtWord(2).Text) Then
                txt2(Index + 1).SetFocus
            Else
                txt3(1).SetFocus
            End If
            Exit Sub
        End If
    Next i

    If KeyAscii = 8 Then
        Exit Sub
    End If
    
    KeyAscii = 0
End Sub

Private Sub txt3_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim A As String
    Dim i As Integer
    
    A = UCase(Chr(KeyAscii))
    For i = 1 To Len(Label1(3).Caption)
        If A = Mid$(Label1(3).Caption, i, 1) Then
            KeyAscii = Asc(A)
            If Index < Len(txtWord(3).Text) Then
                txt3(Index + 1).SetFocus
            Else
                txt4(1).SetFocus
            End If
            Exit Sub
        End If
    Next i

    If KeyAscii = 8 Then
        Exit Sub
    End If
    
    KeyAscii = 0
End Sub

Private Sub txt4_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim A As String
    Dim i As Integer
    
    A = UCase(Chr(KeyAscii))
    For i = 1 To Len(Label1(4).Caption)
        If A = Mid$(Label1(4).Caption, i, 1) Then
            KeyAscii = Asc(A)
            If Index < Len(txtWord(4).Text) Then
                txt4(Index + 1).SetFocus
            Else
                If txtWord(5).Text <> "" Then
                    txt5(1).SetFocus
                Else
                    txt1(1).SetFocus
                End If
            End If
            Exit Sub
        End If
    Next i

    If KeyAscii = 8 Then
        Exit Sub
    End If
    
    KeyAscii = 0
End Sub

Private Sub txt5_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim A As String
    Dim i As Integer
    
    A = UCase(Chr(KeyAscii))
    For i = 1 To Len(Label1(5).Caption)
        If A = Mid$(Label1(5).Caption, i, 1) Then
            KeyAscii = Asc(A)
            If Index < Len(txtWord(5).Text) Then
                txt5(Index + 1).SetFocus
            Else
                txt1(1).SetFocus
            End If
            Exit Sub
        End If
    Next i

    If KeyAscii = 8 Then
        Exit Sub
    End If
    
    KeyAscii = 0
End Sub
Private Sub txtAnswer1_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim A As String
        
    A = UCase(Chr(KeyAscii))
    KeyAscii = Asc(A)
End Sub

Sub GetStart()
    Dim vbSearch As Integer
    Dim i As Integer
    Dim J As Integer
    Dim K As Integer
    Dim NumWords As Integer
    Dim vbTempWord As String
    Dim Word(10) As String
    Dim WordScrambled(10) As String
    Dim Scrambled As String
    Dim tempScrambled As String
    Dim Numbers2Puzzle As String
    Dim FullAnswer As String
    Dim WhichWord As Integer
    
    Words = Data1.Recordset.Fields("Words")
    Scrambled = Data1.Recordset.Fields("Scrambled")
    vbCaption = Data1.Recordset.Fields("Caption")
    NumWords = Val(Left$(Words, 1))
    Numbers2Puzzle = Data1.Recordset.Fields("NumbersToPuzzle")
    FullAnswer = Data1.Recordset.Fields("NumberOfWords")
    For i = 1 To 7
        txt1(i).Text = ""
        txt1(i).Visible = False
        txt2(i).Text = ""
        txt2(i).Visible = False
        txt3(i).Text = ""
        txt3(i).Visible = False
        txt4(i).Text = ""
        txt4(i).Visible = False
        txt5(i).Text = ""
        txt5(i).Visible = False
    Next
    
    For i = 1 To 40
        txtAnswer1(i).Text = ""
        txtAnswer1(i).Visible = False
    Next i
    For i = 1 To 5
        Label1(i).Caption = ""
        Label1(i).Visible = False
    Next i
    
    For i = 1 To NumWords
        Label1(i).Visible = True
        vbSearch = InStr(Words, ",")
        vbTempWord = Mid$(Words, vbSearch + 1, Len(Words) - vbSearch)
        tempScrambled = Mid$(Scrambled, vbSearch + 1, Len(Scrambled) - vbSearch)
        vbLength = InStr(vbTempWord, ",") - 1
        If vbLength > 0 Then
            Word(i) = Left$(vbTempWord, vbLength)
            WordScrambled(i) = Left$(tempScrambled, vbLength)
        Else
            Word(i) = vbTempWord
            WordScrambled(i) = tempScrambled
            i = NumWords
        End If
        txtWord(i).Text = Word(i)
        Label1(i).Caption = WordScrambled(i)
        If i < NumWords Then
            Words = Mid$(vbTempWord, vbLength + 1, Len(vbTempWord) - 1)
            Scrambled = Mid$(tempScrambled, vbLength + 1, Len(tempScrambled) - 1)
        End If
        If i = 5 Then
            Label1(i).Visible = True
        End If
    Select Case i
        Case 1
            For J = 1 To Len(txtWord(i).Text)
                txt1(J).Visible = True
                txt1(J).BackColor = &HFFFFFF
            Next J
        Case 2
            For J = 1 To Len(txtWord(i).Text)
                txt2(J).Visible = True
                txt2(J).BackColor = &HFFFFFF
            Next J
        Case 3
            For J = 1 To Len(txtWord(i).Text)
                txt3(J).Visible = True
                txt3(J).BackColor = &HFFFFFF
            Next J
        Case 4
            For J = 1 To Len(txtWord(i).Text)
                txt4(J).Visible = True
                txt4(J).BackColor = &HFFFFFF
            Next J
        Case 5
            For J = 1 To Len(txtWord(i).Text)
                txt5(J).Visible = True
                txt5(J).BackColor = &HFFFFFF
            Next J
    End Select
    Next i
    
    vbLength = 2
    For i = 1 To NumWords
        vbSearch = InStr(Numbers2Puzzle, "-")
        NumWords = Val(Left$(Numbers2Puzzle, 1))
        For J = 1 To NumWords
            K = Val(Mid$(Numbers2Puzzle, vbLength + 1, 1))
            vbLength = vbLength + 2
            Select Case i
                Case 1
                    txt1(K).BackColor = &HC0C0C0
                    'txt1(K).ForeColor = &H0
                Case 2
                    txt2(K).BackColor = &HC0C0C0
                    'txt1(K).ForeColor = &H0
                Case 3
                    txt3(K).BackColor = &HC0C0C0
                    'txt1(K).ForeColor = &H0
                Case 4
                    txt4(K).BackColor = &HC0C0C0
                    'txt1(K).ForeColor = &H0
                Case 5
                    txt5(K).BackColor = &HC0C0C0
                    'txt1(K).ForeColor = &H0
                Case 6
                    'txt6(K).BackColor = &H80FFFF
            End Select
        
        Next J
        Numbers2Puzzle = Mid$(Numbers2Puzzle, vbSearch + 1, Len(Numbers2Puzzle) - vbSearch + 1)
        vbLength = 2
    Next i
    Caption1.Caption = vbCaption
    NumWords = Left$(FullAnswer, 1)
    vbTempWord = Data1.Recordset.Fields("Lengths")
    TotalLength = 0
    J = 1
    For i = 1 To NumWords * 2 Step 2
        NumLength(J) = Mid$(vbTempWord, i, 1)
        TotalLength = TotalLength + NumLength(J)
        J = J + 1
    Next i
    K = 1
    For i = 1 To NumWords
        For J = 1 To NumLength(i)
            txtAnswer1(K).Visible = True
            K = K + 1
        Next J
        K = K + 1
    Next i
    
    vbSearch = InStr(FullAnswer, ",")
    If vbSearch > 0 Then
        vbLength = Mid$(FullAnswer, 3, 1)
    Else
        GoTo Exit1
    End If
    
    vbTempWord = Mid$(FullAnswer, 5, Len(FullAnswer) - 4)
    vbSearch = InStr(vbTempWord, ",")
    TotalLength = vbSearch - 1
    If TotalLength < 0 Then
        TotalLength = Len(vbTempWord)
        'vbLength = TotalLength
    End If
    J = 1
    GetLength
    If vbLength = 1 Then
        For i = WordLength To (WordLength + TotalLength) - 1
            txtAnswer1(i).Text = Mid$(vbTempWord, J, 1)
            txtAnswer1(i).Locked = True
            txtAnswer1(i).TabStop = False
            J = J + 1
        Next i
    Else
        J = 1
        For i = WordLength + 1 To WordLength + TotalLength
           txtAnswer1(i).Text = Mid(vbTempWord, J, 1)
           txtAnswer1(i).Locked = True
           txtAnswer1(i).TabStop = False
           J = J + 1
        Next i
    End If
    vbTempWord = Mid$(vbTempWord, vbSearch + 1, Len(vbTempWord) - vbSearch)
    vbSearch = InStr(vbTempWord, ",")
    If vbSearch = 0 Then
        GoTo Exit1
    End If
    vbLength = Left$(vbTempWord, 1)
    vbTempWord = Mid$(vbTempWord, vbSearch + 1, Len(vbTempWord) - 1)
    GetLength
    J = 1
    For i = WordLength + 1 To WordLength + Len(vbTempWord)
       txtAnswer1(i).Text = Mid(vbTempWord, J, 1)
       txtAnswer1(i).Locked = True
       txtAnswer1(i).TabStop = False
       J = J + 1
    Next i
    
    vbSearch = InStr(vbTempWord, ",")
    If vbSearch = 0 Then
        GoTo Exit1
    End If
    vbLength = Left$(vbTempWord, 1)
    vbTempWord = Mid$(vbTempWord, vbSearch + 1, Len(vbTempWord) - 1)
    GetLength
    J = 1
    For i = WordLength + 1 To WordLength + Len(vbTempWord)
       txtAnswer1(i).Text = Mid(vbTempWord, J, 1)
       txtAnswer1(i).Locked = True
       txtAnswer1(i).TabStop = False
       J = J + 1
    Next i
    
Exit1:

End Sub
