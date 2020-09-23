VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   5385
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   2160
      Width           =   5175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   5175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mohsen Hosseini
'Mohsen_hosseyni@hotmail.com
'Ascii To UniCode

Private Sub Command1_Click()

Text1.Text = CUni(Text3.Text)
'Text1.Text = Text1.Text + ChrW(&H698)
End Sub

Private Sub Command2_Click()
Text3.Text = ""
Text1.Text = ""
End Sub

Private Sub Form_Load()
    Command1.Caption = CUni("jfndg")
    Command2.Caption = CUni("`h; ;vnk gdsj")
End Sub
