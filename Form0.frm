VERSION 5.00
Begin VB.Form Form0 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "һ�ʻ�"
   ClientHeight    =   6210
   ClientLeft      =   8025
   ClientTop       =   2550
   ClientWidth     =   4545
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form0.frx":0000
   ScaleHeight     =   6210
   ScaleWidth      =   4545
   Begin VB.CommandButton Command4 
      Caption         =   "�˳���Ϸ"
      Height          =   435
      Left            =   900
      TabIndex        =   3
      Top             =   5340
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ת��Ƥ��"
      Height          =   435
      Left            =   900
      TabIndex        =   2
      Top             =   4500
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ѡ��ؿ�"
      Height          =   435
      Left            =   900
      TabIndex        =   1
      Top             =   3600
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���¿�ʼ"
      Height          =   435
      Left            =   900
      TabIndex        =   0
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "һ �� ��"
      BeginProperty Font 
         Name            =   "�����п�"
         Size            =   42
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   780
      TabIndex        =   4
      Top             =   1260
      Width           =   3255
   End
End
Attribute VB_Name = "Form0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
Form1.Show

End Sub

Private Sub Command2_Click()
Form11.Show
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Form_Load()
Me.Left = 7980
Me.Top = 2175
End Sub

