VERSION 5.00
Begin VB.Form Form11 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "关卡选择"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   1665
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   1665
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "华文彩云"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      ItemData        =   "Form11.frx":0000
      Left            =   120
      List            =   "Form11.frx":0022
      TabIndex        =   0
      Top             =   180
      Width           =   1395
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Drive1_Change()

End Sub

Private Sub Form_Load()
Me.Top = Form0.Top
Me.Left = Form0.Left + Form0.Width

End Sub

Private Sub List1_Click()
Dim a As Integer
a = List1.listindex
Select Case a
Case 0
Form1.Show
Form0.Hide
Unload Me
Case 1
Form2.Show
Form0.Hide
Unload Me
Case 2
Form3.Show
Form0.Hide
Unload Me
Case 3
Form4.Show
Form0.Hide
Unload Me
Case 4
Form5.Show
Form0.Hide
Unload Me
Case 5
Form6.Show
Form0.Hide
Unload Me
Case 6
Form7.Show
Form0.Hide
Unload Me
Case 7
Form8.Show
Form0.Hide
Unload Me
Case 8
Form9.Show
Form0.Hide
Unload Me
Case 9
Form10.Show
Form0.Hide
Unload Me
End Select

End Sub
