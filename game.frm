VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "第一关"
   ClientHeight    =   6210
   ClientLeft      =   8025
   ClientTop       =   2550
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   310.5
   ScaleMode       =   2  'Point
   ScaleWidth      =   227.25
   Begin VB.CommandButton Command2 
      Caption         =   "返回首页"
      Height          =   435
      Left            =   3180
      TabIndex        =   4
      Top             =   4920
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "重玩"
      Height          =   435
      Left            =   480
      TabIndex        =   3
      Top             =   4920
      Width           =   915
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "—第一关—"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   4980
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   375
      Index           =   0
      Left            =   900
      Shape           =   3  'Circle
      Top             =   4260
      Width           =   315
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   375
      Index           =   2
      Left            =   3300
      Shape           =   3  'Circle
      Top             =   1800
      Width           =   315
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   375
      Index           =   1
      Left            =   900
      Shape           =   3  'Circle
      Top             =   1800
      Width           =   315
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   10
      Index           =   20
      X1              =   174
      X2              =   54
      Y1              =   102
      Y2              =   222
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   10
      X1              =   51
      X2              =   51
      Y1              =   99
      Y2              =   219
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   10
      Index           =   21
      X1              =   54
      X2              =   174
      Y1              =   99
      Y2              =   99
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   1
      Left            =   900
      TabIndex        =   1
      Top             =   1860
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   2
      Left            =   3300
      TabIndex        =   2
      Top             =   1860
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   4260
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim firstpoint As Boolean
Dim index2 As Integer
Dim n As Integer




Private Sub Command1_Click()
firstpoint = True
n = 0
For i = 0 To 2
Shape1(i).BackColor = &H404040
Next i
Line1(10).BorderColor = &H404040
Line1(20).BorderColor = &H404040
Line1(21).BorderColor = &H404040
End Sub

Private Sub Command2_Click()
Form0.Show
Unload Me
End Sub

Private Sub Form_Load()
firstpoint = True
n = 0
Me.Height = 6630
Me.Width = 4635
Me.Left = 7980
Me.Top = 2175
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
q = MsgBox("您确定要退出游戏？", vbYesNo + 48, "退出")
   If q = 6 Then
   Unload Form0
   Else
   Cancel = 1
   End If
End If
End Sub

Private Sub Label1_Click(Index As Integer)
If firstpoint Then
   Shape1(Index).BackColor = &HC0C000
   firstpoint = False
   index2 = Index
Else
   Select Case tpl(Index, index2)
   Case 10, 20, 21
      If Not Line1(tpl(index2, Index)).BorderColor = &HC0C000 Then
         Line1(tpl(index2, Index)).BorderColor = &HC0C000
         Shape1(Index).BackColor = &HC0C000
         index2 = Index
         n = n + 1
      Else
         Exit Sub
      End If
   End Select
   If n = 3 Then
    ns = MsgBox("...成功...按下yes进入下一关；按下no回到首页", vbYesNoCancel, "恭喜你闯关成功")
    Select Case ns
    Case 6
      Form2.Show
      Unload Me
    Case 7
      Form0.Show
      Unload Me
    End Select
   End If
End If
End Sub

