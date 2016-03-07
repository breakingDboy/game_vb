VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "第四关"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4395
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   4395
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "返回首页"
      Height          =   435
      Left            =   3000
      TabIndex        =   1
      Top             =   5220
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "重玩"
      Height          =   435
      Left            =   480
      TabIndex        =   0
      Top             =   5220
      Width           =   915
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   435
      Index           =   4
      Left            =   300
      TabIndex        =   7
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   435
      Index           =   3
      Left            =   780
      TabIndex        =   6
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   435
      Index           =   2
      Left            =   3120
      TabIndex        =   5
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   435
      Index           =   1
      Left            =   3600
      TabIndex        =   4
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   435
      Index           =   0
      Left            =   1920
      TabIndex        =   3
      Top             =   720
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   4
      Left            =   360
      Shape           =   3  'Circle
      Top             =   1740
      Width           =   315
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   3
      Left            =   840
      Shape           =   3  'Circle
      Top             =   3780
      Width           =   315
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   2
      Left            =   3180
      Shape           =   3  'Circle
      Top             =   3840
      Width           =   315
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   1
      Left            =   3660
      Shape           =   3  'Circle
      Top             =   1740
      Width           =   315
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   0
      Left            =   1980
      Shape           =   3  'Circle
      Top             =   660
      Width           =   315
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   42
      X1              =   3300
      X2              =   480
      Y1              =   4080
      Y2              =   1980
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   20
      X1              =   3300
      X2              =   2100
      Y1              =   4080
      Y2              =   840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   30
      X1              =   2100
      X2              =   960
      Y1              =   840
      Y2              =   4020
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   31
      X1              =   3840
      X2              =   960
      Y1              =   1980
      Y2              =   4080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   41
      X1              =   3780
      X2              =   540
      Y1              =   1980
      Y2              =   1980
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "—第四关—"
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
      Left            =   1560
      TabIndex        =   2
      Top             =   5280
      Width           =   1335
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim firstpoint As Boolean
Dim index2 As Integer
Dim n As Integer
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
Private Sub Command1_Click()
firstpoint = True
n = 0
For i = 0 To 4
Shape1(i).BackColor = &H404040
Next i
Line1(20).BorderColor = &H404040
Line1(42).BorderColor = &H404040
Line1(41).BorderColor = &H404040
Line1(31).BorderColor = &H404040
Line1(30).BorderColor = &H404040
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

Private Sub Label1_Click(Index As Integer)
If firstpoint Then
   Shape1(Index).BackColor = &HC0C000
   firstpoint = False
   index2 = Index
Else
   Select Case tpl(Index, index2)
   Case 20, 42, 41, 31, 30
      If Not Line1(tpl(index2, Index)).BorderColor = &HC0C000 Then
         Line1(tpl(index2, Index)).BorderColor = &HC0C000
         Shape1(Index).BackColor = &HC0C000
         index2 = Index
         n = n + 1
      Else
         Exit Sub
      End If
   End Select
   If n = 5 Then
    ns = MsgBox("...成功...按下yes进入下一关；按下no回到首页", vbYesNoCancel, "恭喜你闯关成功")
    Select Case ns
    Case 6
      Form5.Show
      Unload Me
    Case 7
      Form0.Show
      Unload Me
    End Select
   End If
End If
End Sub


