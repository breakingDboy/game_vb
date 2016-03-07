VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "第五关"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4545
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "返回首页"
      Height          =   435
      Left            =   3180
      TabIndex        =   1
      Top             =   5280
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "重玩"
      Height          =   435
      Left            =   480
      TabIndex        =   0
      Top             =   5280
      Width           =   915
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   435
      Index           =   6
      Left            =   1200
      TabIndex        =   9
      Top             =   1560
      Width           =   435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   435
      Index           =   5
      Left            =   240
      TabIndex        =   8
      Top             =   2280
      Width           =   435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   435
      Index           =   4
      Left            =   240
      TabIndex        =   7
      Top             =   4620
      Width           =   435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   435
      Index           =   3
      Left            =   2760
      TabIndex        =   6
      Top             =   4680
      Width           =   435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   435
      Index           =   2
      Left            =   3660
      TabIndex        =   5
      Top             =   3960
      Width           =   435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   435
      Index           =   1
      Left            =   3600
      TabIndex        =   4
      Top             =   1560
      Width           =   435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   435
      Index           =   0
      Left            =   2700
      TabIndex        =   3
      Top             =   2340
      Width           =   435
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   435
      Index           =   6
      Left            =   1200
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   435
      Index           =   5
      Left            =   300
      Shape           =   3  'Circle
      Top             =   2340
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   435
      Index           =   4
      Left            =   300
      Shape           =   3  'Circle
      Top             =   4620
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   435
      Index           =   3
      Left            =   2760
      Shape           =   3  'Circle
      Top             =   4680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   435
      Index           =   2
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   0
      Left            =   2760
      Shape           =   3  'Circle
      Top             =   2340
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   435
      Index           =   1
      Left            =   3660
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   40
      X1              =   2880
      X2              =   420
      Y1              =   2580
      Y2              =   4920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   53
      X1              =   2940
      X2              =   420
      Y1              =   4860
      Y2              =   2520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   61
      X1              =   3900
      X2              =   1380
      Y1              =   1740
      Y2              =   1740
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   65
      X1              =   1380
      X2              =   420
      Y1              =   1740
      Y2              =   2460
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   32
      X1              =   3900
      X2              =   3000
      Y1              =   4200
      Y2              =   4920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   10
      X1              =   3900
      X2              =   2940
      Y1              =   1800
      Y2              =   2520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   21
      X1              =   3900
      X2              =   3900
      Y1              =   1800
      Y2              =   4200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   50
      X1              =   3000
      X2              =   420
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   43
      X1              =   3000
      X2              =   420
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   30
      X1              =   3000
      X2              =   3000
      Y1              =   2520
      Y2              =   4920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   54
      X1              =   420
      X2              =   420
      Y1              =   2520
      Y2              =   4920
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "—第五关—"
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
      Left            =   1620
      TabIndex        =   2
      Top             =   5340
      Width           =   1335
   End
End
Attribute VB_Name = "Form5"
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
For i = 0 To 6
Shape1(i).BackColor = &H404040
Next i
Line1(10).BorderColor = &H404040
Line1(30).BorderColor = &H404040
Line1(50).BorderColor = &H404040
Line1(61).BorderColor = &H404040
Line1(21).BorderColor = &H404040
Line1(32).BorderColor = &H404040
Line1(65).BorderColor = &H404040
Line1(54).BorderColor = &H404040
Line1(43).BorderColor = &H404040
Line1(40).BorderColor = &H404040
Line1(53).BorderColor = &H404040
Line1(43).BorderColor = &H404040
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
   Case 10, 50, 61, 30, 65, 21, 32, 53, 40, 43, 54
      If Not Line1(tpl(index2, Index)).BorderColor = &HC0C000 Then
         Line1(tpl(index2, Index)).BorderColor = &HC0C000
         Shape1(Index).BackColor = &HC0C000
         index2 = Index
         n = n + 1
      Else
         Exit Sub
      End If
   End Select
   If n = 11 Then
    ns = MsgBox("...成功...按下yes进入下一关；按下no回到首页", vbYesNoCancel, "恭喜你闯关成功")
    Select Case ns
    Case 6
      Form6.Show
      Unload Me
    Case 7
      Form0.Show
      Unload Me
    End Select
   End If
End If
End Sub

