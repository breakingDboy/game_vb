VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "第七关"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4545
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   4545
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "返回首页"
      Height          =   435
      Left            =   3000
      TabIndex        =   1
      Top             =   5160
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "重玩"
      Height          =   435
      Left            =   420
      TabIndex        =   0
      Top             =   5160
      Width           =   915
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   9
      Left            =   1380
      TabIndex        =   12
      Top             =   780
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   8
      Left            =   420
      TabIndex        =   11
      Top             =   780
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   7
      Left            =   780
      TabIndex        =   10
      Top             =   2220
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   6
      Left            =   420
      TabIndex        =   9
      Top             =   3540
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   5
      Left            =   3660
      TabIndex        =   8
      Top             =   3540
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   4
      Left            =   3180
      TabIndex        =   7
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   3
      Left            =   3600
      TabIndex        =   6
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   2
      Left            =   2580
      TabIndex        =   5
      Top             =   780
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   4
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   0
      Left            =   1920
      TabIndex        =   3
      Top             =   3900
      Width           =   435
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   555
      Index           =   9
      Left            =   1380
      Shape           =   3  'Circle
      Top             =   720
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   555
      Index           =   8
      Left            =   420
      Shape           =   3  'Circle
      Top             =   720
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   555
      Index           =   7
      Left            =   780
      Shape           =   3  'Circle
      Top             =   2100
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   555
      Index           =   6
      Left            =   420
      Shape           =   3  'Circle
      Top             =   3420
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   555
      Index           =   0
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   3840
      Width           =   435
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   555
      Index           =   5
      Left            =   3660
      Shape           =   3  'Circle
      Top             =   3420
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   555
      Index           =   4
      Left            =   3180
      Shape           =   3  'Circle
      Top             =   2160
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   555
      Index           =   3
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   720
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   2
      Left            =   2580
      Shape           =   3  'Circle
      Top             =   720
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   555
      Index           =   1
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   120
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   32
      X1              =   3780
      X2              =   2820
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   92
      X1              =   2760
      X2              =   1620
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   21
      X1              =   2280
      X2              =   2760
      Y1              =   360
      Y2              =   900
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   91
      X1              =   2160
      X2              =   1620
      Y1              =   360
      Y2              =   900
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   50
      X1              =   2220
      X2              =   3780
      Y1              =   4140
      Y2              =   3720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   60
      X1              =   2100
      X2              =   660
      Y1              =   4140
      Y2              =   3720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   54
      X1              =   3840
      X2              =   3360
      Y1              =   3720
      Y2              =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   76
      X1              =   960
      X2              =   540
      Y1              =   2340
      Y2              =   3720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   20
      X1              =   2760
      X2              =   2160
      Y1              =   1020
      Y2              =   4140
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   90
      X1              =   2100
      X2              =   1560
      Y1              =   4140
      Y2              =   1020
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   42
      X1              =   3360
      X2              =   2760
      Y1              =   2400
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   97
      X1              =   1560
      X2              =   960
      Y1              =   1020
      Y2              =   2340
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   43
      X1              =   3840
      X2              =   3420
      Y1              =   960
      Y2              =   2340
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   87
      X1              =   960
      X2              =   540
      Y1              =   2400
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   98
      X1              =   1560
      X2              =   600
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   74
      X1              =   3420
      X2              =   960
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "—第七关—"
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
      Left            =   1500
      TabIndex        =   2
      Top             =   5280
      Width           =   1335
   End
End
Attribute VB_Name = "Form7"
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
For i = 0 To 9
Shape1(i).BackColor = &H404040
Next i
Line1(20).BorderColor = &H404040
Line1(60).BorderColor = &H404040
Line1(50).BorderColor = &H404040
Line1(90).BorderColor = &H404040
Line1(21).BorderColor = &H404040
Line1(91).BorderColor = &H404040
Line1(92).BorderColor = &H404040
Line1(98).BorderColor = &H404040
Line1(32).BorderColor = &H404040
Line1(87).BorderColor = &H404040
Line1(97).BorderColor = &H404040
Line1(42).BorderColor = &H404040
Line1(43).BorderColor = &H404040
Line1(76).BorderColor = &H404040
Line1(54).BorderColor = &H404040
Line1(74).BorderColor = &H404040

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
   Case 91, 21, 98, 92, 32, 87, 97, 90, 20, 43, 74, 76, 60, 50, 54, 42
      If Not Line1(tpl(index2, Index)).BorderColor = &HC0C000 Then
         Line1(tpl(index2, Index)).BorderColor = &HC0C000
         Shape1(Index).BackColor = &HC0C000
         index2 = Index
         n = n + 1
      Else
         Exit Sub
      End If
   End Select
   If n = 16 Then
    ns = MsgBox("...成功...按下yes进入下一关；按下no回到首页", vbYesNoCancel, "恭喜你闯关成功")
    Select Case ns
    Case 6
      Form8.Show
      Unload Me
    Case 7
      Form0.Show
      Unload Me
    End Select
   End If
End If
End Sub

