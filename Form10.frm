VERSION 5.00
Begin VB.Form Form10 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "第十关"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4545
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   4545
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "返回首页"
      Height          =   435
      Left            =   3120
      TabIndex        =   1
      Top             =   5280
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "重玩"
      Height          =   435
      Left            =   540
      TabIndex        =   0
      Top             =   5280
      Width           =   915
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   6
      Left            =   1320
      TabIndex        =   9
      Top             =   1980
      Width           =   315
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   5
      Left            =   420
      TabIndex        =   8
      Top             =   3780
      Width           =   315
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   4
      Left            =   1380
      TabIndex        =   7
      Top             =   3840
      Width           =   315
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   3
      Left            =   2580
      TabIndex        =   6
      Top             =   3840
      Width           =   315
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   2
      Left            =   3660
      TabIndex        =   5
      Top             =   3780
      Width           =   315
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   1
      Left            =   2640
      TabIndex        =   4
      Top             =   1980
      Width           =   315
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   3
      Top             =   780
      Width           =   315
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   435
      Index           =   6
      Left            =   1320
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   315
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   435
      Index           =   5
      Left            =   420
      Shape           =   3  'Circle
      Top             =   3780
      Width           =   315
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   435
      Index           =   4
      Left            =   1380
      Shape           =   3  'Circle
      Top             =   3780
      Width           =   315
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   435
      Index           =   3
      Left            =   2580
      Shape           =   3  'Circle
      Top             =   3780
      Width           =   315
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   435
      Index           =   2
      Left            =   3660
      Shape           =   3  'Circle
      Top             =   3780
      Width           =   315
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   435
      Index           =   1
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   315
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   435
      Index           =   0
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   720
      Width           =   315
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   43
      X1              =   2580
      X2              =   1620
      Y1              =   4020
      Y2              =   4020
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   32
      X1              =   3780
      X2              =   2820
      Y1              =   4020
      Y2              =   4020
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   10
      X1              =   2820
      X2              =   2100
      Y1              =   2160
      Y2              =   900
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   60
      X1              =   2100
      X2              =   1440
      Y1              =   840
      Y2              =   2160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   63
      X1              =   2700
      X2              =   1560
      Y1              =   4020
      Y2              =   2100
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   41
      X1              =   1560
      X2              =   2700
      Y1              =   3960
      Y2              =   2100
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   31
      X1              =   2760
      X2              =   2760
      Y1              =   4020
      Y2              =   2100
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   64
      X1              =   1500
      X2              =   1500
      Y1              =   4020
      Y2              =   2100
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   61
      X1              =   2760
      X2              =   1500
      Y1              =   2100
      Y2              =   2100
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   21
      X1              =   3840
      X2              =   2820
      Y1              =   4020
      Y2              =   2160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   65
      X1              =   1440
      X2              =   540
      Y1              =   2160
      Y2              =   4020
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   54
      X1              =   1500
      X2              =   540
      Y1              =   4020
      Y2              =   4020
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "—第十关—"
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
      TabIndex        =   2
      Top             =   5340
      Width           =   1335
   End
End
Attribute VB_Name = "Form10"
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
Line1(60).BorderColor = &H404040
Line1(10).BorderColor = &H404040
Line1(61).BorderColor = &H404040
Line1(65).BorderColor = &H404040
Line1(64).BorderColor = &H404040
Line1(63).BorderColor = &H404040
Line1(32).BorderColor = &H404040
Line1(41).BorderColor = &H404040
Line1(31).BorderColor = &H404040
Line1(21).BorderColor = &H404040
Line1(54).BorderColor = &H404040
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
   Case 60, 10, 61, 65, 64, 63, 31, 21, 54, 43, 32, 41
      If Not Line1(tpl(index2, Index)).BorderColor = &HC0C000 Then
         Line1(tpl(index2, Index)).BorderColor = &HC0C000
         Shape1(Index).BackColor = &HC0C000
         index2 = Index
         n = n + 1
      Else
         Exit Sub
      End If
   End Select
   If n = 12 Then
    ns = MsgBox("...成功...恭喜您通关..~\(≧▽≦)/~。。喵。。", vbYesNoCancel, "恭喜你闯关成功")
    Select Case ns
      Case 6, 7
      Form0.Show
      Unload Me
    End Select
   End If
End If
End Sub
