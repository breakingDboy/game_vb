VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "第六关"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4545
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   4545
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "返回首页"
      Height          =   435
      Left            =   3060
      TabIndex        =   1
      Top             =   5340
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "重玩"
      Height          =   435
      Left            =   540
      TabIndex        =   0
      Top             =   5340
      Width           =   915
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   3
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   4260
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   0
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   3300
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   2
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   2400
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   6
      Left            =   1560
      Shape           =   3  'Circle
      Top             =   2340
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   5
      Left            =   420
      Shape           =   3  'Circle
      Top             =   4260
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   4
      Left            =   2700
      Shape           =   3  'Circle
      Top             =   4260
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   43
      X1              =   4020
      X2              =   2940
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   1
      Left            =   2100
      Shape           =   3  'Circle
      Top             =   1500
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   21
      X1              =   2880
      X2              =   2340
      Y1              =   2580
      Y2              =   1740
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   435
      Index           =   6
      Left            =   1560
      TabIndex        =   9
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   435
      Index           =   1
      Left            =   2100
      TabIndex        =   4
      Top             =   1500
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   435
      Index           =   5
      Left            =   420
      TabIndex        =   8
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   435
      Index           =   0
      Left            =   2340
      TabIndex        =   3
      Top             =   3420
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   435
      Index           =   4
      Left            =   2700
      TabIndex        =   7
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   435
      Index           =   3
      Left            =   3840
      TabIndex        =   6
      Top             =   4260
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   435
      Index           =   2
      Left            =   2640
      TabIndex        =   5
      Top             =   2400
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   40
      X1              =   2880
      X2              =   2340
      Y1              =   4500
      Y2              =   3600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   61
      X1              =   2160
      X2              =   1680
      Y1              =   1740
      Y2              =   2580
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   20
      X1              =   2820
      X2              =   2340
      Y1              =   2580
      Y2              =   3600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   62
      X1              =   2760
      X2              =   1740
      Y1              =   2580
      Y2              =   2580
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   60
      X1              =   2280
      X2              =   1740
      Y1              =   3480
      Y2              =   2580
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   32
      X1              =   4080
      X2              =   2940
      Y1              =   4560
      Y2              =   2700
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   65
      X1              =   1680
      X2              =   540
      Y1              =   2580
      Y2              =   4500
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   54
      X1              =   2880
      X2              =   540
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "—第六关—"
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
      Top             =   5400
      Width           =   1335
   End
End
Attribute VB_Name = "Form6"
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
Line1(20).BorderColor = &H404040
Line1(21).BorderColor = &H404040
Line1(62).BorderColor = &H404040
Line1(61).BorderColor = &H404040
Line1(65).BorderColor = &H404040
Line1(54).BorderColor = &H404040
Line1(60).BorderColor = &H404040
Line1(32).BorderColor = &H404040
Line1(40).BorderColor = &H404040
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
   Case 20, 40, 60, 62, 61, 21, 32, 43, 54, 65
      If Not Line1(tpl(index2, Index)).BorderColor = &HC0C000 Then
         Line1(tpl(index2, Index)).BorderColor = &HC0C000
         Shape1(Index).BackColor = &HC0C000
         index2 = Index
         n = n + 1
      Else
         Exit Sub
      End If
   End Select
   If n = 10 Then
    ns = MsgBox("...成功...按下yes进入下一关；按下no回到首页", vbYesNoCancel, "恭喜你闯关成功")
    Select Case ns
    Case 6
      Form7.Show
      Unload Me
    Case 7
      Form0.Show
      Unload Me
    End Select
   End If
End If
End Sub

