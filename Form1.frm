VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "小游戏 - 井字过三关"
   ClientHeight    =   3840
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   4860
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   735
      Left            =   3480
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Start 
      Caption         =   "开始"
      Height          =   615
      Left            =   3480
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3500
      Left            =   120
      ScaleHeight     =   3465
      ScaleWidth      =   3165
      TabIndex        =   0
      Top             =   120
      Width           =   3200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   90
   End
   Begin VB.Menu Setting 
      Caption         =   "选项"
      Begin VB.Menu VSofCom 
         Caption         =   " 1P VS Com"
      End
      Begin VB.Menu GamerVBGamer 
         Caption         =   " 1P VB 2P"
      End
      Begin VB.Menu Nothing 
         Caption         =   "-"
      End
      Begin VB.Menu ThanksTo 
         Caption         =   " 鸣谢"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim GameStart As Boolean
Dim S(3, 3) As Long                 ''''注意：1代表绿子，2代表红子，0代表没有子
Dim PutChess As Long                ''''注意：1代表圈，2代表叉
Dim VSToCom As Boolean              ''''注意：PutChess中  1代表圈，2代表叉(电脑)

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
GameStart = False
Label1.Caption = "玩家对战"
End Sub

Private Sub Start_Click()
Picture1.Cls

Picture1.Line (100, Picture1.Height / 3)-(Picture1.Width - 200, Picture1.Height / 3)
Picture1.Line (100, (Picture1.Height / 3) * 2)-(Picture1.Width - 200, (Picture1.Height / 3) * 2)
Picture1.Line (Picture1.Width / 3, 100)-(Picture1.Width / 3, Picture1.Height - 200)
Picture1.Line ((Picture1.Width / 3) * 2, 100)-((Picture1.Width / 3) * 2, Picture1.Height - 200)
GameStart = True
PutChess = 1   ''圈走先
For i = 1 To 3
For j = 1 To 3
S(i, j) = 0
Next
Next

Start.Caption = "重新开始"
End Sub

Private Function ChackWinning() As Boolean         '''''''    ((检测是否有人胜利了  ))

For i = 1 To 3         ''' 对于 横，竖 型胜利判断
  If S(i, 1) = 1 And S(i, 2) = 1 And S(i, 3) = 1 Then    ''''没有办法啦，只能这样
  Winning 1
  ChackWinning = True                      ''''' 返回True 可以告诉用户下一步该结束了
  Exit Function                            ''''' 一来可以提高点速度，二来是为了防止判断出错
  ElseIf S(i, 1) = 2 And S(i, 2) = 2 And S(i, 3) = 2 Then
  Winning 2
    ChackWinning = True
  Exit Function
  ElseIf S(1, i) = 1 And S(2, i) = 1 And S(3, i) = 1 Then
  Winning 1
    ChackWinning = True
  Exit Function
  ElseIf S(1, i) = 2 And S(2, i) = 2 And S(3, i) = 2 Then
  Winning 2
    ChackWinning = True
  Exit Function
  End If
Next

''''   对于对角线的判断
If S(1, 1) = 1 And S(2, 2) = 1 And S(3, 3) = 1 Then
Winning 1
  ChackWinning = True
Exit Function
ElseIf S(1, 1) = 2 And S(2, 2) = 2 And S(3, 3) = 2 Then
Winning 2
  ChackWinning = True
Exit Function
ElseIf S(1, 3) = 1 And S(2, 2) = 1 And S(3, 1) = 1 Then
Winning 1
  ChackWinning = True
Exit Function
ElseIf S(1, 3) = 2 And S(2, 2) = 2 And S(3, 1) = 2 Then
Winning 2
  ChackWinning = True
Exit Function
End If

Dim Sum As Long      ''计数器，用于计算S区域中是否放满了子
  For j = 1 To 3
    For k = 1 To 3
    If S(j, k) <> 0 Then Sum = Sum + 1
    Next
  Next
If Sum = 9 Then
MsgBox "平局"
Start.Caption = "开始"
 GameStart = False
  ChackWinning = True
Picture1.Cls
End If

End Function

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If GameStart = True Then
 If SetChess(X, Y, PutChess) = 0 Then        ''''成功放子
 Drowing X, Y, PutChess '''成功画子
  If ChackWinning = True Then Exit Sub     '''' 赢了的话就退出，不然程序会出错
  If VSToCom = True And PutChess = 2 Then MainCom         '''''      <---  This
 End If
End If
End Sub
''''                  "Can doing
Private Sub Drowing(ByVal X As Long, ByVal Y As Long, ByVal Chess As Long)         '''' 画图过程
 If Chess = 1 Then
 DCir X, Y  ''''要画出圈
 PutChess = 2
 ElseIf Chess = 2 Then
 DX X, Y        ''''要画出叉
 PutChess = 1
 End If
End Sub
Private Sub DCir(ByVal X As Long, ByVal Y As Long)
''''   S(1,1)
If X < Picture1.Width / 3 And Y < Picture1.Height / 3 Then    ''' 在第一条线以上
Picture1.Circle ((Picture1.Width / 3) / 2, (Picture1.Height / 3) / 2), 400, 0    ''''Picture1.Width / 3) / 2   Y轴：3分之一的一半   (Picture1.Height / 3) / 2)   X轴：3分之一的一半   400  半径刚好
''''   S(2,1)
ElseIf X < Picture1.Width / 3 And Y < (Picture1.Height / 3) * 2 Then     ''' 在第二条线以上
Picture1.Circle ((Picture1.Width / 3) / 2, (Picture1.Height / 3) + (Picture1.Height / 3) / 2), 400, 0      ''''(Picture1.Height / 3) + (Picture1.Height / 3) / 2)   Y轴：3分之一 + 3分之一的一半
''''   S(3,1)
ElseIf X < Picture1.Width / 3 And Y > (Picture1.Height / 3) * 2 Then     ''' 在第二条线以下
Picture1.Circle ((Picture1.Width / 3) / 2, (Picture1.Height / 3) * 2 + (Picture1.Height / 3) / 2), 400, 0            ''''(Picture1.Height / 3) * 2 + (Picture1.Height / 3) / 2)     Y轴：3分之二 + 3分之一的一半
''''   S(1,2)
ElseIf X < (Picture1.Width / 3) * 2 And Y < Picture1.Height / 3 Then
Picture1.Circle ((Picture1.Width / 3) / 2 + (Picture1.Width / 3), (Picture1.Height / 3) / 2), 400, 0
''''   S(2,2)
ElseIf X < (Picture1.Width / 3) * 2 And Y < (Picture1.Width / 3) * 2 Then
Picture1.Circle ((Picture1.Width / 3) / 2 + (Picture1.Width / 3), (Picture1.Height / 3) + (Picture1.Height / 3) / 2), 400, 0
''''   S(3,2)
ElseIf X < (Picture1.Width / 3) * 2 And Y > (Picture1.Width / 3) * 2 Then
Picture1.Circle ((Picture1.Width / 3) / 2 + (Picture1.Width / 3), (Picture1.Height / 3) * 2 + (Picture1.Height / 3) / 2), 400, 0
''''   S(1,3)
ElseIf X > (Picture1.Width / 3) * 2 And Y < Picture1.Height / 3 Then
Picture1.Circle ((Picture1.Width / 3) / 2 + 2 * (Picture1.Width / 3), (Picture1.Height / 3) / 2), 400, 0
''''   S(2,3)
ElseIf X > (Picture1.Width / 3) * 2 And Y < (Picture1.Height / 3) * 2 Then
Picture1.Circle ((Picture1.Width / 3) / 2 + 2 * (Picture1.Width / 3), (Picture1.Height / 3) + (Picture1.Height / 3) / 2), 400, 0
''''   S(3,3)
ElseIf X > (Picture1.Width / 3) * 2 And Y > (Picture1.Height / 3) * 2 Then
Picture1.Circle ((Picture1.Width / 3) / 2 + 2 * (Picture1.Width / 3), (Picture1.Height / 3) * 2 + (Picture1.Height / 3) / 2), 400, 0
End If
End Sub
Private Sub DX(ByVal X As Long, ByVal Y As Long)    ''''  画 X
''''   S(1,1)
If X < Picture1.Width / 3 And Y < Picture1.Height / 3 Then    ''' 在第一条线以上
Picture1.Line (Picture1.Width / 18, Picture1.Height / 18)-((Picture1.Width / 18) * 5, (Picture1.Height / 18) * 5)
Picture1.Line (Picture1.Width / 18, (Picture1.Height / 18) * 5)-((Picture1.Width / 18) * 5, Picture1.Height / 18)
''''   S(2,1)
ElseIf X < Picture1.Width / 3 And Y < (Picture1.Height / 3) * 2 Then     ''' 在第二条线以上
Picture1.Line (Picture1.Width / 18, (Picture1.Height / 18) * 7)-((Picture1.Width / 18) * 5, (Picture1.Height / 18) * 11)
Picture1.Line (Picture1.Width / 18, (Picture1.Height / 18) * 11)-((Picture1.Width / 18) * 5, (Picture1.Height / 18) * 7)     ''''(Picture1.Height / 3) + (Picture1.Height / 3) / 2)   Y轴：3分之一 + 3分之一的一半
''''   S(3,1)
ElseIf X < Picture1.Width / 3 And Y > (Picture1.Height / 3) * 2 Then     ''' 在第二条线以下
Picture1.Line (Picture1.Width / 18, (Picture1.Height / 18) * 13)-((Picture1.Width / 18) * 5, (Picture1.Height / 18) * 17)
Picture1.Line (Picture1.Width / 18, (Picture1.Height / 18) * 17)-((Picture1.Width / 18) * 5, (Picture1.Height / 18) * 13)          ''''(Picture1.Height / 3) * 2 + (Picture1.Height / 3) / 2)     Y轴：3分之二 + 3分之一的一半
''''   S(1,2)
ElseIf X < (Picture1.Width / 3) * 2 And Y < Picture1.Height / 3 Then
Picture1.Line ((Picture1.Width / 18) * 7, Picture1.Height / 18)-((Picture1.Width / 18) * 11, (Picture1.Height / 18) * 5)
Picture1.Line ((Picture1.Width / 18) * 7, (Picture1.Height / 18) * 5)-((Picture1.Width / 18) * 11, Picture1.Height / 18)
''''   S(2,2)
ElseIf X < (Picture1.Width / 3) * 2 And Y < (Picture1.Width / 3) * 2 Then
Picture1.Line ((Picture1.Width / 18) * 7, (Picture1.Height / 18) * 7)-((Picture1.Width / 18) * 11, (Picture1.Height / 18) * 11)
Picture1.Line ((Picture1.Width / 18) * 7, (Picture1.Height / 18) * 11)-((Picture1.Width / 18) * 11, (Picture1.Height / 18) * 7)
''''   S(3,2)
ElseIf X < (Picture1.Width / 3) * 2 And Y > (Picture1.Width / 3) * 2 Then
Picture1.Line ((Picture1.Width / 18) * 7, (Picture1.Height / 18) * 13)-((Picture1.Width / 18) * 11, (Picture1.Height / 18) * 17)
Picture1.Line ((Picture1.Width / 18) * 7, (Picture1.Height / 18) * 17)-((Picture1.Width / 18) * 11, (Picture1.Height / 18) * 13)
''''   S(1,3)
ElseIf X > (Picture1.Width / 3) * 2 And Y < Picture1.Height / 3 Then
Picture1.Line ((Picture1.Width / 18) * 13, Picture1.Height / 18)-((Picture1.Width / 18) * 17, (Picture1.Height / 18) * 5)
Picture1.Line ((Picture1.Width / 18) * 13, (Picture1.Height / 18) * 5)-((Picture1.Width / 18) * 17, Picture1.Height / 18)
''''   S(2,3)
ElseIf X > (Picture1.Width / 3) * 2 And Y < (Picture1.Height / 3) * 2 Then
Picture1.Line ((Picture1.Width / 18) * 13, (Picture1.Height / 18) * 7)-((Picture1.Width / 18) * 17, (Picture1.Height / 18) * 11)
Picture1.Line ((Picture1.Width / 18) * 13, (Picture1.Height / 18) * 11)-((Picture1.Width / 18) * 17, (Picture1.Height / 18) * 7)
''''   S(3,3)
ElseIf X > (Picture1.Width / 3) * 2 And Y > (Picture1.Height / 3) * 2 Then
Picture1.Line ((Picture1.Width / 18) * 13, (Picture1.Height / 18) * 13)-((Picture1.Width / 18) * 17, (Picture1.Height / 18) * 17)
Picture1.Line ((Picture1.Width / 18) * 13, (Picture1.Height / 18) * 17)-((Picture1.Width / 18) * 17, (Picture1.Height / 18) * 13)
End If
End Sub

Private Function SetChess(ByVal X As Long, ByVal Y As Long, ByVal Chess As Long) As Long    ''''意义和 S 数组相同  ((    用于返回该区域是否有棋子   ))

''''   S(1,1)
If X < Picture1.Width / 3 And Y < Picture1.Height / 3 Then    ''' 在第一条线以上

If S(1, 1) = 0 Then
S(1, 1) = Chess
SetChess = 0
Else               '''必须加，若不加的话，则程序会自动重复画图
SetChess = 1
End If

''''   S(2,1)
ElseIf X < Picture1.Width / 3 And Y < (Picture1.Height / 3) * 2 Then     ''' 在第二条线以上
If S(2, 1) = 0 Then
S(2, 1) = Chess
SetChess = 0
Else
SetChess = 1
End If

''''   S(3,1)
ElseIf X < Picture1.Width / 3 And Y > (Picture1.Height / 3) * 2 Then     ''' 在第二条线以下
If S(3, 1) = 0 Then
S(3, 1) = Chess
SetChess = 0
Else
SetChess = 1
End If

''''   S(1,2)
ElseIf X < (Picture1.Width / 3) * 2 And Y < Picture1.Height / 3 Then
If S(1, 2) = 0 Then
S(1, 2) = Chess
SetChess = 0
Else
SetChess = 1
End If

''''   S(2,2)
ElseIf X < (Picture1.Width / 3) * 2 And Y < (Picture1.Width / 3) * 2 Then
If S(2, 2) = 0 Then
S(2, 2) = Chess
SetChess = 0
Else
SetChess = 1
End If

''''   S(3,2)
ElseIf X < (Picture1.Width / 3) * 2 And Y > (Picture1.Width / 3) * 2 Then
If S(3, 2) = 0 Then
S(3, 2) = Chess
SetChess = 0
Else
SetChess = 1
End If

''''   S(1,3)
ElseIf X > (Picture1.Width / 3) * 2 And Y < Picture1.Height / 3 Then
If S(1, 3) = 0 Then
S(1, 3) = Chess
SetChess = 0
Else
SetChess = 1
End If

''''   S(2,3)
ElseIf X > (Picture1.Width / 3) * 2 And Y < (Picture1.Height / 3) * 2 Then
If S(2, 3) = 0 Then
S(2, 3) = Chess
SetChess = 0
Else
SetChess = 1
End If

''''   S(3,3)
ElseIf X > (Picture1.Width / 3) * 2 And Y > (Picture1.Height / 3) * 2 Then
If S(3, 3) = 0 Then
S(3, 3) = Chess
SetChess = 0
Else
SetChess = 1
End If
End If
End Function
''''                  "Can doing
Private Sub Winning(ByVal Chess As Long)
If Chess = 1 Then
 If (MsgBox("圈胜") = vbOK) Then
 Picture1.Cls
 End If
ElseIf Chess = 2 Then
 If (MsgBox("叉胜") = vbOK) Then
 Picture1.Cls
 End If
 End If
 
 Start.Caption = "开始"
 GameStart = False
End Sub

Private Sub ThanksTo_Click()
Dim ThanksFor As String
ThanksFor = "  特别鸣谢:" & vbCrLf & vbCrLf & vbCrLf
ThanksFor = ThanksFor & "  绍款(代码测试)    " & vbCrLf & "  冠霖(代码测试) "
MsgBox ThanksFor
End Sub

Private Sub VSofCom_Click()
VSToCom = True
MsgBox "已经设置为 :人机对战"
Label1.Caption = "人机对战"
End Sub

Private Sub GamerVBGamer_Click()
VSToCom = False
MsgBox "已经设置为 :玩家对战"
Label1.Caption = "玩家对战"
End Sub

'''''' -------------------------Computer  Put  Chess---------------------

Private Sub MainCom()

 If NextStep(2) = True Then
 ChackWinning
 Exit Sub     ''判断自己能否胜利      Exit Sub是为了不让RndPutChess有机会运行
 End If
 If NextStep(1) = True Then Exit Sub          ''判断对方能否胜利

 If ToFocu4 = True Then Exit Sub          ''对于中间
If ToFocu1 = True Then Exit Sub
If ToFocu2 = True Then Exit Sub

RndPutChess
End Sub

Private Function ReturnTheSiteX(ByVal SiteX As Long) As Long
If SiteX = 1 Then
ReturnTheSiteX = (Picture1.Width / 3) / 2     ''''第一条线的一半
ElseIf SiteX = 2 Then
ReturnTheSiteX = Picture1.Width / 2           ''''中间（整体的一半）
ElseIf SiteX = 3 Then
ReturnTheSiteX = (Picture1.Width / 3) / 2 + (Picture1.Width / 3) * 2   ''''最后一条的一半
End If
End Function

Private Function ReturnTheSiteY(ByVal SiteY As Long) As Long
If SiteY = 1 Then
ReturnTheSiteY = Picture1.Height / 6      ''''第一条线的一半
ElseIf SiteY = 2 Then
ReturnTheSiteY = Picture1.Height / 2          ''''中间（整体的一半）
ElseIf SiteY = 3 Then
ReturnTheSiteY = (Picture1.Height / 3) * 2 + Picture1.Height / 6 ''''最后一条的一半
End If
End Function

Private Function NextStep(ByVal Color As Long) As Boolean           ''''注意：1代表绿子，2代表红子

Dim Chess1(3), Chess2(3) As Long
For i = 1 To 3
   For j = 1 To 3
   Chess1(j) = S(i, j)       ''横排
 Chess2(j) = S(j, i)        ''竖排
Next

If Chess1(1) = Chess1(2) And Chess1(1) = Color Then ''''''   Chess1(1) = Chess1(2) = 1  ---不能这样
If SetChess(ReturnTheSiteY(3), ReturnTheSiteX(i), 2) = 0 Then    ''走S(i,3)        ''不知为何要这样 倒转
Drowing ReturnTheSiteY(3), ReturnTheSiteX(i), 2
NextStep = True              '''''''<--------------------------------       必须要这样，不返回值的话，MainCom有时会连续放两颗子的
Exit Function              '''''''<--------------------------------       必须要这样，不退出的话，有时会连续放两颗子的
End If
ElseIf Chess1(1) = Chess1(3) And Chess1(1) = Color Then
If SetChess(ReturnTheSiteY(2), ReturnTheSiteX(i), 2) = 0 Then    ''走S(i,2)
Drowing ReturnTheSiteY(2), ReturnTheSiteX(i), 2
NextStep = True
Exit Function
End If
ElseIf Chess1(2) = Chess1(3) And Chess1(3) = Color Then
If SetChess(ReturnTheSiteY(1), ReturnTheSiteX(i), 2) = 0 Then    ''走S(i,1)
Drowing ReturnTheSiteY(1), ReturnTheSiteX(i), 2
NextStep = True
Exit Function
End If
End If
 ''''    上面是横排攻略
If Chess2(1) = Chess2(2) And Chess2(1) = Color Then
If SetChess(ReturnTheSiteY(i), ReturnTheSiteX(3), 2) = 0 Then    ''走S(3,i)
Drowing ReturnTheSiteY(i), ReturnTheSiteX(3), 2
NextStep = True
Exit Function
End If
ElseIf Chess2(1) = Chess2(3) And Chess2(1) = Color Then
If SetChess(ReturnTheSiteY(i), ReturnTheSiteX(2), 2) = 0 Then    ''走S(2,i)
Drowing ReturnTheSiteY(i), ReturnTheSiteX(2), 2
NextStep = True
Exit Function
End If
ElseIf Chess2(2) = Chess2(3) And Chess2(3) = Color Then
If SetChess(ReturnTheSiteY(i), ReturnTheSiteX(1), 2) = 0 Then    ''走S(1,i)
Drowing ReturnTheSiteY(i), ReturnTheSiteX(1), 2
NextStep = True
Exit Function
End If
End If
Next
 ''''    上面是竖排攻略
 ''''    下面是对角线攻略
 For j = 1 To 3
 Chess1(j) = S(j, j)        ''""" 左上角与右下角 的对角线
 Chess2(j) = S(4 - j, j)     ''""" 左下角与右上角 的对角线
 Next
  ''""" 左上角与右下角 的对角线攻略
If Chess1(1) = Chess1(2) And Chess1(1) = Color Then                   ''S(1,1)  S(2,2)
If SetChess(ReturnTheSiteY(3), ReturnTheSiteX(3), 2) = 0 Then    ''走S(3,3)
Drowing ReturnTheSiteY(3), ReturnTheSiteX(3), 2
NextStep = True
Exit Function
End If
ElseIf Chess1(1) = Chess1(3) And Chess1(1) = Color Then                  ''S(1,1)  S(3,3)
If SetChess(ReturnTheSiteY(2), ReturnTheSiteX(2), 2) = 0 Then    ''走S(2,2)
Drowing ReturnTheSiteY(2), ReturnTheSiteX(2), 2
NextStep = True
Exit Function
End If
ElseIf Chess1(2) = Chess1(3) And Chess1(3) = Color Then                 ''S(2,2)  S(3,3)
If SetChess(ReturnTheSiteY(1), ReturnTheSiteX(1), 2) = 0 Then    ''走S(1,1)
Drowing ReturnTheSiteY(1), ReturnTheSiteX(1), 2
NextStep = True
Exit Function
End If
End If

  ''""" 左下角与右上角 的对角线攻略
If Chess2(1) = Chess2(2) And Chess2(1) = Color Then                   ''S(3,1)  S(2,2)
If SetChess(ReturnTheSiteY(3), ReturnTheSiteX(1), 2) = 0 Then    ''走S(1,3)
Drowing ReturnTheSiteY(3), ReturnTheSiteX(1), 2
NextStep = True
Exit Function
End If
ElseIf Chess2(1) = Chess2(3) And Chess2(1) = Color Then                  ''S(3,1)  S(1,3)
If SetChess(ReturnTheSiteY(2), ReturnTheSiteX(2), 2) = 0 Then    ''走S(2,2)
Drowing ReturnTheSiteY(2), ReturnTheSiteX(2), 2
NextStep = True
Exit Function
End If
ElseIf Chess2(2) = Chess2(3) And Chess2(3) = Color Then                   ''S(2,2)  S(1,3)
If SetChess(ReturnTheSiteY(1), ReturnTheSiteX(3), 2) = 0 Then    ''走S(1,3)
Drowing ReturnTheSiteY(1), ReturnTheSiteX(3), 2
NextStep = True
Exit Function
End If
End If

NextStep = False
End Function

Private Sub RndPutChess()
Dim Rnd1, Rnd2 As Long
Rnd1 = Int(Rnd * 3 + 1)
Rnd2 = Int(Rnd * 3 + 1)
If SetChess(ReturnTheSiteY(Rnd1), ReturnTheSiteX(Rnd2), 2) = 0 Then
Drowing ReturnTheSiteY(Rnd1), ReturnTheSiteX(Rnd2), 2
Else
RndPutChess
End If
End Sub

Private Function ToFocu4() As Boolean            ''''''  对于四角 有这个函数
If S(1, 1) = 1 Or S(1, 3) = 1 Or S(3, 1) = 1 Or S(3, 3) = 1 Then
If SetChess(ReturnTheSiteY(2), ReturnTheSiteX(2), 2) = 0 Then
Drowing ReturnTheSiteY(2), ReturnTheSiteX(2), 2
ToFocu4 = True
Exit Function         '''' 成功就要退出，别给机会运行到下面
End If
End If

ToFocu4 = False
End Function

Private Function ToFocu2() As Boolean            ''''''  对于四角加中间 有这个函数
If S(2, 1) <> 0 And S(1, 2) <> 0 And S(3, 2) <> 0 And S(2, 3) <> 0 Then
ToFocu2 = False
Exit Function
End If

If S(1, 1) = 1 And S(3, 3) = 1 And S(2, 2) = 2 Or S(3, 1) = 1 And S(1, 3) = 1 And S(2, 2) Then
Dim nRnd As Long
nRnd = Int(Rnd(1) * 4 + 1)
If nRnd = 1 Then                       ''''  放置在 S(2,1) 上
If SetChess(ReturnTheSiteY(1), ReturnTheSiteX(2), 2) = 0 Then
Drowing ReturnTheSiteY(1), ReturnTheSiteX(2), 2
ToFocu2 = True
Exit Function
Else
ToFocu2
End If

ElseIf nRnd = 2 Then                       ''''  放置在 S(1,2) 上
If SetChess(ReturnTheSiteY(2), ReturnTheSiteX(1), 2) = 0 Then
Drowing ReturnTheSiteY(2), ReturnTheSiteX(1), 2
ToFocu2 = True
Exit Function
Else
ToFocu2
End If

ElseIf nRnd = 3 Then                       ''''  放置在 S(3,2) 上
If SetChess(ReturnTheSiteY(2), ReturnTheSiteX(3), 2) = 0 Then
Drowing ReturnTheSiteY(2), ReturnTheSiteX(3), 2
ToFocu2 = True
Exit Function
Else
ToFocu2
End If

ElseIf nRnd = 4 Then                       ''''  放置在 S(2,3) 上
If SetChess(ReturnTheSiteY(3), ReturnTheSiteX(2), 2) = 0 Then
Drowing ReturnTheSiteY(3), ReturnTheSiteX(2), 2
ToFocu2 = True
Exit Function
Else
ToFocu2
End If
End If
End If

ToFocu2 = False
End Function

Private Function ToFocu1() As Boolean            ''''''  对于中间 有这个函数
If S(1, 1) <> 0 And S(1, 3) <> 0 And S(3, 1) <> 0 And S(3, 3) <> 0 Then
ToFocu1 = False
Exit Function
End If

If S(2, 2) = 1 Then

Dim nRnd As Long
nRnd = Int(Rnd(1) * 4 + 1)
If nRnd = 1 Then                       ''''  放置在 S(1,1) 上
If SetChess(ReturnTheSiteY(1), ReturnTheSiteX(1), 2) = 0 Then
Drowing ReturnTheSiteY(1), ReturnTheSiteX(1), 2
ToFocu1 = True
Exit Function
Else
ToFocu1
End If

ElseIf nRnd = 2 Then                       ''''  放置在 S(1,3) 上
If SetChess(ReturnTheSiteY(3), ReturnTheSiteX(1), 2) = 0 Then
Drowing ReturnTheSiteY(3), ReturnTheSiteX(1), 2
ToFocu1 = True
Exit Function
Else
ToFocu1
End If

ElseIf nRnd = 3 Then                       ''''  放置在 S(3,1) 上
If SetChess(ReturnTheSiteY(1), ReturnTheSiteX(3), 2) = 0 Then
Drowing ReturnTheSiteY(1), ReturnTheSiteX(3), 2
ToFocu1 = True
Exit Function
Else
ToFocu1
End If

ElseIf nRnd = 4 Then                       ''''  放置在 S(3,3) 上
If SetChess(ReturnTheSiteY(3), ReturnTheSiteX(3), 2) = 0 Then
Drowing ReturnTheSiteY(3), ReturnTheSiteX(3), 2
ToFocu1 = True
Exit Function
Else
ToFocu1
End If

End If
End If

ToFocu1 = False
End Function
