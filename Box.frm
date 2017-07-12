VERSION 5.00
Begin VB.Form Box 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   6045
   ClientLeft      =   6960
   ClientTop       =   4425
   ClientWidth     =   15255
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   14.25
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   15255
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808080&
      Caption         =   "pause"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13680
      TabIndex        =   2
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Jump!"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12120
      TabIndex        =   0
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   240
      Top             =   240
   End
   Begin VB.Label MMM 
      Caption         =   "maxscore"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10920
      TabIndex        =   3
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Shape PY 
      FillStyle       =   7  'Diagonal Cross
      Height          =   495
      Left            =   1300
      Top             =   4080
      Width           =   495
   End
   Begin VB.Shape Shape 
      FillStyle       =   4  'Upward Diagonal
      Height          =   1455
      Index           =   1
      Left            =   10680
      Top             =   3120
      Width           =   405
   End
   Begin VB.Shape Shape 
      FillStyle       =   5  'Downward Diagonal
      Height          =   2295
      Index           =   0
      Left            =   5500
      Top             =   2280
      Width           =   400
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   0
      Top             =   4560
      Width           =   15255
   End
   Begin VB.Label TTT 
      Caption         =   "score"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10920
      TabIndex        =   1
      Top             =   480
      Width           =   3375
   End
End
Attribute VB_Name = "Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim jumping As Integer, v As Integer
Dim speed As Integer, time As Integer, mxscore As Double, score As Double
Dim r, g, b
Dim f As Boolean

Private Sub CrtTXT(Path As String) '创建文件
    Open Path For Output As #1
    Close #1
End Sub

Private Function GetTXT(Path As String) '读取文本
    Open Path For Input As #1
     GetTXT = StrConv(InputB(LOF(1), 1), vbUnicode)
    Close #1
End Function

Private Sub SetTXT(Strl As String, Path As String) '写入文本
    Open Path For Output As #1
    Print #1, Strl;
    Close #1
End Sub

Public Function getanum(X As Integer) As Integer
    getanum = Int(Rnd * X)
End Function


Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 Then
        If jumping = 0 Then jumping = 1: v = -170 * 3
    End If
    If KeyCode = 80 Then
        If f Then
            f = False: Timer1.Interval = 0
            Else: f = True: Timer1.Interval = 50
        End If
    End If
End Sub

Private Sub Command1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 Then jumping = -1
End Sub


Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If jumping = 0 Then jumping = 1: v = -170 * 3
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    jumping = -1
End Sub

Private Sub Command2_Click()
    If f Then
        f = False: Timer1.Interval = 0
    Else: f = True: Timer1.Interval = 50
    End If
End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 32 Then
        If jumping = 0 Then jumping = 1: v = -170 * 3
    End If
    If KeyCode = 80 Then
        If f Then
            f = False: Timer1.Interval = 0
            Else: f = True: Timer1.Interval = 50
        End If
    End If
End Sub

Private Sub Command2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 Then jumping = -1
End Sub


Public Sub Main()
    mxscore = Val(GetTXT("C:\Users\yanghong\Documents\box\mxscore"))
    MMM.Caption = "maxscore" & " : " & Str(mxscore)
    f = True
    r = 250: g = 250: b = 250
    Box.BackColor = RGB(r, g, b)
    Randomize
    score = 0
    
    PY.Left = 1300
    PY.Top = 4080
    v = 0
    Shape(0).Left = 16000
    Shape(1).Left = 16000 + 8000
    Shape(0).Height = getanum(1500) + 500
    Shape(1).Height = getanum(1500) + 500
    Shape(0).Top = Shape1.Top - Shape(0).Height
    Shape(1).Top = Shape1.Top - Shape(1).Height
    speed = 70 * 5
    
End Sub

Public Sub chk()
    
    ll = PY.Left
    rr = PY.Left + PY.Width
    uu = PY.Top
    dd = PY.Top + PY.Height
    
    For i = 0 To 1
        
        Dim l As Integer, r As Integer, u As Integer, d As Integer
        l = Shape(i).Left
        r = Shape(i).Left + Shape(i).Width
        u = Shape(i).Top
        d = Shape(i).Top + Shape(i).Height
        
        Dim chk As Boolean
        chk = True
        If dd >= u Then
            If l <= ll And ll <= r Then chk = False
            If l <= rr And rr <= r Then chk = False
        End If
        If Not chk Then
            MsgBox "爆炸:" + Chr(10) & Chr(13) + "score : " + CStr(score), 0, "Boom"
            Call SetTXT(Str(mxscore), "C:\Users\yanghong\Documents\box\mxscore")
            Call Main
        End If
    Next i
    
End Sub




Private Sub score_Click()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Form_Load()
    'CrtTXT ("C:\Users\yanghong\Documents\box\mxscore")
    Call Main
End Sub

Private Sub Timer1_Timer()
    Dim i As Integer, k As Integer
    
    For i = 0 To 1
        Shape(i).Left = Shape(i).Left - speed
        
        If Shape(i).Left < -400 Then
            Shape(i).Left = 15600
            Shape(i).Height = getanum(1500) + 500
            Shape(i).Top = Shape1.Top - Shape(i).Height
        End If
        
    Next i
    
    
    If jumping = 1 Then v = v + 5 * 9
    If jumping = -1 Then v = v + 15 * 9
    PY.Top = PY.Top + v
    If v > 0 Then jumping = -1
    If PY.Top + PY.Height > Shape1.Top Then
        jumping = 0
        v = 0
        PY.Top = Shape1.Top - PY.Height
    End If
    
    Call chk
    
    score = score + speed / 500
    score = Int(score * 10000) / 10000
    time = time + 1
    If time Mod 20 = 0 And g >= 0 Then
        g = g - 1
        b = b - 1
        Box.BackColor = RGB(r, g, b)
    End If
    If time = 100 And g >= 0 Then
        speed = speed + 10: time = 0
    End If
    
    TTT.Caption = "score   " & " : " & Str(score)
    If score > mxscore Then
        mxscore = score
    End If
    MMM.Caption = "maxscore" & " : " & Str(mxscore)
    
End Sub

