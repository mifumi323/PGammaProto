VERSION 5.00
Begin VB.Form FormMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '固定(実線)
   Caption         =   "パネルγproto"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9585
   Icon            =   "FormMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   500
   ScaleMode       =   3  'ﾋﾟｸｾﾙ
   ScaleWidth      =   639
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton Command1 
      Caption         =   "消す"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   615
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   3840
      Left            =   3240
      Picture         =   "FormMain.frx":0442
      ScaleHeight     =   256
      ScaleMode       =   3  'ﾋﾟｸｾﾙ
      ScaleWidth      =   256
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.Label lScore 
      Caption         =   "得点"
      Height          =   945
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   960
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Sub BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long)

Const Tate = 28     '縦に並べる数
Const Yoko = 32     '横に並べる数
Const Iro = 3       '色数

Dim Data(Yoko - 1, Tate - 1) As Integer
Dim Score As Long
Dim XX As Integer, YY As Integer

Dim Flag As Boolean     '消え中
Dim Mouse As Boolean

Private Sub Command1_Click()
    If Not Flag Then
        Dim Ochita As Boolean, Kieta As Boolean
        Dim I As Integer, J As Integer
        Dim Rensa As Integer
        Dim Plus As Long
        Rensa = 0
        Flag = True
        Do
            ' 空中ブロックを落とす
            Do
                Ochita = False
                For I = 0 To Yoko - 1
                    For J = Tate - 1 To 1 Step -1   ' 下から検証
                        If Data(I, J) = 0 And Data(I, J - 1) <> 0 Then
                            Data(I, J) = Data(I, J - 1)
                            Data(I, J - 1) = 0
                            Ochita = True
                        End If
                    Next
                Next
                If Ochita = False Then Exit Do
                Wait
            Loop
            Wait
            Kieta = False
            For I = 0 To Yoko - 2
                For J = 0 To Tate - 2
                    If Data(I, J) <> 0 And Data(I, J) Mod 100 = Data(I + 1, J) Mod 100 And Data(I, J) Mod 100 = Data(I, J + 1) Mod 100 And Data(I, J) Mod 100 = Data(I + 1, J + 1) Mod 100 Then
                        Data(I, J) = Data(I, J) Mod 100 + 100
                        Data(I + 1, J) = Data(I + 1, J) Mod 100 + 100
                        Data(I, J + 1) = Data(I, J + 1) Mod 100 + 100
                        Data(I + 1, J + 1) = Data(I + 1, J + 1) Mod 100 + 100
                        Plus = Plus + 4 * (Rensa + 1)   ' 消し得点 = 基本得点 * 連鎖数
                        Kieta = True
                    End If
                Next
            Next
            If Kieta = False Then Exit Do
            Rensa = Rensa + 1
            Plus = Plus + Rensa     ' コンボボーナス
            For I = 0 To Yoko - 1
                For J = 0 To Tate - 1
                    If Data(I, J) >= 100 Then Data(I, J) = 99 + Rensa
                Next
            Next
            Wait
            For I = 0 To Yoko - 1
                For J = 0 To Tate - 1
                    If Data(I, J) >= 100 Then Data(I, J) = 0
                Next
            Next
            lScore.Caption = "得点：" & vbNewLine & Score & vbNewLine & " +" & Plus
            Wait
        Loop
        If Plus Then
            Score = Score + Plus
            lScore.Caption = "得点：" & vbNewLine & Score
        End If
        Flag = False
    End If
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        Dim FileName As String
        Dim I As Long
        Do
            FileName = "pgamma" & Format(I, "000") & ".bmp"
            If Dir(FileName) = "" Then Exit Do
            I = I + 1
            DoEvents
        Loop
        SavePicture Image, FileName
        Caption = "Save:" & FileName
    End If
End Sub

Private Sub Form_Load()
    Width = Width + (640 - ScaleWidth) * Screen.TwipsPerPixelX
    Height = Height + (480 - ScaleHeight) * Screen.TwipsPerPixelY
    OnDraw
End Sub

Public Sub OnDraw()
    Cls
    Dim I As Integer, J As Integer, X As Long, Y As Long
    For I = 0 To Yoko - 1
        For J = 0 To Tate - 1
            X = (ScaleWidth - Yoko * 15) / 2 + I * 15
            Y = J * 15
            BitBlt hDC, X, Y, 15, 15, pic.hDC, (Data(I, J) Mod 100) * 15, (Data(I, J) \ 100) * 15, &HCC0020
        Next
    Next
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Flag Then Exit Sub
    Dim I As Integer, J As Integer
    XX = (X - (ScaleWidth - Yoko * 15) / 2) \ 15
    YY = Y \ 15
    If Not (0 <= XX And XX < Yoko And 0 <= YY And YY < Tate) Then Exit Sub
    If Button = vbLeftButton Then Data(XX, YY) = Data(XX, YY) + 1
    If Button = vbRightButton Then Data(XX, YY) = Data(XX, YY) - 1
    If Data(XX, YY) < 1 Then Data(XX, YY) = Iro
    If Data(XX, YY) > Iro Then Data(XX, YY) = 1
    OnDraw
    Mouse = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Flag Or Not Mouse Then Exit Sub
    Dim I As Integer, J As Integer
    I = (X - (ScaleWidth - Yoko * 15) / 2) \ 15
    J = Y \ 15
    If I <> XX Or J <> YY Then
        XX = I
        YY = J
        If Not (0 <= XX And XX < Yoko And 0 <= YY And YY < Tate) Then Exit Sub
        If Button = vbLeftButton Then Data(XX, YY) = Data(XX, YY) + 1
        If Button = vbRightButton Then Data(XX, YY) = Data(XX, YY) - 1
        If Data(XX, YY) < 1 Then Data(XX, YY) = Iro
        If Data(XX, YY) > Iro Then Data(XX, YY) = 1
        OnDraw
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Mouse = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub

Public Sub Wait()
    Dim a As Integer
    OnDraw
    For a = 0 To 19
        Sleep 10
        DoEvents
    Next
End Sub
