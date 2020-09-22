VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Brownian line"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   4980
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4440
      Top             =   4920
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   4800
      Left            =   120
      ScaleHeight     =   4740
      ScaleMode       =   0  'User
      ScaleWidth      =   5664.986
      TabIndex        =   0
      Top             =   120
      Width           =   4800
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   5040
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const Pi = 3.14159263

Dim N As Single         'starting point on X-axis

Dim Xs As Single:       Dim Ys As Single    'picturebox start point
Dim X0 As Single:       Dim Y0 As Single    'start point of line
Dim X1 As Single:       Dim Y1 As Single    'end point of the line
Dim X2 As Single:       Dim Y2 As Single    'end point of the line

Dim W As Integer        'magnitude of randomnes
Dim Dimension As Byte   'selection the randomness in movment i 1D or 2D

Private Sub Form_Load()
    Timer1.Enabled = True:      Timer1.Interval = 1
    
    N = 0
    X0 = 0:                     Y0 = 0
    W = 2000:
    Dimension = InputBox("Select " & vbCrLf & "[1] 1D randomnes or " & vbCrLf & "[2] 2D randomnes", 100, 1, 300)
End Sub

Private Sub Timer1_Timer()
    Call Draw
    N = N + 1:     Label1.Caption = N
    'stop condition expresed in number of 0 to 180° and 180 to 360° intervals
    If N > 4500 Then Timer1.Enabled = False
End Sub

Private Function Draw()
    Select Case Dimension
        Case 1
            Xs = 200:                   Ys = 1500   'positioning on picturebox
            X1 = N
            Y1 = Ys + W * (0.2 - Rnd(2))
        Case 2
            Xs = 800:                   Ys = 1500   'positioning on picturebox
            X1 = N + W / 2 * (0.2 - Rnd(2))
            Y1 = Ys + W * 3 / 4 * (-0.1 - Rnd(2))
    End Select
        
        If (N Mod 40) = 0 Then
            'draw every 40 point. but if you zoom in (decreas step form 40 to 10 or 1) self similarity
            'will apperar - see avi animation. Zoom is not included because it is algorithm that this
            'presentation is focused on. Extra code would only blur the intention of code.
            Picture1.Line (Xs + X2, Ys + Y2)-(Xs + X1, Ys + Y1)
            X2 = X1:                    Y2 = Y1
            X0 = X1:                    Y0 = Y1
        Else
            X0 = X1:                    Y0 = Y1
        End If
        Picture1.Refresh
        'Debug.Print Format(N, "00.0") & "   " & Format(X0, "#.00") & "  " & Format(Y0, "#.00") _
                      & "   " & Format(X1, "#.00") & "    " & Format(Y1, "#.00")
        
        
End Function

Private Sub Picture1_Click()
    Unload Me
End Sub

Private Sub Form_Click()
    Unload Me
End Sub
