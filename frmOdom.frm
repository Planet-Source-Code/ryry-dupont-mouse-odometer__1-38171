VERSION 5.00
Begin VB.Form frmOdom 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2220
   LinkTopic       =   "Form1"
   Picture         =   "frmOdom.frx":0000
   ScaleHeight     =   147
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   148
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrMove 
      Left            =   240
      Top             =   0
   End
   Begin VB.PictureBox picNumW 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2805
      Left            =   720
      Picture         =   "frmOdom.frx":FF36
      ScaleHeight     =   187
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picNumB 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2805
      Left            =   480
      Picture         =   "frmOdom.frx":11CB0
      ScaleHeight     =   187
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picNumM 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   480
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   91
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picCM 
      Height          =   495
      Left            =   120
      ScaleHeight     =   0.767
      ScaleMode       =   7  'Centimeter
      ScaleWidth      =   0.767
      TabIndex        =   1
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picP 
      Height          =   495
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblTip 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Line linB 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      X1              =   0
      X2              =   73
      Y1              =   0
      Y2              =   74
   End
   Begin VB.Line linA 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      X1              =   0
      X2              =   73
      Y1              =   0
      Y2              =   74
   End
   Begin VB.Label lblSpeed 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0.00 cm/s"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "frmOdom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINT_TYPE) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hdcDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal dwRop As Long) As Long
Private Type POINT_TYPE
    X As Long
    Y As Long
End Type

Dim CMperP, KMperP, OLD As POINT_TYPE, CooRD As POINT_TYPE
Dim LASTtime As Long, THIStime As Long
Dim TOT, CURR, theSPEED
Dim theDATE, LPx, LPy
Const PI = 3.14159
Private Sub Form_DblClick()
    Unload frmOdom
End Sub
Private Sub Form_Load()
    Call LoadProc
    Call MakeRound(frmOdom, 800)
    'get CM per pixel using 2 scaled pictureboxes
    CMperP = picCM.ScaleHeight / picP.ScaleHeight
    KMperP = CMperP / 100000
    GetCursorPos CooRD
    THIStime = GetTickCount
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        'set up to move the form around
        LPx = X
        LPy = Y
        tmrMove.Interval = 1
    End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'stop moving form
    If tmrMove.Interval = 1 Then tmrMove.Interval = 0
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'save current distance
    Open App.Path & "\settings.dat" For Output As #1
        Print #1, TOT
        Print #1, theDATE
    Close #1
End Sub
Private Sub Timer1_Timer()
    'the main operations
    OLD = CooRD
    GetCursorPos CooRD
    'calculate the distance between the last recorded point and the newest recorded point
    CURR = Sqr((OLD.X - CooRD.X) ^ 2 + (OLD.Y - CooRD.Y) ^ 2) * CMperP
    LASTtime = THIStime
    THIStime = GetTickCount
    'calculate cursor speed in cm/s
    theSPEED = CURR / ((THIStime - LASTtime) / 1000)
    'add to the total distance(recorded in KM)
    TOT = TOT + CURR / 100000
    lblSpeed.Caption = Round(theSPEED, 2) & " cm/s"
    'move odometer pointer
    linA.X1 = Cos(((theSPEED * 2 + 130) * PI) / 180) * 72.5 + linA.X2
    linA.Y1 = Sin(((theSPEED * 2 + 130) * PI) / 180) * 72.5 + linA.Y2
    linB.X1 = Cos(((theSPEED * 2 + 310) * PI) / 180) * 10 + linB.X2
    linB.Y1 = Sin(((theSPEED * 2 + 310) * PI) / 180) * 10 + linB.Y2
    'update the milage
    If theSPEED > 0 Then DoMilage (TOT * 1000)
End Sub
Private Sub tmrMove_Timer()
    'move the form around the screen
    Dim CURRloc As POINT_TYPE
    GetCursorPos CURRloc
    frmOdom.Left = (CURRloc.X - LPx) * Screen.TwipsPerPixelX
    frmOdom.Top = (CURRloc.Y - LPy) * Screen.TwipsPerPixelY
End Sub
Private Sub DoMilage(AMT)
    'this is a little odometer code i threw together.
    'all of the nested if/then statements are to make
    'it have a realistic rollover effect...move the
    'mouse slow when the last digit is a 9 to see
    'what i mean
    AMT = Format(AMT, "000000.00")
    If Right(AMT, 2) > 90 Then
        BitBlt picNumM.hDC, picNumM.Width - 26, 0, 13, 17, picNumB.hDC, 0, Mid(AMT, 6, 1) * 17 + Right(AMT, 1) * 1.7, vbSrcCopy
        If Mid(AMT, 6, 1) = 9 Then
            BitBlt picNumM.hDC, picNumM.Width - 39, 0, 13, 17, picNumB.hDC, 0, Mid(AMT, 5, 1) * 17 + Right(AMT, 1) * 1.7, vbSrcCopy
            If Mid(AMT, 5, 1) = 9 Then
                BitBlt picNumM.hDC, picNumM.Width - 52, 0, 13, 17, picNumB.hDC, 0, Mid(AMT, 4, 1) * 17 + Right(AMT, 1) * 1.7, vbSrcCopy
                If Mid(AMT, 4, 1) = 9 Then
                    BitBlt picNumM.hDC, picNumM.Width - 65, 0, 13, 17, picNumB.hDC, 0, Mid(AMT, 3, 1) * 17 + Right(AMT, 1) * 1.7, vbSrcCopy
                    If Mid(AMT, 3, 1) = 9 Then
                        BitBlt picNumM.hDC, picNumM.Width - 78, 0, 13, 17, picNumB.hDC, 0, Mid(AMT, 2, 1) * 17 + Right(AMT, 1) * 1.7, vbSrcCopy
                        If Mid(AMT, 2, 1) = 9 Then
                            BitBlt picNumM.hDC, picNumM.Width - 91, 0, 13, 17, picNumB.hDC, 0, Left(AMT, 1) * 17 + Right(AMT, 1) * 1.7, vbSrcCopy
                        Else
                            BitBlt picNumM.hDC, picNumM.Width - 91, 0, 13, 17, picNumB.hDC, 0, Left(AMT, 1) * 17, vbSrcCopy
                        End If
                    Else
                        BitBlt picNumM.hDC, picNumM.Width - 78, 0, 13, 17, picNumB.hDC, 0, Mid(AMT, 2, 1) * 17, vbSrcCopy
                        BitBlt picNumM.hDC, picNumM.Width - 91, 0, 13, 17, picNumB.hDC, 0, Left(AMT, 1) * 17, vbSrcCopy
                    End If
                Else
                    BitBlt picNumM.hDC, picNumM.Width - 65, 0, 13, 17, picNumB.hDC, 0, Mid(AMT, 3, 1) * 17, vbSrcCopy
                    BitBlt picNumM.hDC, picNumM.Width - 78, 0, 13, 17, picNumB.hDC, 0, Mid(AMT, 2, 1) * 17, vbSrcCopy
                    BitBlt picNumM.hDC, picNumM.Width - 91, 0, 13, 17, picNumB.hDC, 0, Left(AMT, 1) * 17, vbSrcCopy
                End If
            Else
                BitBlt picNumM.hDC, picNumM.Width - 52, 0, 13, 17, picNumB.hDC, 0, Mid(AMT, 4, 1) * 17, vbSrcCopy
                BitBlt picNumM.hDC, picNumM.Width - 65, 0, 13, 17, picNumB.hDC, 0, Mid(AMT, 3, 1) * 17, vbSrcCopy
                BitBlt picNumM.hDC, picNumM.Width - 78, 0, 13, 17, picNumB.hDC, 0, Mid(AMT, 2, 1) * 17, vbSrcCopy
                BitBlt picNumM.hDC, picNumM.Width - 91, 0, 13, 17, picNumB.hDC, 0, Left(AMT, 1) * 17, vbSrcCopy
            End If
        Else
            BitBlt picNumM.hDC, picNumM.Width - 39, 0, 13, 17, picNumB.hDC, 0, Mid(AMT, 5, 1) * 17, vbSrcCopy
            BitBlt picNumM.hDC, picNumM.Width - 52, 0, 13, 17, picNumB.hDC, 0, Mid(AMT, 4, 1) * 17, vbSrcCopy
            BitBlt picNumM.hDC, picNumM.Width - 65, 0, 13, 17, picNumB.hDC, 0, Mid(AMT, 3, 1) * 17, vbSrcCopy
            BitBlt picNumM.hDC, picNumM.Width - 78, 0, 13, 17, picNumB.hDC, 0, Mid(AMT, 2, 1) * 17, vbSrcCopy
            BitBlt picNumM.hDC, picNumM.Width - 91, 0, 13, 17, picNumB.hDC, 0, Left(AMT, 1) * 17, vbSrcCopy
        End If
        BitBlt picNumM.hDC, picNumM.Width - 13, 0, 13, 17, picNumW.hDC, 0, Right(AMT, 2) * 1.7, vbSrcCopy
    Else
        'each number is 17 high and 13 wide...copy from the
        'numbered picture to the main milage listing number
        'by number
        BitBlt picNumM.hDC, picNumM.Width - 13, 0, 13, 17, picNumW.hDC, 0, Right(AMT, 2) * 1.7, vbSrcCopy
        BitBlt picNumM.hDC, picNumM.Width - 26, 0, 13, 17, picNumB.hDC, 0, Mid(AMT, 6, 1) * 17, vbSrcCopy
        BitBlt picNumM.hDC, picNumM.Width - 39, 0, 13, 17, picNumB.hDC, 0, Mid(AMT, 5, 1) * 17, vbSrcCopy
        BitBlt picNumM.hDC, picNumM.Width - 52, 0, 13, 17, picNumB.hDC, 0, Mid(AMT, 4, 1) * 17, vbSrcCopy
        BitBlt picNumM.hDC, picNumM.Width - 65, 0, 13, 17, picNumB.hDC, 0, Mid(AMT, 3, 1) * 17, vbSrcCopy
        BitBlt picNumM.hDC, picNumM.Width - 78, 0, 13, 17, picNumB.hDC, 0, Mid(AMT, 2, 1) * 17, vbSrcCopy
        BitBlt picNumM.hDC, picNumM.Width - 91, 0, 13, 17, picNumB.hDC, 0, Left(AMT, 1) * 17, vbSrcCopy
    End If
    picNumM.Refresh
    BitBlt frmOdom.hDC, picNumM.Left, picNumM.Top, picNumM.Width, picNumM.Height, picNumM.hDC, 0, 0, vbSrcCopy
    frmOdom.Refresh
    lblTip.ToolTipText = Round(TOT * 1000, 2) & " meters since " & theDATE
End Sub
Private Sub LoadProc()
    'if file doesnt exist, create it
    If Dir$(App.Path & "\settings.dat") = "" Then
        Open App.Path & "\settings.dat" For Output As #1
            Print #1, 0
            Print #1, Date
        Close #1
    End If
    'load current distance
    Open App.Path & "\settings.dat" For Input As #1
        Input #1, TOT
        Line Input #1, theDATE
    Close #1
    Call DoMilage(TOT * 1000)
End Sub
