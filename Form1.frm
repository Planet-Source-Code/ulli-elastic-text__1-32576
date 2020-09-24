VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "ElasticText"
   ClientHeight    =   3105
   ClientLeft      =   5115
   ClientTop       =   2745
   ClientWidth     =   8175
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   207
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   2550
      Width           =   1275
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z '32 bit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds)
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC, lpRect As RECT, ByVal hBrush) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject) As Long
Private Declare Function GetTextCharacterExtra Lib "gdi32" (ByVal hDC) As Long
Private Declare Function SetTextCharacterExtra Lib "gdi32" (ByVal hDC, ByVal nCharExtra) As Long

Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Private Done                As Boolean
Private TextRec             As RECT
Private StandardSpacing     As Long
Private Canvas              As Object

Private Sub cmdOK_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    Show
    DoEvents
    Do
        BackColor = RGB(Rnd * 127 + 128, Rnd * 127 + 128, Rnd * 127 + 128)
        FontBold = True
        If Not Done Then
            FontSize = 14
            BounceText "Planet Source Code Forever", 16, 16, 40, 4, , &HFFFFFF - BackColor
        End If
        If Not Done Then
            FontSize = 12
            BounceText "Hi folks! Whishing you a successful day...", 36, 64, 30, 3
        End If
        If Not Done Then
            FontSize = 10
            BounceText "...and don't forget to vote EXCELLENT :-)", 56, 110, 60, 5, 20, vbRed
        End If
        If Not Done Then
            Sleep 1111
        End If
    Loop Until Done

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Done = True

End Sub

Private Sub BounceText(ByVal Text As String, _
                       ByVal x, _
                       ByVal y, _
                       Optional ByVal Speed = 60, _
                       Optional ByVal Loops = 1, _
                       Optional ByVal StartSpacing = 0, _
                       Optional ByVal TextColor As OLE_COLOR = vbButtonText)

  Dim hBrush        As Long 'brush handle
  Dim MinSpacing    As Long
  Dim MaxSpacing    As Long
  Dim CurrSpacing   As Long
  Dim Shrink        As Boolean

    ScaleMode = vbPixels
    StandardSpacing = GetTextCharacterExtra(hDC)
    MinSpacing = -9
    MaxSpacing = FontSize * 3

    With TextRec
        .Left = x
        .Top = y
        .Right = ScaleWidth
        .Bottom = y + TextHeight("A")
    End With 'TEXTREC

    hBrush = CreateSolidBrush(BackColor)
    ForeColor = TextColor
    CurrSpacing = StartSpacing

    Done = False
    Shrink = (CurrSpacing > StandardSpacing)
    Do
        Select Case Shrink
          Case True
            CurrSpacing = CurrSpacing - 1
            If CurrSpacing < MinSpacing Then
                Shrink = False
                Loops = Loops - 1
            End If
          Case False
            CurrSpacing = CurrSpacing + 1
            If CurrSpacing > MaxSpacing Then
                Shrink = True
                Loops = Loops - 1
            End If
        End Select

        SetTextCharacterExtra hDC, CurrSpacing
        CurrentX = x
        CurrentY = y
        FillRect hDC, TextRec, hBrush 'erase previous text
        Print Text
        DoEvents
        Sleep Speed
    Loop Until (CurrSpacing = StandardSpacing And Loops = 0) Or Done
    DeleteObject hBrush 'kill Brush

End Sub

':) Ulli's VB Code Formatter V2.10.8 (11.03.2002 10:26:53) 20 + 100 = 120 Lines
