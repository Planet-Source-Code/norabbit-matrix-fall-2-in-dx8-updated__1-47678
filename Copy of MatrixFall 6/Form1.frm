VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   7995
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Private Sub Form_Click()
    '
    If PauseSz = False Then
        '
        PauseSz = True
        '
    Else
        '
        'on remet les compteurs à zéro
        FPS_NbrImg = 0
        lFpsTmp = GetTickCount - 1
        '
        PauseSz = False
        '
    End If
    '
End Sub

Private Sub Form_DblClick()
    '
    Form_Click
    '
End Sub
'
Private Sub Form_KeyPress(KeyAscii As Integer)
    '
    Select Case KeyAscii
        '
        Case vbKeyEscape
            '
             bRunning = False
            '
        Case 102
            '
            If AffFps = True Then
                '
                AffFps = False
                '
            Else
                '
                AffFps = True
                '
            End If
            '
        '
    End Select
    '
End Sub
'
'lancement de la procédure principale
Public Sub LancerProcP()
    '
    'on initialise quelques variables
    PauseSz = False
    '
    'on lance le programme
    Initialise Me, ModeAffSzX, ModeAffSzY
    '
    Unload Me
    '
End Sub

Private Sub Form_Load()
    '
    Me.WindowState = 0
    '
    'Me.Top = 0
    'Me.Left = 0
    'Me.Width = Screen.Width
    'Me.Height = Screen.Height
    '
End Sub

'
Private Sub Form_Unload(Cancel As Integer)
    '
    bRunning = False
    '
End Sub
