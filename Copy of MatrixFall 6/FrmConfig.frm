VERSION 5.00
Begin VB.Form FrmConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuration de Matrix fall 2"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   Icon            =   "FrmConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   5715
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "Affichage"
      Height          =   345
      Left            =   3810
      TabIndex        =   28
      Top             =   1230
      Width           =   1800
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Textes"
      Height          =   345
      Left            =   1950
      TabIndex        =   27
      Top             =   1230
      Width           =   1830
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Options graphiques"
      Height          =   345
      Left            =   120
      TabIndex        =   26
      Top             =   1230
      Width           =   1800
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   120
      MousePointer    =   10  'Up Arrow
      Picture         =   "FrmConfig.frx":08CA
      ScaleHeight     =   1005
      ScaleWidth      =   5475
      TabIndex        =   18
      Top             =   120
      Width           =   5475
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   4020
      TabIndex        =   15
      Top             =   6150
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sauver la configuration"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   6120
      Width           =   2475
   End
   Begin VB.Frame Frame2 
      Caption         =   "Affichage : "
      Height          =   3915
      Left            =   120
      TabIndex        =   10
      Top             =   1740
      Width           =   5475
      Begin VB.TextBox Text6 
         Height          =   315
         Left            =   180
         TabIndex        =   30
         Top             =   1950
         Width           =   945
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   2790
         TabIndex        =   17
         Top             =   1200
         Width           =   945
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   2790
         TabIndex        =   13
         Top             =   780
         Width           =   945
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   2790
         TabIndex        =   12
         Top             =   360
         Width           =   945
      End
      Begin VB.Label Label12 
         Caption         =   $"FrmConfig.frx":127E4
         Height          =   1215
         Left            =   180
         TabIndex        =   32
         Top             =   2340
         Width           =   5115
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "nombre de lignes à afficher avant de relancer l'affichage des textes :"
         Height          =   195
         Left            =   180
         TabIndex        =   31
         Top             =   1680
         Width           =   4815
      End
      Begin VB.Label Label9 
         Caption         =   "nombre max d'images par seconde :"
         Height          =   195
         Left            =   180
         TabIndex        =   16
         Top             =   1260
         Width           =   2565
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "largeur des lettres :"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "hauteur des lettres :"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   420
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Textes : "
      Height          =   3915
      Left            =   120
      TabIndex        =   1
      Top             =   1740
      Width           =   5475
      Begin VB.CommandButton Command3 
         Caption         =   "Supprimer l'élément"
         Height          =   375
         Left            =   1860
         TabIndex        =   9
         Top             =   3360
         Width           =   1635
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   750
         TabIndex        =   6
         Text            =   "200"
         Top             =   2850
         Width           =   795
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Ajouter le texte"
         Height          =   375
         Left            =   3600
         TabIndex        =   4
         Top             =   3360
         Width           =   1635
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   750
         TabIndex        =   3
         Top             =   2400
         Width           =   4545
      End
      Begin VB.ListBox List1 
         Height          =   2010
         Left            =   180
         TabIndex        =   2
         Top             =   300
         Width           =   5115
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "ms  (la pause survient avant l'affichage du texte)"
         Height          =   195
         Left            =   1650
         TabIndex        =   8
         Top             =   2910
         Width           =   3405
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "pause :"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   2880
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "texte :"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   2460
         Width           =   435
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Options graphiques : "
      Height          =   3915
      Left            =   120
      TabIndex        =   19
      Top             =   1740
      Width           =   5475
      Begin VB.ComboBox CmbRendu 
         Height          =   315
         Left            =   180
         TabIndex        =   22
         Top             =   1230
         Width           =   5145
      End
      Begin VB.ComboBox CmbCartes 
         Height          =   315
         Left            =   180
         TabIndex        =   21
         Top             =   540
         Width           =   5145
      End
      Begin VB.ComboBox CmbAff 
         Height          =   315
         Left            =   180
         TabIndex        =   20
         Top             =   1860
         Width           =   5145
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Matériel disponible :"
         Height          =   195
         Left            =   180
         TabIndex        =   25
         Top             =   300
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Mode de rendu :"
         Height          =   195
         Left            =   180
         TabIndex        =   24
         Top             =   990
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Modes d'affichage :"
         Height          =   195
         Left            =   180
         TabIndex        =   23
         Top             =   1620
         Width           =   1395
      End
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "e-mail : thomas.john@swing.be"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   5760
      Width           =   5475
   End
End
Attribute VB_Name = "FrmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'******************************************************************
'*                                                                *
'* LA MAJEUR PARTIE DE CE CODE A ETE COPIE D'UNE SOURCE SE        *
'* TROUVANT A L'ADRESSE SUIVANTE :                                *
'* http://216.5.163.53/DirectX4VB/Downloads/DirectX8/Graph_02.zip *                                                           *
'*                                                                *
'* TRADUCTION ET MODIFICATIONS PAR THOMAS JOHN                    *
'*                                                                *
'******************************************************************
'
'cette variable va contenir le nombre de cartes vidéo ("adapter") que DirectX aura trouvé
Dim nAdapters As Long
'
'cette structure sert à contenir les informations concernant une carte vidéo ("adapter")
Dim AdapterInfo As D3DADAPTER_IDENTIFIER8
'
'nombre de modes d'affichage trouvé
Dim nModes As Long
'
Private Sub Command1_Click()
    '
    On Error Resume Next
    '
    'on efface l'ancien fichier s'il existe
    Kill App.Path & "\matrixfall.ini"
    '
    Dim FichSz As Integer
    Dim i As Integer
    '
    'on récupère un numéro de fichier libre
    FichSz = FreeFile
    '
    'on sauve la configuration choisie
    Open App.Path & "\matrixfall.ini" For Binary As #FichSz
    '
    'le mode d'accélération et d'affichage choisi
    Put #FichSz, 1, "accm : " & Left$(CmbRendu.List(CmbRendu.ListIndex), 3) & vbCrLf
    Put #FichSz, , "mode : " & CmbAff.List(CmbAff.ListIndex) & vbCrLf
    '
    'la carte de rendu choisie
    Put #FichSz, , "cart : " & Left$(CmbCartes.List(CmbCartes.ListIndex), 1) & vbCrLf
    '
    'la hauteur et la largeur des lettres
    Put #FichSz, , "haut : " & Text3.Text & vbCrLf
    Put #FichSz, , "larg : " & Text4.Text & vbCrLf
    '
    'le cycle de rendu
    Put #FichSz, , "cycl : " & Text5.Text & vbCrLf
    '
    'la limite de lignes affichées
    Put #FichSz, , "liml : " & Text6.Text & vbCrLf
    '
    'les textes
    If List1.ListCount > 0 Then
        '
        For i = 0 To List1.ListCount - 1
            '
            Put #FichSz, , "txtm : " & List1.List(i) & vbCrLf
            '
        Next
        '
    End If
    '
    Close #FichSz
    '
End Sub

Private Sub Command2_Click()
    '
    List1.AddItem Text2.Text & ";" & Text1.Text
    '
End Sub

Private Sub Command3_Click()
    '
    List1.RemoveItem List1.ListIndex
    '
End Sub
'
Private Sub Command4_Click()
    '
    Unload Me
    '
End Sub

Private Sub Command5_Click()
    '
    Frame1.Visible = False
    Frame2.Visible = False
    Frame3.Visible = True
    '
End Sub

Private Sub Command6_Click()
    '
    Frame1.Visible = True
    Frame2.Visible = False
    Frame3.Visible = False
    '
End Sub

Private Sub Command7_Click()
    '
    Frame1.Visible = False
    Frame2.Visible = True
    Frame3.Visible = False
    '
End Sub

'
Private Sub Form_Load()
    '
    'on ajoute à la liste les textes
    Dim i As Integer
    '
    For i = 0 To UBound(TexteSz)
        '
        'on vérifie que le texte ne soit pas vide
        If TexteSz(i) <> vbNullString Then List1.AddItem PauseTexte(i) & ";" & TexteSz(i)
        '
    Next
    '
    Text3.Text = HauteurLettreSz
    Text4.Text = LargeurLettreSz
    Text5.Text = CycleRenduSz
    Text6.Text = LimiteLignesAffSz
    '
    'on crée les objets utiles
    '
    'on crée notre objet principal
    Set Dx = New DirectX8
    '
    'on crée l'interface Direct3D via notre objet principal
    Set D3D = Dx.Direct3DCreate()
    '
    'on énumère le nombre de cartes disponibles
    EnumerateAdapters
    '
    'on vérifie si l'accélération matérielle est disponible
    EnumerateDevices
    '
    'on récupère les modes d'affichage disponibles
    EnumerateDispModes
    '
End Sub
'
'cette fonction récupère le nombre de cartes vidéo ("adapter") trouvées et les affiche
Private Sub EnumerateAdapters()
    '
    Dim i As Integer, sTemp As String, J As Integer
    '
    'on récupère le nombre (1 ou 2 dans la plupart des cas)
    nAdapters = D3D.GetAdapterCount
    '
    'on récupère les détails concernant la ou les carte(s) trouvée(s)
    For i = 0 To nAdapters - 1
        '
        D3D.GetAdapterIdentifier i, 0, AdapterInfo
        '
        'on récupère le nom de la carte courante
        sTemp = ""
        '
        For J = 0 To 511
            '
            sTemp = sTemp & Chr$(AdapterInfo.Description(J))
            '
        Next J
        '
        'on enlève les caractères indésirables
        sTemp = Replace(sTemp, Chr$(32), " ")
        '
        'on ajoute le nom à notre liste
        CmbCartes.AddItem i & " - " & sTemp
        '
        'on sélectionne le premier élément de la liste
        CmbCartes.ListIndex = 0
        '
    Next i
    '
End Sub
'
'nous permet de savoir si l'accélération matérielle est supportée
Private Sub EnumerateDevices()
    '
    On Local Error Resume Next
    '
    Dim Caps As D3DCAPS8
    '
    D3D.GetDeviceCaps CmbRendu.List(0), D3DDEVTYPE_HAL, Caps
    '
    If Err.Number = D3DERR_NOTAVAILABLE Then
        '
        'si il y a erreur, c'est que soit aucun carte n'est présente, soit la carte ne gère pas l'accélération matérielle
        'REF est toujours disponible
        CmbRendu.AddItem "Reference Rasterizer (REF)"
        '
    Else
        '
        'la carte gère l'accélération matérielle
        CmbRendu.AddItem "HAL - Hardware Acceleration"
        '
        'REF est toujours disponible
        CmbRendu.AddItem "REF - Reference Rasterizer"
        '
    End If
    '
    'on sélectionne le premier élément de la liste
    CmbRendu.ListIndex = 0
    '
End Sub
'
'énumère et affiche le nombre de mode d'affichage disponibles pour cette carte
Private Sub EnumerateDispModes()
    '
    'on efface le combo
    CmbAff.Clear
    '
    Dim i As Integer, ModeTemp As D3DDISPLAYMODE, Renderer As Long, sTmp As String
    '
    'on récupère d'abord le type d'accélération disponible (hardware ou software)
    If UCase(Left(CmbRendu.Text, 3)) = "REF" Then
        '
        Renderer = 2
        '
    Else
        '
        Renderer = 1
        '
    End If
    '
    'on récupère le nombre de modes d'affichage
    nModes = D3D.GetAdapterModeCount(CmbCartes.ListIndex)
    '
    'on fait une boucle afin de tous les ajouter au combobox
    For i = 0 To nModes - 1
        '
        Call D3D.EnumAdapterModes(CmbCartes.ListIndex, i, ModeTemp)
        '
        'on sépare les modes en 2 catégories (16 et 32 bits)
        If ModeTemp.Format = D3DFMT_R8G8B8 Or ModeTemp.Format = D3DFMT_X8R8G8B8 Or ModeTemp.Format = D3DFMT_A8R8G8B8 Then
            '
            'on vérifie si ce mode est acceptable et valide
            If D3D.CheckDeviceType(CmbCartes.ListIndex, Renderer, ModeTemp.Format, ModeTemp.Format, False) >= 0 Then
                '
                'si oui, on l'ajoute à notre liste s'il n'existe pas déjà
                If VerifElement(ModeTemp.Width & "x" & ModeTemp.Height & "x32") = -1 Then CmbAff.AddItem ModeTemp.Width & "x" & ModeTemp.Height & "x32" '& "    [FMT: " & ModeTemp.Format & "]"
                '
            End If
            '
        Else
            '
            'on fait la même chose qu'en haut
            If D3D.CheckDeviceType(CmbCartes.ListIndex, Renderer, ModeTemp.Format, ModeTemp.Format, False) >= 0 Then
                '
                If VerifElement(ModeTemp.Width & "x" & ModeTemp.Height & "x16") = -1 Then CmbAff.AddItem ModeTemp.Width & "x" & ModeTemp.Height & "x16" '& "    [FMT: " & ModeTemp.Format & "]"
                '
            End If
            '
        End If
        '
    Next i
    '
    'on sélectionne le bon élément
    i = VerifElement(ModeAffSzX & "x" & ModeAffSzY & "x" & ModeBit)
    '
    If i > 0 Then
        '
        CmbAff.ListIndex = i
        '
    Else
        '
        CmbAff.ListIndex = CmbAff.ListCount - 1
        '
    End If
    '
End Sub
'
Private Sub List1_Click()
    '
    Dim sTmp() As String
    '
    sTmp() = Split(List1.List(List1.ListIndex), ";")
    '
    Text1.Text = sTmp(1)
    Text2.Text = sTmp(0)
    '
End Sub
'
Private Sub Text1_Change()
    '
    If List1.ListIndex = -1 Then Exit Sub
    '
    List1.List(List1.ListIndex) = Text2.Text & ";" & Text1.Text
    '
End Sub
'
Private Sub Text2_Change()
    '
    If List1.ListIndex = -1 Then Exit Sub
    '
    List1.List(List1.ListIndex) = Text2.Text & ";" & Text1.Text
    '
End Sub
'
'fontion vérifiant l'existance d'un élément dans un combobox
'il renvoie 0 si l'élément est plus petit que 800x600
Private Function VerifElement(TexteSz As String) As Integer
    '
    Dim i As Integer
    '
    'on vérifie si l'élément est plus petit que 800x600
    If CLng(Replace(Left(TexteSz, 4), "x", "")) < 800 Then
        '
        VerifElement = 0
        '
        Exit Function
        '
    End If
    '
    If CmbAff.ListCount = -1 Then
        '
        VerifElement = -1
        '
        Exit Function
        '
    End If
    '
    For i = 0 To CmbAff.ListCount - 1
        '
        If LCase(TexteSz) = LCase(CmbAff.List(i)) Then
            '
            'il existe
            VerifElement = i
            '
            Exit Function
            '
        End If
        '
    Next
    '
    'si on arrive jusqu'ici, c'est qu'il n'a pas été trouvé
    VerifElement = -1
    '
End Function
