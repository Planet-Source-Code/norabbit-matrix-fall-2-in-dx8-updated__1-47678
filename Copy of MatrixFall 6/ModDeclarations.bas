Attribute VB_Name = "ModDeclarations"
'*************************************************************************
'*                                                                       *
'* MODULE DX8                                                            *
'*                                                                       *
'* DECLARATION DES VARIABLES PRINCIPALES + FONCTIONS                     *
'*                                                                       *
'* traduit et compl�t� par Thomas John (thomas.john@swing.be)            *
'*                                                                       *
'* source : http://216.5.163.53/DirectX4VB (DirectX 4 VB, Jack Hoxley)   *
'*                                                                       *
'*************************************************************************
'
'l'objet principal
Public Dx As DirectX8
'
'cet objet contr�le tout ce qui est 3D
Public D3D As Direct3D8
'
'cet objet repr�sente le "hardware" (la carte graphique) utilis� pour le rendu
Public D3DDevice As Direct3DDevice8
'
'une "librairie d'aide"
'D3DX8 est une classe d'aide qui contient une multitude de fonctions destin�es � faciliter la programmation en DX8
Public D3DX As D3DX8
'
'variable servant � d�tecter si le programme tourne ou pas
Public bRunning As Boolean
'
'description des diff�rents types de vertex
Public Const FVF_TLVERTEX = (D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR)
Public Const FVF_LVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX1)
Public Const FVF_VERTEX = (D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1)
'
'cette structure repr�sente un vertex 2D (identique � la structure "D3DTLVERTEX" pour Dx7)
Public Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    rhw As Single
    Color As Long
    Specular As Long
    tU As Single
    tV As Single
End Type
'
'autre type de vertex
Public Type LITVERTEX
    X As Single
    Y As Single
    Z As Single
    Color As Long
    Specular As Long
    tU As Single
    tV As Single
End Type
'
'autre type de vertex
Public Type VERTEX
    X As Single
    Y As Single
    Z As Single
    nx As Single
    ny As Single
    nz As Single
    tU As Single
    tV As Single
End Type
'
'fonte
Public MainFont As D3DXFont
Public MainFontDesc As IFont
Public TextRect As RECT
Public fnt As New StdFont
'
'Pi
Public Const pi As Single = 3.14159265358979
'
Public matWorld As D3DMATRIX '//How the vertices are positioned
'o� la cam�ra se trouve et vers o� pointe-t-elle
Public matView As D3DMATRIX
'comment la cam�ra projecte le monde 3D sur un �cran (surface) 2D
Public matProj As D3DMATRIX
'
'calcul du fps (images par seconde)
Public FPS_NbrFps As Long
Public FPS_NbrImg As Long
Public lFpsTmp As Long
'
'dimensions de l'affichage
Public DimH As Long
Public DimL As Long
'
'infos
Public InfoSz As String
'
'classe se chargeant de faire appara�tre du texte
Public ClsTxt As New ClsTexte
'
'classe se chargeant de faire appara�tre plusieurs lignes (public car elle sera utilis�e par une autre classe -> ClsTxt)
Public ClsPlusL As New ClsPlusLignes
'
'texture de fonte matrix
Public MatrixTex_Blanc As Direct3DTexture8
Public MatrixTex_Vert As Direct3DTexture8
Public MatrixTex_Trainee As Direct3DTexture8
Public MatrixTex_Normal As Direct3DTexture8
'
'pause
Public PauseSz As Boolean
'
'affichage du nombre d'images par seconde
Public AffFps As Boolean
'
'permet de ralentir la vitesse d'affichage � un certain nombre d'images par seconde
Public FpsMod As Long
'
'liste contenant toutes les coordonn�es X utilis�es par les lignes
Public ListeCooX As New Collection
'
'dimensions des lettres
Public HauteurLettreSz As Long
Public LargeurLettreSz As Long
'
'mode d'affichage
Public ModeAffSzX As Long
Public ModeAffSzY As Long
Public ModeBit As Long
'
'acc�l�ration mat�rielle ou software
Public AccMatSoft As String
'
'carte choisie
Public CarteChoixSz As Long
'
'textes
Public TexteSz() As String
'
'pause entre l'affichage de chaque texte
Public PauseTexte() As Long
'
'cycle de rendu en millisecondes (temps d'attente entre chaque rendu --> diminue le nombre d'image par seconde et donc la vitesse)
Public CycleRenduSz As Long
'
'limite de lignes affich�es
Public LimiteLignesAffSz As Long
'
'
'
'*******************************************************************
'* Initialise
'*******************************************************************
'
Public Function Initialise(FrmObjet As Form, DimLargeur As Long, DimHauteur As Long) As Boolean
    '
    On Error Resume Next
    '
    'd�crit notre mode d'affichage
    Dim DispMode As D3DDISPLAYMODE
    '
    'd�crit notre mode de vue
    Dim D3DWindow As D3DPRESENT_PARAMETERS
    '
    'pour les filtreurs de texture
    Dim Caps As D3DCAPS8 '//For Texture Filters
    '
    'on cr�e notre objet principal
    Set Dx = New DirectX8
    '
    'on cr�e l'interface Direct3D via notre objet principal
    Set D3D = Dx.Direct3DCreate()
    '
    'on cr�e notre librairie d'aide
    Set D3DX = New D3DX8
    '
    'DispMode.Format = D3DFMT_X8R8G8B8
    'DispMode.Format = D3DFMT_A8R8G8B8
    DispMode.Format = D3DFMT_R5G6B5 'si ce mode ne fonctionne pas, utilisez celui juste au-dessus
    DispMode.Width = DimLargeur
    DispMode.Height = DimHauteur
    '
    DimL = DimLargeur
    DimH = DimHauteur
    '
    D3DWindow.BackBufferCount = 1 '1 BackBuffer
    D3DWindow.BackBufferWidth = DispMode.Width
    D3DWindow.BackBufferHeight = DispMode.Height
    D3DWindow.hDeviceWindow = FrmObjet.hWnd
    D3DWindow.EnableAutoDepthStencil = 1
    D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
    D3DWindow.BackBufferFormat = DispMode.Format
    '
    If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, D3DRTYPE_SURFACE, D3DFMT_D32) = D3D_OK Then
        '
        'on peut utiliser un Z-Buffer de 16-bit
        D3DWindow.AutoDepthStencilFormat = D3DFMT_D32
        '
    Else
        '
        If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, D3DRTYPE_SURFACE, D3DFMT_D24) = D3D_OK Then
            '
            'on peut utiliser un Z-Buffer de 16-bit
            D3DWindow.AutoDepthStencilFormat = D3DFMT_D24
            '
        Else
            '
            If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, D3DRTYPE_SURFACE, D3DFMT_D16) = D3D_OK Then
                '
                'on peut utiliser un Z-Buffer de 16-bit
                D3DWindow.AutoDepthStencilFormat = D3DFMT_D16
                '
            End If
            '
        End If
        '
    End If
    '
    'on montre notre feuille pour �tre s�r
    FrmObjet.Show
    '
    'on la met au-dessus de toutes
    SetWindowPos FrmObjet.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    '
    'cette ligne cr�e un "device" qui utilise la carte graphique ("hardware") pour effectuer les calculs si possible,
    'ou le processeur ("software") et utilise comme objet de r�ception notre feuille principale
    'on lance le mode hardware ou software selon les options charg�es
    Select Case AccMatSoft
        '
        Case "REF"
            '
            Set D3DDevice = D3D.CreateDevice(CarteChoixSz, D3DDEVTYPE_REF, FrmObjet.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)
            '
        Case "HAL"
            '
            Set D3DDevice = D3D.CreateDevice(CarteChoixSz, D3DDEVTYPE_HAL, FrmObjet.hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, D3DWindow)
            '
        '
    End Select
    '
    'on demande au vertex shader d'utiliser notre format de vertex
    D3DDevice.SetVertexShader FVF_TLVERTEX
    '
    'nos points (vertices) n'ont pas besoin de lumi�re, donc on d�sactive cette option
    D3DDevice.SetRenderState D3DRS_LIGHTING, False
    '
    'D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    '
    'd�clarations utiles pour le rendu de textures transparantes
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
    '
    'filtrage de texture : donne un meilleur r�sultat lors d'un redimensionnement d'une texture
    D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
    D3DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
    '
    'on active notre Z-Buffer
    D3DDevice.SetRenderState D3DRS_ZENABLE, 1
    '
    '
    'la matrice "World"
    D3DXMatrixIdentity matWorld
    D3DDevice.SetTransform D3DTS_WORLD, matWorld
    '
    'la matrice "View"
    D3DXMatrixLookAtLH matView, MakeVector(0, 9, -9), MakeVector(0, 0, 0), MakeVector(0, 1, 0)
    D3DDevice.SetTransform D3DTS_VIEW, matView
    '
    'la matrice de projection
    D3DXMatrixPerspectiveFovLH matProj, pi / 4, 1, 0.1, 500
    D3DDevice.SetTransform D3DTS_PROJECTION, matProj
    '
    '
    'initialisation du rendu du texte
    fnt.Name = "Tahoma"
    fnt.Size = 12
    fnt.Bold = True
    Set MainFontDesc = fnt
    Set MainFont = D3DX.CreateFont(D3DDevice, MainFontDesc.hFont)
    TextRect.Top = 1
    TextRect.Left = 1
    TextRect.bottom = 480
    TextRect.Right = 640
    '
    '**************************************
    '** chargement des textures          **
    '**************************************
    '
    Set MatrixTex_Blanc = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\matrixfall\fontes_blanches.png", 256, 256, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, 0, ByVal 0, ByVal 0)
    Set MatrixTex_Vert = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\matrixfall\fontes_vertes.png", 256, 256, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, 0, ByVal 0, ByVal 0)
    Set MatrixTex_Trainee = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\matrixfall\trainee.png", 256, 256, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, 0, ByVal 0, ByVal 0)
    Set MatrixTex_Normal = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\matrixfall\fontes_normales.png", 256, 256, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, 0, ByVal 0, ByVal 0)
    '
    '**************************************
    '** fin du chargement des textures   **
    '**************************************
    '
    'on cache le curseur
    ShowCursor False
    '
    'si on arrive jusqu'ici, c'est que tout s'est bien pass�
    Initialise = True
    '
    'on g�re les erreurs survenues durant l'initialisation ici
    If Err Then
        '
        EcrireLog "### " & Time & " ###" & vbCrLf & "erreur lors de l'initialisation :" & vbCrLf & D3DX.GetErrorString(Err.Number) & vbCrLf
        '
    End If
    '
    Dim RotateAngle As Single
    Dim matTemp As D3DMATRIX 'contient des donn�es temporaires
    '
    bRunning = True
    '
    '-1 pour �viter la division par z�ro
    lFpsTmp = GetTickCount - 1
    '
    Do While bRunning
        '
        'on v�rifie si il y a une pause
        If PauseSz = False Then
            '
            'nombre d'images par secondes
            FPS_NbrFps = FPS_NbrImg / ((GetTickCount - lFpsTmp) / 1000)
            '
            'on v�rifie que la variable qui contient le nombre d'images rendues ne soit pas trop grande
            If FPS_NbrImg > 1000000 Then
                '
                FPS_NbrImg = 0
                '
                lFpsTmp = GetTickCount - 1
                '
            End If
            '
            'on rend la sc�ne si le nombre d'images par seconde a �t� atteint
            If FPS_NbrFps <= CycleRenduSz Then
                '
                'on incr�ment le nombre d'images rendues
                FPS_NbrImg = FPS_NbrImg + 1
                '
                'on "rend" (dessine) la sc�ne
                Render
                '
            End If
            '
        End If
        '
        'on laisse vb respirer
        DoEvents
        '
    Loop
    '
    'on g�re les erreurs survenues lors du rendu ici
    If Err Then
        '
        EcrireLog "### " & Time & " ###" & vbCrLf & "erreur lors du rendu :" & vbCrLf & D3DX.GetErrorString(Err.Number) & vbCrLf
        '
    End If
    '
    'on affiche le curseur
    ShowCursor True
    '
    'la boucle s'est termin�e signifiant que le programme va se fermer
    'il faut avant tout d�charger les objets qu'on a charg� pr�c�demment
    '
    On Error Resume Next 'pour �tre s�r
    '
    Set D3DDevice = Nothing
    Set D3D = Nothing
    Set Dx = Nothing
    Set D3DX = Nothing
    '
    'on g�re les erreurs survenues lors du d�chargement des objets ici
    If Err Then
        '
        EcrireLog "### " & Time & " ###" & vbCrLf & "erreur lors du d�chargement des objets :" & vbCrLf & D3DX.GetErrorString(Err.Number) & vbCrLf
        '
    End If
    '
End Function
'
'PROCEDURE DE RENDU DE LA SCENE
Public Sub Render()
    '
    Dim i As Integer
    '
    'on efface la surface et on lui donne la couleur noir
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1#, 0
    '
    'on commence le rendu
    D3DDevice.BeginScene
        '
        'tous les appels de rendu doivent �tre fait entre "BeginScene" et "EndScene"
        '
        'rendu du texte
        If AffFps = True Then D3DX.DrawText MainFont, &HFF78B478, FPS_NbrFps & " fps", TextRect, DT_LEFT
        '
        'on affiche le texte
        ClsTxt.Afficher
        '
        'on affiche les lignes
        ClsPlusL.Afficher
        '
    D3DDevice.EndScene
    '
    'on met � jour l'image � l'�cran
    D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
    '
End Sub
'
'fonction permettant de cr�er un vecteur en une ligne
Public Function MakeVector(X As Single, Y As Single, Z As Single) As D3DVECTOR
    '
    MakeVector.X = X
    MakeVector.Y = Y
    MakeVector.Z = Z
    '
End Function
'
'fontion permettant de remplir un objet en une seule ligne
Public Function CreateTLVertex(X As Single, Y As Single, Z As Single, rhw As Single, Color As Long, Specular As Long, tU As Single, tV As Single) As TLVERTEX
    '
    '//NB: whilst you can pass floating point values for the coordinates (single)
    '       there is little point - Direct3D will just approximate the coordinate by rounding
    '       which may well produce unwanted results....
    '
    With CreateTLVertex
        '
        .X = X
        .Y = Y
        .Z = Z
        .rhw = rhw
        .Color = Color
        .Specular = Specular
        .tU = tU
        .tV = tV
        '
    End With
    '
End Function
'
'fontion permettant de remplir un objet en une seule ligne
Public Function CreateLitVertex(X As Single, Y As Single, Z As Single, Color As Long, Specular As Long, tU As Single, tV As Single) As LITVERTEX
    '
    With CreateLitVertex
        '
        .X = X
        .Y = Y
        .Z = Z
        .Color = Color
        .Specular = Specular
        .tU = tU
        .tV = tV
        '
    End With
End Function
'
'fonction renvoyant un nombre al�atoire situ� entre un minimum et un maximum
'(merci � raff (VbFrance) pour le code disponible � l'adresse suivante : http://www.vbfrance.com/article.aspx?ID=7209)
'(code l�g�rement modifi�)
Public Function RndNbr(MinSz As Long, MaxSz As Long) As Long

     RndNbr = (Rnd * (MaxSz - MinSz)) + MinSz
                
End Function
'
'convertit une donn�e hex en long
Public Function Hex2Long(hHex) As Long
    '
    Hex2Long = "&H" & hHex
    '
End Function
'
'lecture des options
Public Sub ChargerOptions()
    '
    On Error Resume Next
    '
    Dim FichSz As Integer
    Dim sTmp As String
    Dim sTmp2() As String
    Dim sTmp3() As String
    '
    FichSz = FreeFile
    '
    Open App.Path & "\matrixfall.ini" For Binary As #FichSz
    '
    sTmp = Space(LOF(FichSz))
    '
    Get #FichSz, , sTmp
    '
    Close #FichSz
    '
    'on r�cup�re chaque information en s�parant celles-ci par "VbCrLf"
    sTmp2() = Split(sTmp, vbCrLf)
    '
    'si le tableau ne contient aucune donn�e, on va directement � la fin de la proc�dure
    If UBound(sTmp2) = -1 Then GoTo FIN_PROC
    '
    'je r�utilise "FichSz" comme variable d'incr�mentation
    For FichSz = 0 To UBound(sTmp2)
        '
        If sTmp2(FichSz) = vbNullString Then GoTo FIN_PROC
        '
        'on r�cup�re les infos en fonction de leur nom
        Select Case Left$(sTmp2(FichSz), 4)
            '
            'type d'acc�l�ration (hardware ou software)
            Case "accm"
                '
                AccMatSoft = Right(sTmp2(FichSz), 3)
                '
            'mode d'affichage
            Case "mode"
                '
                sTmp3() = Split(Mid(sTmp2(FichSz), 8, Len(sTmp2(FichSz)) - 7), "x")
                '
                'on v�rifie que toutes les infos sont l�
                If UBound(sTmp3) = 2 Then
                    '
                    ModeAffSzX = CLng(sTmp3(0))
                    ModeAffSzY = CLng(sTmp3(1))
                    ModeBit = CLng(sTmp3(2))
                    '
                Else
                    '
                    'sinon, on met le mode par d�faut
                    ModeAffSzX = 800
                    ModeAffSzY = 600
                    ModeBit = 16
                    '
                End If
                '
            Case "haut"
                '
                HauteurLettreSz = CLng(Mid(sTmp2(FichSz), 8, Len(sTmp2(FichSz)) - 7))
                '
            Case "larg"
                '
                LargeurLettreSz = CLng(Mid(sTmp2(FichSz), 8, Len(sTmp2(FichSz)) - 7))
                '
            Case "cycl"
                '
                CycleRenduSz = CLng(Mid(sTmp2(FichSz), 8, Len(sTmp2(FichSz)) - 7))
                '
            Case "cart"
                '
                CarteChoixSz = CLng(Mid(sTmp2(FichSz), 8, Len(sTmp2(FichSz)) - 7))
                '
            Case "liml"
                '
                LimiteLignesAffSz = CLng(Mid(sTmp2(FichSz), 8, Len(sTmp2(FichSz)) - 7))
                '
            'texte + pause
            Case "txtm"
                '
                'on s�pare le texte de la pause
                sTmp3() = Split(Mid(sTmp2(FichSz), 8, Len(sTmp2(FichSz)) - 7), ";", 2)
                '
                'on v�rifie que les 2 informations sont l�
                If UBound(sTmp3) = 1 Then
                    '
                    'on les ajoute � la liste
                    '
                    'on v�rifie que l'index 0 du tableau TexteSz contient bien quelque chose
                    If TexteSz(0) = vbNullString Then
                        '
                        TexteSz(0) = sTmp3(1)
                        PauseTexte(0) = CLng(sTmp3(0))
                        '
                    Else
                        '
                        ReDim Preserve TexteSz(0 To UBound(TexteSz) + 1)
                        ReDim Preserve PauseTexte(0 To UBound(PauseTexte) + 1)
                        TexteSz(UBound(TexteSz)) = sTmp3(1)
                        PauseTexte(UBound(PauseTexte)) = CLng(sTmp3(0))
                        '
                    End If
                    '
                End If
                '
            '
        End Select
        '
    Next
    '
FIN_PROC:
    '
    'on v�rifie que les informations importantes sont pr�sentes
    If HauteurLettreSz = 0 Then
        '
        HauteurLettreSz = 25
        '
    End If
    '
    If LargeurLettreSz = 0 Then
        '
        LargeurLettreSz = HauteurLettreSz * 1.3
        '
    End If
    '
    If ModeAffSzX < 800 Or ModeAffSzY < 600 Or ModeBit = 0 Then
        '
        ModeAffSzX = 800
        ModeAffSzY = 600
        ModeBit = 16
        '
    End If
    '
    If LimiteLignesAffSz <= 0 Then
        '
        LimiteLignesAffSz = 200
        '
    End If
    '
End Sub
'
'fonction inscrivant dans le fichier log_matrix_fall.txt les erreurs et autres
Public Sub EcrireLog(TexteSz As String)
    '
    Dim FichSz As Integer
    '
    FichSz = FreeFile
    '
    Open App.Path & "\log_matrix_fall.txt" For Binary As #FichSz
    Seek #FichSz, LOF(FichSz)
    Put #FichSz, , TexteSz & vbCrLf
    Close #FichSz
    '
End Sub
