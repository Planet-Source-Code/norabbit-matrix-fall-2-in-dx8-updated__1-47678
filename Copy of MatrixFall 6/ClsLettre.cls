VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLettre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************************************
'*                                                                                *
'* CETTE CLASSE GERE UNE SEULE LETTRE                                             *
'*                                                                                *
'* Thomas John (thomas.john@swing.be)                                             *
'*                                                                                *
'**********************************************************************************
'
'structure vertex pour la fonte matrix
Dim vertMatrix(3) As TLVERTEX
'
'stocke la g�om�trie d'un objet
Dim VBuffer As Direct3DVertexBuffer8
'
'opacit� de la lettre
Dim OpaciteSz As Long
'
'num�ro de la lettre
Dim NumLettreSz As Long
'
'coordonn�es de la lettre
Dim XSz As Single
Dim YSz As Single
'
'dimensions
Dim HauteurSz As Long
Dim LargeurSz As Long
'
'couleur de la lettre
Dim RSz As Long
Dim VSz As Long
Dim BSz As Long
'
'temps de transition 0 (du noir au blanc)
Dim TransTemps0 As Long
'temps �coul�
Dim TempsSz0 As Long
'
'temps de transition 1 (caract�re lumineux)
Dim TransTemps1 As Long
'temps �coul�
Dim TempsSz1 As Long
'
'temps de transition 2
Dim TransTemps2 As Long
'temps �coul� 2
Dim TempsSz2 As Long
'
'temps de transition 3 (opacit� de la lettre s'estompe)
Dim TransTemps3 As Long
'temps �coul�
Dim TempsSz3 As Long
'
'temps de transition 4 (changement de lettre)
Dim TransTemps4 As Long
'temps �coul�
Dim TempsSz4 As Long
'
'la lettre peut elle changer dans le temps
Dim ChangementSz As Long
'
'variations de couleur
Dim VarCouleur As Long
'
'�tat de la lettre
Dim EtatSz As Long
'
'
'*********************************
'* PROPRIETES
'*********************************
'
'ETAT DE LA LETTRE
Public Property Let Etat(Valeur As Long)
    '
    EtatSz = Valeur
    '
End Property
'
Public Property Get Etat() As Long
    '
    Etat = EtatSz
    '
End Property
'
'VARIATIONS DE COULEUR
Public Property Let Variation(Valeur As Long)
    '
    VarCouleur = Valeur
    '
End Property
'
Public Property Get Variation() As Long
    '
    Variation = VarCouleur
    '
End Property
'
'VITESSE A LAQUELLE LA LETTRE CHANGE
Public Property Let Transition4(Valeur As Long)
    '
    TransTemps4 = Valeur
    '
End Property
'
Public Property Get Transition4() As Long
    '
    Transition4 = TransTemps4
    '
End Property
'
'VITESSE A LAQUELLE L'OPACITE DE LA LETTRE CHANGE
Public Property Let Transition3(Valeur As Long)
    '
    TransTemps3 = Valeur
    '
End Property
'
Public Property Get Transition3() As Long
    '
    Transition3 = TransTemps3
    '
End Property
'
'TEMPS DE LA 2EME TRANSITION
Public Property Let Transition2(Valeur As Long)
    '
    TransTemps2 = Valeur
    '
End Property
'
Public Property Get Transition2() As Long
    '
    Transition2 = TransTemps2
    '
End Property
'
'TEMPS DE VIE DU CARACTERE LUMINEUX
Public Property Let Transition1(Valeur As Long)
    '
    TransTemps1 = Valeur
    '
End Property
'
Public Property Get Transition1() As Long
    '
    Transition1 = TransTemps1
    '
End Property
'
'TEMPS QUE LA LETTRE VA METTRE A APPARAITRE
Public Property Let Transition0(Valeur As Long)
    '
    TransTemps0 = Valeur
    '
End Property
'
Public Property Get Transition0() As Long
    '
    Transition0 = TransTemps0
    '
End Property
'
'LA LETTRE PEUT ELLE CHANGER DANS LE TEMPS
Public Property Let Changement(Valeur As Long)
    '
    ChangementSz = Valeur
    '
End Property
'
Public Property Get Changement() As Long
    '
    Changement = ChangementSz
    '
End Property
'
'NUMERO DE LA LETTRE
Public Property Let Lettre(Valeur As Long)
    '
    NumLettreSz = Valeur
    '
End Property
'
Public Property Get Lettre() As Long
    '
    Lettre = NumLettreSz
    '
End Property
'
'OPACITE
Public Property Let Opacite(Valeur As Long)
    '
    OpaciteSz = Valeur
    '
End Property
'
Public Property Get Opacite() As Long
    '
    Opacite = OpaciteSz
    '
End Property
'
'COORDONNEE X
Public Property Let X(Valeur As Single)
    '
    XSz = Valeur
    '
End Property
'
Public Property Get X() As Single
    '
    X = XSz
    '
End Property
'
'COORDONNEE Y
Public Property Let Y(Valeur As Single)
    '
    YSz = Valeur
    '
End Property
'
Public Property Get Y() As Single
    '
    Y = YSz
    '
End Property
'
'HAUTEUR
Public Property Let Hauteur(Valeur As Long)
    '
    HauteurSz = Valeur
    '
End Property
'
Public Property Get Hauteur() As Long
    '
    Hauteur = HauteurSz
    '
End Property
'
'LARGEUR
Public Property Let Largeur(Valeur As Long)
    '
    LargeurSz = Valeur
    '
End Property
'
Public Property Get Largeur() As Long
    '
    Largeur = LargeurSz
    '
End Property
'
'*********************************
'* FIN DES PROPRIETES
'*********************************
'
'proc�dure permettan d'affiche la fonte choisie (via une texture)
Private Sub AfficherFonte(Num�roFonte As Long, Optional CouleurSz As Long = vbWhite)
    '
    Dim CharX As Integer
    Dim CharY As Integer
    Dim cTmp As Long
    '
    'on calcule les coordonn�es
    cTmp = (Num�roFonte - 1) \ 13
    CharY = 25 * cTmp
    cTmp = (Num�roFonte - 1) Mod 13
    CharX = 19 * cTmp
    '
    'on d�termine la couleur (on ajoute la valeur alpha)
    CouleurSz = Hex2Long(Hex(OpaciteSz) & Hex(CouleurSz))
    '
    'on g�n�re les points
    vertMatrix(0) = CreateTLVertex(XSz, YSz, 0, 1, CouleurSz, 0, (1 / 256) * CharX, (1 / 256) * CharY)
    vertMatrix(1) = CreateTLVertex(XSz + LargeurSz, YSz, 0, 1, CouleurSz, 0, ((1 / 256) * CharX) + ((1 / 256) * 19), (1 / 256) * CharY)
    vertMatrix(2) = CreateTLVertex(XSz, YSz + HauteurSz, 0, 1, CouleurSz, 0, (1 / 256) * CharX, ((1 / 256) * CharY) + ((1 / 256) * 25))
    vertMatrix(3) = CreateTLVertex(XSz + LargeurSz, YSz + HauteurSz, 0, 1, CouleurSz, 0, ((1 / 256) * CharX) + ((1 / 256) * 19), ((1 / 256) * CharY) + ((1 / 256) * 25))
    '
    'on dessine l'image
    Select Case EtatSz
        '
        Case 1 'lettre blanche
            '
            D3DDevice.SetTexture 0, MatrixTex_Blanc
            '
        Case 2 'lettre verte
            '
            D3DDevice.SetTexture 0, MatrixTex_Vert
            '
        '
    End Select
    '
    Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_MODULATE)
    'Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAARG1, D3DTA_DIFFUSE)
    'Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAARG2, D3DTA_CURRENT)
    '
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, vertMatrix(0), Len(vertMatrix(0))
    '
End Sub
'
'proc�dure qui va s'occuper d'afficher la lettre dans le temps
Public Sub AfficherTemps()
    '
    'tant que la lettre est blanche, on se contente de l'afficher sans la modifier
    If EtatSz = 1 Then
        '
        AfficherFonte NumLettreSz
        '
        'on quite la proc�dure ici
        Exit Sub
        '
    End If
    '
    'transition du noir vers le blanc
    'If TempsSz0 < TransTemps0 Then
        '
        'on change l'opacit� de la lettre en fonction du temps que cela doit prendre
        'OpaciteSz = OpaciteSz - ((OpaciteSz - 255) / (TransTemps0 - TempsSz0))
        '
        'AfficherFonte NumLettreSz, RGB(RSz, VSz, BSz)
        '
        'TempsSz0 = TempsSz0 + 1
        '
        'Exit Sub
        '
    'End If
    '
    'transition de la couleur blanche vers la couleur verte
    'If TempsSz1 < TransTemps1 Then
        '
        'on change l'�tat
        'EtatSz = 2
        '
        'on change la couleur de la lettre en fonction du temps que cela doit prendre
        'RSz = RSz - ((RSz - (255 + VarCouleur)) / (TransTemps1 - TempsSz1))
        'VSz = VSz - ((VSz - (255 + VarCouleur)) / (TransTemps1 - TempsSz1))
        'BSz = BSz - ((BSz - (255 + VarCouleur)) / (TransTemps1 - TempsSz1))
        '
        'AfficherFonte NumLettreSz, RGB(RSz, VSz, BSz)
        '
        'TempsSz1 = TempsSz1 + 1
        '
        'Exit Sub
        '
    'End If
    '
    'on effectue la transition suivante
    If TempsSz2 < TransTemps2 Then
        '
        'ici, la couleur reste la m�me
        AfficherFonte NumLettreSz, RGB(RSz + VarCouleur, VSz + VarCouleur, BSz + VarCouleur)
        '
        TempsSz2 = TempsSz2 + 1
        '
        'si la lettre peut changer dans le temps, on la change
        If ChangementSz = 1 Then
            '
            'on attend que le temps n�cessaire avant d'effectuer le changement se soit �coul�
            If TempsSz4 = TransTemps4 Then
                '
                NumLettreSz = RndNbr(1, 52)
                '
                TempsSz4 = 0
                '
            Else
                '
                TempsSz4 = TempsSz4 + 1
                '
            End If
            '
        End If
        '
        '
        Exit Sub
        '
    End If
    '
    'on effectue la transition suivante
    If TempsSz3 < TransTemps3 Then
        '
        'ici, l'opacit� s'estompe
        OpaciteSz = OpaciteSz - (OpaciteSz / (TransTemps3 - TempsSz3))
        '
        AfficherFonte NumLettreSz, RGB(RSz, VSz, BSz)
        '
        TempsSz3 = TempsSz3 + 1
        '
        Exit Sub
        '
    End If
    '
    'la "vie" de la lettre se termine ici
    EtatSz = 0
    '
End Sub

Private Sub Class_Initialize()
    '
    TransTemps0 = 1
    '
    EtatSz = 1
    '
    OpaciteSz = 255
    '
    RSz = 255
    VSz = 255
    BSz = 255
    '
End Sub
