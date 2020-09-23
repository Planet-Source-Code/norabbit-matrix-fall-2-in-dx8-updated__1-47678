VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsTrainee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************************************
'*                                                                                *
'* CETTE CLASSE GERE UNE TRAINEE                                                  *
'*                                                                                *
'* Thomas John (thomas.john@swing.be)                                             *
'*                                                                                *
'**********************************************************************************
'
'structure vertex pour la trainée
Dim vertTrainee(3) As TLVERTEX
'
'stocke la géométrie d'un objet
Dim VBuffer As Direct3DVertexBuffer8
'
'numéro de la trainée
Dim NumTraineeSz As Long
'
'coordonnées de la trainée
Dim XSz As Single
Dim YSz As Single
'
'dimensions
Dim HauteurSz As Long
Dim LargeurSz As Long
'
'couleur de la trainée
Dim RSz As Long
Dim VSz As Long
Dim BSz As Long
'
'etat de la trainée
Dim EtatSz As Long
'
'opacité de la trainée
Dim OpaciteSz As Long
'
'opacité maximum à atteindre
Dim OpaciteSzMax As Long
'
'temps de transition 0 (opacité de la trainée s'accentue)
Dim TransTemps0 As Long
'temps écoulé
Dim TempsSz0 As Long
'
'temps de transition 3 (opacité de la trainée s'estompe)
Dim TransTemps3 As Long
'temps écoulé
Dim TempsSz3 As Long
'
'
'*********************************
'* PROPRIETES
'*********************************
'
'VITESSE A LAQUELLE L'OPACITE DE LA TRAINEE CHANGE
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
'OPACITE MAX
Public Property Let OpaciteMax(Valeur As Long)
    '
    OpaciteSzMax = Valeur
    '
End Property
'
Public Property Get OpaciteMax() As Long
    '
    OpaciteMax = OpaciteSzMax
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
'ETAT
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
'NUMERO DE LA TRAINEE
Public Property Let Trainée(Valeur As Long)
    '
    NumTraineeSz = Valeur
    '
End Property
'
Public Property Get Trainée() As Long
    '
    Trainée = NumTraineeSz
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
'procédure permettan d'affiche la trainée choisie
Private Sub AfficherTrainee(NuméroSz As Long, Optional CouleurSz As Long = vbWhite)
    '
    Dim CharX As Integer
    Dim CharY As Integer
    Dim cTmp As Long
    '
    'on calcule les coordonnées
    cTmp = (NuméroSz - 1) \ 13
    CharY = 134 * cTmp
    cTmp = (NuméroSz - 1) Mod 13
    CharX = 19 * cTmp
    '
    'on détermine la couleur en y ajoutant la valeur "alpha"
    CouleurSz = Hex2Long(Hex(OpaciteSz) & Hex(CouleurSz))
    '
    'on génère les points
    vertTrainee(0) = CreateTLVertex(XSz, YSz, 0, 1, CouleurSz, 0, (1 / 256) * CharX, (1 / 256) * CharY)
    vertTrainee(1) = CreateTLVertex(XSz + LargeurSz, YSz, 0, 1, CouleurSz, 0, ((1 / 256) * CharX) + ((1 / 256) * 19), (1 / 256) * CharY)
    vertTrainee(2) = CreateTLVertex(XSz, YSz + HauteurSz, 0, 1, CouleurSz, 0, (1 / 256) * CharX, ((1 / 256) * CharY) + ((1 / 256) * 134))
    vertTrainee(3) = CreateTLVertex(XSz + LargeurSz, YSz + HauteurSz, 0, 1, CouleurSz, 0, ((1 / 256) * CharX) + ((1 / 256) * 19), ((1 / 256) * CharY) + ((1 / 256) * 134))
    '
    'on dessine l'image
    D3DDevice.SetTexture 0, MatrixTex_Trainee
    '
    Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_MODULATE)
    'Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAARG1, D3DTA_DIFFUSE)
    'Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAARG2, D3DTA_CURRENT)
    '
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, vertTrainee(0), Len(vertTrainee(0))
    '
End Sub
'
'procédure qui va s'occuper d'afficher la trainée dans le temps
Public Sub AfficherTemps()
    '
    'on ajuste le rendu de la trainée en fonction de son état
    Select Case EtatSz
        '
        Case 1 'trainée blanche
            '
            RSz = 255
            VSz = 255
            BSz = 255
            '
        Case 2 'trainée verte
            '
            RSz = 124
            VSz = 209
            BSz = 144
            '
        Case 3 'on diminue l'opacité
            '
            If TempsSz3 < TransTemps3 Then
                '
                OpaciteSz = OpaciteSz - (OpaciteSz / (TransTemps3 - TempsSz3))
                '
                AfficherTrainee NumTraineeSz, RGB(RSz, VSz, BSz)
                '
                TempsSz3 = TempsSz3 + 1
                '
            End If
            '
            'on sort de la procédure afin de ne pas effectuer la fonction qui suit
            Exit Sub
            '
        '
    End Select
    '
    'l'opacité de la lettre s'accentue
    If TempsSz0 < TransTemps0 Then
        '
        'on change la couleur de la lettre en fonction du temps que cela doit prendre
        OpaciteSz = OpaciteSz - ((OpaciteSz - OpaciteSzMax) / (TransTemps0 - TempsSz0))
        '
        AfficherTrainee NumTraineeSz, RGB(RSz, VSz, BSz)
        '
        TempsSz0 = TempsSz0 + 1
        '
        'on sort de la procédure afin de ne pas réafficher la trainée
        Exit Sub
        '
    End If
    '
    'on affiche la trainée
    AfficherTrainee NumTraineeSz, RGB(RSz, VSz, BSz)
    '
End Sub
'
'initialisation des variables
Public Sub Init()
    '
    TempsSz0 = 0
    '
    TempsSz3 = 0
    '
    OpaciteSz = 0
    '
End Sub

Private Sub Class_Initialize()
    '
    TransTemps0 = 20
    OpaciteSz = 0
    '
End Sub
