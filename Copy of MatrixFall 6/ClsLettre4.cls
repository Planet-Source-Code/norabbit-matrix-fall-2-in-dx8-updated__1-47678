VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLettre4"
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
'stocke la géométrie d'un objet
Dim VBuffer As Direct3DVertexBuffer8
'
'opacité de la lettre
Dim OpaciteSz As Long
'
'numéro de la lettre
Dim NumLettreSz As Long
'
'coordonnées de la lettre
Dim XSz As Single
Dim YSz As Single
'
'dimensions
Dim HauteurSz As Single
Dim LargeurSz As Single
'
'couleur de la lettre
Dim RSz As Long
Dim VSz As Long
Dim BSz As Long
'
'temps de transition 3 (opacité de la lettre s'estompe)
Dim TransTemps3 As Long
'temps écoulé
Dim TempsSz3 As Long
'
'temps de transition 4 (changement de lettre)
Dim TransTemps4 As Long
'temps écoulé
Dim TempsSz4 As Long
'
'la lettre peut elle changer dans le temps
Dim ChangementSz As Long
'
'variations de couleur
Dim VarCouleur As Long
'
'état de la lettre
Dim EtatSz As Long
'
'état de la modification
Dim EtatModSz As Long
'
'
'*********************************
'* PROPRIETES
'*********************************
'
'ETAT DE LA MODIFICATION
Public Property Let EtatModif(Valeur As Long)
    '
    EtatModSz = Valeur
    '
End Property
'
Public Property Get EtatModif() As Long
    '
    EtatModif = EtatModSz
    '
End Property
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
Public Property Let Hauteur(Valeur As Single)
    '
    HauteurSz = Valeur
    '
End Property
'
Public Property Get Hauteur() As Single
    '
    Hauteur = HauteurSz
    '
End Property
'
'LARGEUR
Public Property Let Largeur(Valeur As Single)
    '
    LargeurSz = Valeur
    '
End Property
'
Public Property Get Largeur() As Single
    '
    Largeur = LargeurSz
    '
End Property
'
'*********************************
'* FIN DES PROPRIETES
'*********************************
'
'procédure permettan d'afficher la fonte choisie (via une texture)
Private Sub AfficherFonte(NuméroFonte As Long, Optional CouleurSz As Long = vbWhite)
    '
    Dim CharX As Integer
    Dim CharY As Integer
    Dim cTmp As Long
    '
    'on calcule les coordonnées
    cTmp = (NuméroFonte - 1) \ 13
    CharY = 25 * cTmp
    cTmp = (NuméroFonte - 1) Mod 13
    CharX = 19 * cTmp
    '
    'on détermine la couleur
    CouleurSz = Hex2Long(Hex(OpaciteSz) & Hex(CouleurSz))
    '
    'on génère les points
    vertMatrix(0) = CreateTLVertex(XSz, YSz, 0, 1, CouleurSz, 0, (1 / 256) * CharX, (1 / 256) * CharY)
    vertMatrix(1) = CreateTLVertex(XSz + LargeurSz, YSz, 0, 1, CouleurSz, 0, ((1 / 256) * CharX) + ((1 / 256) * 19), (1 / 256) * CharY)
    vertMatrix(2) = CreateTLVertex(XSz, YSz + HauteurSz, 0, 1, CouleurSz, 0, (1 / 256) * CharX, ((1 / 256) * CharY) + ((1 / 256) * 25))
    vertMatrix(3) = CreateTLVertex(XSz + LargeurSz, YSz + HauteurSz, 0, 1, CouleurSz, 0, ((1 / 256) * CharX) + ((1 / 256) * 19), ((1 / 256) * CharY) + ((1 / 256) * 25))
    '
    'on dessine l'image suivant son état (lettre normale ou "bizarre")
    Select Case EtatSz
        '
        Case 1
            '
            D3DDevice.SetTexture 0, MatrixTex_Normal
            '
        Case 2
            '
            D3DDevice.SetTexture 0, MatrixTex_Vert
            '
        Case 3
            '
            D3DDevice.SetTexture 0, MatrixTex_Vert
            '
        Case 4
            '
            D3DDevice.SetTexture 0, MatrixTex_Normal
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
'procédure qui va s'occuper d'afficher la lettre dans le temps
Public Sub AfficherTemps()
    '
    'on affiche la lettre en fonction de son état
    Select Case EtatSz
        '
        Case 1 'lettre normale
            '
            'on opère une modification sur la lettre en fonction de l'état de modification
            Select Case EtatModSz
                '
                Case 1 'les dimensions de la lettre augmentes
                    '
                    XSz = XSz - 0.05
                    YSz = YSz - 0.05
                    '
                    HauteurSz = HauteurSz + 0.1
                    LargeurSz = LargeurSz + 0.1
                    '
                '
            End Select
            '
            AfficherFonte NumLettreSz, RGB(255, 255, 255)
            '
            Exit Sub
            '
        Case 2 'lettre "bizarre"
            '
            AfficherFonte NumLettreSz, RGB(255, 255, 255)
            '
            'si la lettre peut changer dans le temps, on la change
            If ChangementSz = 1 Then
                '
                'on attend que le temps nécessaire avant d'effectuer le changement
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
            Exit Sub
            '
        Case Else 'fade-off
            '
            'on effectue la transition
            If TempsSz3 < TransTemps3 Then
                '
                'ici, l'opacité s'estompe
                OpaciteSz = OpaciteSz - (OpaciteSz / (TransTemps3 - TempsSz3))
                '
                '-------------------------------------------------------------
                '- toutes sortes de modifications sur la lettre est possible -
                '-------------------------------------------------------------
                'si l'état est de 4, alors on change les coordonnées et les dimensions de la lettre
                If EtatSz = 4 Then
                    '
                    XSz = XSz + 0.1
                    YSz = YSz + 0.1
                    '
                    HauteurSz = HauteurSz - 0.2
                    LargeurSz = LargeurSz - 0.2
                    '
                End If
                '
                AfficherFonte NumLettreSz, RGB(RSz, VSz, BSz)
                '
                TempsSz3 = TempsSz3 + 1
                '
                Exit Sub
                '
            End If
            '
        '
    End Select
    '
    'si on arrive jusqu'ici, c'est que la "vie" de la lettre s'est écoulée
    EtatSz = 0
    '
End Sub

Private Sub Class_Initialize()
    '
    TransTemps0 = 1
    '
    EtatSz = 1
    EtatModSz = 0
    '
    OpaciteSz = 255
    '
    RSz = 255
    VSz = 255
    BSz = 255
    '
End Sub


