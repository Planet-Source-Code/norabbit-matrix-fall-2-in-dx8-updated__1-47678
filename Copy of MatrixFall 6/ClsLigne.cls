VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLigne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************************************
'*                                                                                *
'* CETTE CLASSE GERE UNE SEULE LIGNE DE PLUSIEURS LETTRES                         *
'*                                                                                *
'* Thomas John (thomas.john@swing.be)                                             *
'*                                                                                *
'**********************************************************************************
'
'classe g�rant une lettre
Dim ClsL() As ClsLettre
'
'nombre de lettres charg�es
Dim NbrLettres As Long
'
'num�ro de la lettre courante
Dim CourLettre As Integer
'
'coordonn�es
Dim XSz As Single
Dim YSz As Single
'
'dimensions
Dim HauteurSz As Long
Dim LargeurSz As Long
'
Dim i As Integer
'
'vitesse
Dim VitesseSz As Long
Dim VitesseCnt As Long
'
'temps que va mettre la couleur de la lettre � virer au vert
Dim TransitionSz1 As Long
'
'temps pendant lequel la couleur de la lettre ne changera pas
Dim TransitionSz2 As Long
'
'temps que va mettre la lettre � s'estomper
Dim TransitionSz3 As Long
'
'etat de la ligne
Dim EtatSz As Single '1 = active ; 0 = inactive
'
'nous pr�vient si la cr�ation de lettres est termin�e
Dim OkFinSz As Boolean
'
'
'*********************************
'* PROPRIETES
'*********************************
'
'ETAT DE LA LIGNE
Public Property Get Etat() As Single
    '
    Etat = EtatSz
    '
End Property
'
'TRANSITION 1
Public Property Let Transition1(Valeur As Single)
    '
    TransitionSz1 = Valeur
    '
End Property
'
Public Property Get Transition1() As Single
    '
    Transition1 = TransitionSz1
    '
End Property
'
'TRANSITION 2
Public Property Let Transition2(Valeur As Single)
    '
    TransitionSz2 = Valeur
    '
End Property
'
Public Property Get Transition2() As Single
    '
    Transition2 = TransitionSz2
    '
End Property
'
'TRANSITION 3
Public Property Let Transition3(Valeur As Single)
    '
    TransitionSz3 = Valeur
    '
End Property
'
Public Property Get Transition3() As Single
    '
    Transition3 = TransitionSz3
    '
End Property
'
'VITESSE
Public Property Let Vitesse(Valeur As Single)
    '
    VitesseSz = Valeur
    '
End Property
'
Public Property Get Vitesse() As Single
    '
    Vitesse = VitesseSz
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
'
Private Sub Class_Initialize()
    '
    'on initialise certaines variables
    ReDim ClsL(0 To 0)
    '
    NbrLettres = 0
    YSz = 0
    EtatSz = 1
    '
End Sub
'
'AFFICHAGE DES LETTRES LES UNES APRES LES AUTRES
Public Sub Afficher()
    '
    'si l'�tat = 0, alors on en a fini avec cette ligne
    If EtatSz = 0 Then Exit Sub
    '
    'on ne cr�e pas de nouvelle lettre tant que le temps (VitesseSz) n'a pas �t� atteint
    If VitesseCnt = VitesseSz Then
        '
        VitesseCnt = 0
        '
        'si la cr�ation de lettres est termin�e, on passe � l'affichage
        If OkFinSz = True Then GoTo AFF_PROC
        '
        NbrLettres = NbrLettres + 1
        'YSz = (NbrLettres * HauteurSz) - HauteurSz + 20
        YSz = (NbrLettres * (HauteurSz - 3)) - HauteurSz
        '
        'on v�rifie que la coordonn� y ne d�passe pas la taille de l'affichage choisi
        If YSz < DimH Then
            '
            'on cr�e une nouvelle lettre
            ReDim Preserve ClsL(0 To NbrLettres)
            '
            Set ClsL(NbrLettres) = New ClsLettre
            '
            With ClsL(NbrLettres)
                '
                .Hauteur = HauteurSz
                .Largeur = LargeurSz
                .X = XSz
                .Y = YSz
                .Opacite = 255
                .Transition0 = 1 'RndNbr(1, 5)
                .Transition1 = TransitionSz1 '8 'RndNbr(5, 24)
                .Transition2 = RndNbr(50, 150)
                .Transition3 = RndNbr(10, 35)
                .Changement = RndNbr(1, 3)
                .Transition4 = RndNbr(5, 15)
                .Variation = RndNbr(-100, 0)
                .Lettre = RndNbr(1, 57)
                '
            End With
            '
            'on change l'�tat de la lettre pr�c�dente s'il y en a une
            If NbrLettres > 1 Then ClsL(NbrLettres - 1).Etat = 2
            '
        Else
            '
            'on change l'�tat de la derni�re lettre cr��e
            ClsL(NbrLettres - 1).Etat = 2
            '
            'on informe de la fin de la cr�ation des lettres
            OkFinSz = True
            '
        End If
        '
    Else
        '
        VitesseCnt = VitesseCnt + 1
        '
    End If
    '
AFF_PROC:
    '
    'si aucune lettre n'a encore �t� cr��e, on quitte ici
    If UBound(ClsL) = 0 Then Exit Sub
    '
    'on affiche les lettres
    For i = 1 To UBound(ClsL)
        '
        With ClsL(i)
            '
            'si l'�tat de la lettre = 0, on ne l'affiche pas
            If .Etat <> 0 Then .AfficherTemps
            '
        End With
        '
    Next
    '
    'on v�rifie l'�tat de toutes les lettres
    For i = 1 To UBound(ClsL)
        '
        'si l'�tat d'une des lettres est diff�rent de 0 alors on quitte la proc�dure ici
        If ClsL(i).Etat <> 0 Then Exit Sub
        '
    Next
    '
    'si on est arriv� jusqu'ici, c'est que plus aucune lettre ne doit �tre affich�e, donc on initialise la ligne
    EtatSz = 0
    ReDim ClsL(0 To 0)
    '
    NbrLettres = 0
    YSz = 0
    OkFinSz = False
    '
    'on enl�ve cette ligne de la liste
    ListeCooX.Remove "k" & XSz
    '
End Sub
