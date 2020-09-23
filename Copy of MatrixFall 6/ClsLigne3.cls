VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLigne3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************************************
'*                                                                                *
'* CETTE CLASSE GERE UNE SEULE LIGNE DE PLUSIEURS LETTRES (n°3)                   *
'*                                                                                *
'* Thomas John (thomas.john@swing.be)                                             *
'*                                                                                *
'**********************************************************************************
'
'classe gérant une lettre
Dim ClsL() As ClsLettre3
'
'classe gérant une trainée
Dim ClsT() As ClsTrainee
'
'nombre de lettres chargées
Dim NbrLettres As Long
'
'numéro de la lettre courante
Dim CourLettre As Integer
'
'coordonnées
Dim XSz As Single
Dim YSz As Single
Dim YSzDep As Single 'départ de la ligne
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
'temps que va mettre l'opacité à arriver à son max
Dim TransitionSz0 As Long
'
'temps que va mettre la couleur de la lettre à virer vers la nouvelle couleur
Dim TransitionSz1 As Long
'
'temps pendant lequel la couleur de la lettre ne changera pas
Dim TransitionSz2 As Long
'
'temps que va mettre la lettre à s'estomper
Dim TransitionSz3 As Long
'
'etat de la ligne
Dim EtatSz As Single '1 = active ; 0 = inactive
'
'flou des lettres
Dim FlouSz As Long
'
'hauteur de la trainée
Dim HTrainee As Long
'
'temps de vie maximum de la ligne
Dim TempsVieMax As Long
'
'temps de vie courant de la ligne
Dim TempsVieCour As Long
'
'pointeur de fin de ligne (ce pointeur lancera la dernière transition d'un lettre (opacité) lorsqu'il passera dessus)
Dim PointeurSz As Integer
'
'permet de savoir si les opérations de fin de ligne sont terminèes
Dim OkFinSz As Boolean
'
'opacite des trainées
Dim Trainee1Op As Long
Dim Trainee2Op As Long
'
'
'*********************************
'* PROPRIETES
'*********************************
'
'OPACITE DE LA DEUXIEMME TRAINEE
Public Property Let Trainee2Alpha(Valeur As Long)
    '
    Trainee2Op = Valeur
    '
End Property
'
Public Property Get Trainee2Alpha() As Long
    '
    Trainee2Alpha = Trainee2Op
    '
End Property
'
'OPACITE DE LA PREMIERE TRAINEE
Public Property Let Trainee1Alpha(Valeur As Long)
    '
    Trainee1Op = Valeur
    '
End Property
'
Public Property Get Trainee1Alpha() As Long
    '
    Trainee1Alpha = Trainee1Op
    '
End Property
'
'TEMPS DE VIE MAXIMUM DE LE LIGNE
Public Property Let Vie(Valeur As Long)
    '
    TempsVieMax = Valeur
    '
End Property
'
Public Property Get Vie() As Long
    '
    Vie = TempsVieMax
    '
End Property
'
'ETAT DE LA LIGNE
Public Property Get Etat() As Single
    '
    Etat = EtatSz
    '
End Property
'
'TRANSITION 0
Public Property Let Transition0(Valeur As Single)
    '
    TransitionSz0 = Valeur
    '
End Property
'
Public Property Get Transition0() As Single
    '
    Transition0 = TransitionSz0
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
'COORDONNEE Y DE DEPART
Public Property Let YDep(Valeur As Single)
    '
    YSzDep = Valeur
    '
End Property
'
Public Property Get YDep() As Single
    '
    YDep = YSzDep
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
    ReDim ClsT(0 To 2)
    '
    'trainee de devant
    Set ClsT(1) = New ClsTrainee
    '
    Trainee1Op = 255
    '
    With ClsT(1)
        '
        .Transition3 = 10
        .Hauteur = 134
        .Trainée = RndNbr(1, 3)
        .Etat = 1
        '
    End With
    '
    'trainee de derrière
    Set ClsT(2) = New ClsTrainee
    '
    Trainee2Op = 255
    '
    With ClsT(2)
        '
        .Transition3 = 10
        .Hauteur = 134
        .Trainée = RndNbr(4, 8)
        .Etat = 2
        '
    End With
    '
    NbrLettres = 0
    YSz = 0
    EtatSz = 1
    PointeurSz = 0
    OkFinSz = False
    '
End Sub
'
'AFFICHAGE DES LETTRES LES UNES APRES LES AUTRES
Public Sub Afficher()
    '
    'on ne fait rien si l'état = 0
    If EtatSz = 0 Then Exit Sub
    '
    'on ne crée pas de nouvelle lettre tant que le temps (VitesseSz) n'a pas été atteint
    If VitesseCnt = VitesseSz Then
        '
        VitesseCnt = 0
        '
        'si les opérations de fin de ligne sont terminées, on va directement à l'affichage des lettre et trainées
        If OkFinSz = True Then
            '
            'on incrémente toujours le pointeur
            PointeurSz = PointeurSz + 1
            '
            GoTo AFF_PROC
            '
        End If
        '
        'on incrément le temps de vie courant de la ligne
        TempsVieCour = TempsVieCour + 1
        '
        'on vérifie si la durée de vie de la ligne est expirée
        If TempsVieCour >= TempsVieMax Then
            '
            'on incrément le pointeur (designe le numéro de la lettre)
            PointeurSz = PointeurSz + 1
            '
            'on change l'état de la trainée de devant afin qu'elle disparaisse
            ClsT(1).Etat = 3
            '
            'on vérifie que le pointeur ne soit pas trop grand
            If PointeurSz <= UBound(ClsL) Then
                '
                'on change l'état de la dernière lettre créée afin qu'elle devienne verte
                ClsL(UBound(ClsL)).Etat = 2
                '
                With ClsL(PointeurSz)
                    '
                    'on change l'état de la lettre à "4" qui signifie qu'elle peut effectuer la dernière transition
                    .Etat = 4
                    '
                    'on change la durée de la deuxième transition
                    .Transition2 = 0
                    '
                End With
                '
            Else
                '
                'on change l'état de la trainée de derrière afin qu'elle disparaisse
                ClsT(2).Etat = 3
                '
                'on informe que les opérations de fin de ligne sont terminées
                OkFinSz = True
                '
            End If
            '
            'on ne doit plus créer de lettre donc on passe la procédure qui suit
            GoTo AFF_PROC
            '
        End If
        '
        NbrLettres = NbrLettres + 1
        YSz = (NbrLettres * (HauteurSz - 3)) - HauteurSz + YSzDep
        '
        'on vérifie que la coordonné y ne dépasse pas la taille de l'affichage choisi
        If YSz < DimH Then
            '
            'on crée une nouvelle lettre
            ReDim Preserve ClsL(0 To NbrLettres)
            '
            Set ClsL(NbrLettres) = New ClsLettre3
            '
            With ClsL(NbrLettres)
                '
                .Hauteur = HauteurSz
                .Largeur = LargeurSz
                .X = XSz
                .Y = YSz
                .Transition0 = TransitionSz0
                .Transition1 = TransitionSz1 'RndNbr(5, 24)
                .Transition2 = TransitionSz2 'RndNbr(50, 150)
                .Transition3 = TransitionSz3 'RndNbr(10, 35)
                .Changement = RndNbr(1, 3)
                .Transition4 = RndNbr(5, 15)
                .Variation = RndNbr(-100, 0)
                .Lettre = RndNbr(1, 57)
                'ClsT(1).Etat = .Etat
                '
            End With
            '
            'on change l'état de l'avant-dernière lettre afin qu'elle devienne verte
            If NbrLettres > 1 Then ClsL(NbrLettres - 1).Etat = 2
            '
        End If
        '
        'on augment le multiplicateur qui permet de calculer la hauteur de la trainee de devant
        HTrainee = HTrainee + 1
        '
    Else
        '
        VitesseCnt = VitesseCnt + 1
        '
    End If
    '
AFF_PROC:
    '
    'on continue si le nombre de lettres à afficher est supérieur à 1
    If UBound(ClsL) = 0 Then Exit Sub
    '
    'on  affiche les lettres
    For i = 1 To UBound(ClsL)
        '
        With ClsL(i)
            '
            'on affiche que si l'état est <> 0
            If .Etat <> 0 Then
                '
                .AfficherTemps
                '
            End If
            '
        End With
        '
    Next
    '
    'on affiche la trainée de devant si la "transition 0" < 4
    If TransitionSz0 < 4 Then
        '
        With ClsT(1)
            '
            .OpaciteMax = Trainee1Op
            .Largeur = LargeurSz
            .X = XSz
            .Y = ((HTrainee) * (HauteurSz - 3)) - ClsT(1).Hauteur + YSzDep
            .AfficherTemps
            '
        End With
        '
    End If
    '
    If PointeurSz > 0 Then
        '
        'on affiche la trainée de derrière
        With ClsT(2)
            '
            .OpaciteMax = Trainee2Op
            .Largeur = LargeurSz
            .X = XSz
            .Y = ((PointeurSz * 2) * (HauteurSz - 3)) - ClsT(2).Hauteur + YSzDep
            .AfficherTemps
            '
        End With
        '
    End If
    '
    'on vérifie encore une fois si la durée de vie de la ligne est expirée
    If TempsVieCour >= TempsVieMax Then
        '
        'si oui, on vérifie si toutes les lettres ont disparu
        For i = 1 To UBound(ClsL)
            '
            'si une seule lettre est encore visibe, on quitte la procédure
            If ClsL(i).Etat <> 0 Then Exit Sub
            '
        Next
        '
    Else
        '
        'sinon, on quitte la procédure
        Exit Sub
        '
    End If
    '
    'si on arrive jusque ici, c'est que toutes les lettres ont disparu, donc on réinitialise la ligne
    EtatSz = 0
    '
    ReDim ClsL(0 To 0)
    '
    With ClsT(1)
        '
        .Init
        .Etat = 1
        '
    End With
    '
    With ClsT(2)
        '
        .Init
        .Etat = 2
        '
    End With
    '
    OkFinSz = False
    PointeurSz = 0
    TempsVieCour = 0
    NbrLettres = 0
    YSz = 0
    HTrainee = 0
    '
    'on enlève cette ligne de la liste
    ListeCooX.Remove "k" & XSz
    '
    '
End Sub
