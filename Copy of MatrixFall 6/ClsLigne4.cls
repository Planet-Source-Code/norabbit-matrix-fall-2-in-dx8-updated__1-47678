VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLigne4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************************************
'*                                                                                *
'* CETTE CLASSE GERE UNE SEULE LIGNE DE PLUSIEURS LETTRES (n°4)                   *
'*                                                                                *
'* Thomas John (thomas.john@swing.be)                                             *
'*                                                                                *
'**********************************************************************************
'
'classe gérant une lettre
Dim ClsL() As ClsLettre4
'
'classe gérant une trainée
Dim ClsT() As ClsTrainee
'
'lettre choisie (1-26)
Dim LettreChoix As Long
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
Dim YSzArr As Single 'arrivée de la ligne (dernière lettre)
'
'dimensions
Dim HauteurSz As Long
Dim LargeurSz As Long
'
'vitesse
Dim VitesseSz As Long
Dim VitesseCnt As Long
'
'temps que va mettre la lettre à s'estomper
Dim TransitionSz3 As Long
'
'etat de la ligne
Dim EtatSz As Single '1 = active ; 0 = inactive
'
'hauteur de la trainée
Dim HTrainee As Long
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
'LETTRE PRINCIPALE
Public Property Let Lettre(Valeur As Long)
    '
    LettreChoix = Valeur
    '
End Property
'
Public Property Get Lettre() As Long
    '
    Lettre = LettreChoix
    '
End Property
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
'ETAT DE LA LIGNE
Public Property Get Etat() As Single
    '
    Etat = EtatSz
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
'COORDONNEE Y D'ARRIVEE
Public Property Let YArr(Valeur As Single)
    '
    YSzArr = Valeur
    '
End Property
'
Public Property Get YArr() As Single
    '
    YArr = YSzArr
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
    With ClsT(1)
        '
        .Transition3 = 20
        .Hauteur = 134
        .Trainée = 3
        .Etat = 1
        '
    End With
    '
    'trainee de derrière
    Set ClsT(2) = New ClsTrainee
    '
    With ClsT(2)
        '
        .Transition3 = 20
        .Hauteur = 134
        .Trainée = RndNbr(5, 8)
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
'DEMANDE DE DEMARRAGE DE LA PROCEDURE DE FADE-OUT DE LA LETTRE PRINCIPALE
Public Sub Fin()
    '
    ClsL(UBound(ClsL)).Etat = 4
    '
End Sub
'
'AFFICHAGE DES LETTRES LES UNES APRES LES AUTRES
Public Sub Afficher()
    '
    'on ne fait rien tant que le temps (vitesse) n'a pas été atteind
    If VitesseCnt = VitesseSz Then
        '
        VitesseCnt = 0
        '
        'si la procédure d'apparition de la lettre est terminée, on entamme la procédure d'attente jusqu'à ce qu'on "nous" dise d'entammer la procédure de fade-out de la lettre
        If OkFinSz = True Then
            '
            'on vérifie l'état de la lettre. Si il est de 0, alors on stoppe la procédure ici (la lettre de fin s'est "effacée" --> après fade-off)
            If ClsL(UBound(ClsL)).Etat = 0 Then
                '
                'on change l'état de la ligne
                EtatSz = 0
                '
                Exit Sub
                '
            End If
            '
            'on  affiche les lettres à afficher
            For i = 1 To UBound(ClsL)
                '
                'on affiche que si l'état est <> 0
                If ClsL(i).Etat <> 0 Then
                    '
                    ClsL(i).AfficherTemps
                    '
                End If
                '
            Next
            '
            'on affiche la trainée de derrière
            With ClsT(2)
                '
                .Y = ((PointeurSz) * (HauteurSz - 3)) - ClsT(2).Hauteur + YSzDep
                .AfficherTemps
                '
            End With
            '
            'on incrémente le pointeur afin de faire tomber la trainée de derrière
            PointeurSz = PointeurSz + 1
            'on stoppe la procédure ici
            Exit Sub
            '
        End If
        '
        NbrLettres = NbrLettres + 1
        YSz = (NbrLettres * (HauteurSz - 3)) - HauteurSz + YSzDep
        '
        'on vérifie que la coordonné y ne dépasse pas la coordonnée d'arrivée
        If YSz <= YSzArr Then
            '
            'on crée une nouvelle lettre
            ReDim Preserve ClsL(0 To NbrLettres)
            '
            Set ClsL(NbrLettres) = New ClsLettre4
            '
            With ClsL(NbrLettres)
                '
                .Hauteur = HauteurSz
                .Largeur = LargeurSz
                .X = XSz
                .Y = YSz
                .Transition3 = TransitionSz3
                .Changement = RndNbr(1, 3)
                .Transition4 = RndNbr(5, 15)
                .Variation = RndNbr(-100, 0)
                .Lettre = LettreChoix
                '
            End With
            '
            'on change l'état de l'avant-dernière lettre ainsi que quelques paramètres
            If NbrLettres > 1 Then
                '
                With ClsL(NbrLettres - 1)
                    '
                    .Etat = 2
                    .Lettre = RndNbr(1, 57)
                    .Changement = RndNbr(1, 3)
                    '
                End With
                '
            End If
            '
            'on augment le multiplicateur qui permet de calculer la hauteur de la trainee de devant
            HTrainee = HTrainee + 1
            '
        Else
            '
            'on ne va pas plus bas, on commence à faire diminuer l'opacité des premières lettres
            PointeurSz = PointeurSz + 1
            '
            'on "fade-out" la trainée de devant
            ClsT(1).Etat = 3
            '
            'on vérifie si le pointeur n'est pas trop grand
            If PointeurSz > UBound(ClsL) - 1 Then
                '
                'on informe de la fin de la procédure d'apparition de la lettre
                OkFinSz = True
                '
                'on "fade-out" la trainée de derrière
                ClsT(2).Etat = 3
                '
                'on change l'état de modification de la lettre principale
                'ClsL(UBound(ClsL)).EtatModif = 1
                '
                'on change la durée de la transition 3 (fade-out)
                ClsL(UBound(ClsL)).Transition3 = TransitionSz3 * 3
                '
                'on va directement à l'affichage
                GoTo AFFICHAGE_L
                '
            End If
            '
            'on "fade-off" la lettre pointée
            ClsL(PointeurSz).Etat = 3
            '
        End If
        '
    Else
        '
        VitesseCnt = VitesseCnt + 1
        '
    End If
    '
AFFICHAGE_L:
    '
    'on continue si le nombre de lettres à afficher est supérieur à 1
    If UBound(ClsL) = 0 Then Exit Sub
    '
    'on  affiche les lettres
    For i = 1 To UBound(ClsL)
        '
        'on affiche que si l'état est <> 0
        If ClsL(i).Etat <> 0 Then
            '
            ClsL(i).AfficherTemps
            '
        End If
        '
    Next
    '
    'on affiche la trainée de devant
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
    If PointeurSz > 0 Then
        '
        'on affiche la trainée de derrière
        With ClsT(2)
            '
            .OpaciteMax = Trainee2Op
            .Largeur = LargeurSz
            .X = XSz
            .Y = ((PointeurSz) * (HauteurSz - 3)) - ClsT(2).Hauteur + YSzDep
            .AfficherTemps
            '
        End With
        '
    End If
    '
End Sub


