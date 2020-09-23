VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLigne2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************************************
'*                                                                                *
'* CETTE CLASSE GERE UNE SEULE LIGNE DE PLUSIEURS LETTRES (N°2)                   *
'*                                                                                *
'* Thomas John (thomas.john@swing.be)                                             *
'*                                                                                *
'**********************************************************************************
'
'classe gérant une lettre
Dim ClsL() As ClsLettre2
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
'temps que va mettre la couleur de la lettre à virer au vert
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
'nombre de lettres que cette ligne contient
Dim NbrLettresCont As Long
'
'position du pointeur destiné à débuter la procédure de noircissement des lettres
Dim PointeurSz As Integer
'
'nombre de lettre à créer par cycle
Dim NbrLettresCycle As Long
'
'hauteur de la trainée
Dim HTrainee As Long
'
'nous permet de savoir si on a fini de créer de nouvelles lettres
Dim OkFinSz As Boolean
'
'effacement aléatoire des lettres
Dim EffAleaSz As Long
'
'
'*********************************
'* PROPRIETES
'*********************************
'
'EFFACEMENT ALEATOIRE DES LETTRES
Public Property Let EffacementAlea(Valeur As Long)
    '
    EffAleaSz = Valeur
    '
End Property
'
Public Property Get EffacementAlea() As Long
    '
    EffacementAlea = EffAleaSz
    '
End Property
'
'NOMBRE DE LETTRES A CREER PAR CYCLE
Public Property Let NombreLettresCycle(Valeur As Long)
    '
    NbrLettresCycle = Valeur
    '
End Property
'
Public Property Get NombreLettresCycle() As Long
    '
    NombreLettresCycle = NbrLettresCycle
    '
End Property
'
'NOMBRE DE LETTRES QUE CETTE LIGNE DOIT CONTENIR
Public Property Let NombreLettres(Valeur As Long)
    '
    NbrLettresCont = Valeur
    '
End Property
'
Public Property Get NombreLettres() As Long
    '
    NombreLettres = NbrLettresCont
    '
End Property
'
'FLOU DES LETTRES
Public Property Let Flou(Valeur As Long)
    '
    If Valeur = 1 Then
        '
        FlouSz = 26
        '
    Else
        '
        FlouSz = 0
        '
    End If
    '
End Property
'
Public Property Get Flou() As Long
    '
    Flou = FlouSz
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
    'trainée blanche
    Set ClsT(1) = New ClsTrainee
    '
    'trainée verte
    Set ClsT(2) = New ClsTrainee
    '
    With ClsT(1)
        '
        .OpaciteMax = 120
        .Trainée = RndNbr(1, 3)
        .Etat = 1
        .Hauteur = 134
        .Largeur = 19
        '
    End With
    '
    With ClsT(2)
        '
        .OpaciteMax = 120
        .Trainée = RndNbr(5, 8)
        .Etat = 2
        .Hauteur = 134
        .Largeur = 19
        '
    End With
    '
    EtatSz = 1
    NbrLettresCont = 1
    NbrLettresCycle = 1
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
        'on vérifie si il faut encore créer de nouvelles lettres
        If OkFinSz = True Then
            '
            'on augmente toujours la variable "NbrLettres"
            NbrLettres = NbrLettres + 1
            '
            GoTo AFFICHAGE_LETTRE
            '
        End If
        '
        'on crée autant de nouvelles lettres qu'il en est demandé par cycle
        For i = 1 To NbrLettresCycle
            '
            NbrLettres = NbrLettres + 1
            'YSz = (NbrLettres * HauteurSz) - HauteurSz + 20
            YSz = (NbrLettres * (HauteurSz - 3)) - HauteurSz
            '
            'on vérifie que la coordonné y ne dépasse pas la taille de l'affichage choisi
            If YSz < DimH Then
                '
                ReDim Preserve ClsL(0 To NbrLettres)
                '
                Set ClsL(NbrLettres) = New ClsLettre2
                '
                With ClsL(NbrLettres)
                    '
                    .Opacite = 255
                    .Hauteur = HauteurSz
                    .Largeur = LargeurSz
                    .X = XSz
                    .Y = YSz
                    .Transition3 = TransitionSz3
                    .Changement = RndNbr(1, 4)
                    .Transition4 = RndNbr(4, 10)
                    .Lettre = RndNbr(1, 57)
                    '
                    Select Case RndNbr(1, 3)
                        '
                        Case 1
                            '
                            .Variation = 0
                            '
                        Case 2
                            '
                            .Variation = -50
                            '
                        Case 3
                            '
                            .Variation = -100
                            '
                        '
                    End Select
                    '
                    .Etat = 1
                    '
                End With
                '
                'on change l'état de la lettre précédente si le nombre de lettres est supérieur à 1
                If NbrLettres > 1 Then ClsL(NbrLettres - 1).Etat = 2
                '
            Else
                '
                OkFinSz = True
                '
                GoTo AFFICHAGE_LETTRE
                '
            End If
            '
        Next
        '
AFFICHAGE_LETTRE:
        '
        'on commence le noircissement des lettres lorsque le nombre demandé a été atteint
        If NbrLettres > NbrLettresCont Then
            '
            For i = 1 To NbrLettresCycle
                '
                PointeurSz = PointeurSz + 1
                '
                'on vérifie que le pointeur ne dépasse pas le nombre de lettres
                If PointeurSz <= UBound(ClsL) Then
                    '
                    With ClsL(PointeurSz)
                        '
                        If .Etat <> 0 Then .Etat = 3
                        '
                    End With
                    '
                Else
                    '
                    Exit For
                    '
                End If
                '
            Next
            '
        End If
        '
        'on change la hauteur de la trainee
        HTrainee = HTrainee + NbrLettresCycle
    Else
        '
        VitesseCnt = VitesseCnt + 1
        '
    End If
    '
    'si il n'y a aucune lettre à afficher, on affiche rien
    If UBound(ClsL) = 0 Then Exit Sub
    '
    'on affiche les lettres
    For i = 1 To UBound(ClsL)
        '
        With ClsL(i)
            '
            'si l'état de la lettre = 0 alors on ne l'affiche pas
            If .Etat <> 0 Then .AfficherTemps
            '
            'si l'effacement aléatoire des lettres est activé, on opére...
            If EffAleaSz = 1 Then
                '
                If RndNbr(1, 10) = 1 Then
                    '
                    .Etat = 3
                    '
                    .Transition3 = TransitionSz3 * 3
                    '
                End If
                '
            End If
            '
        End With
        '
    Next
    '
    'on affiche les trainées
    'on affiche la trainée de derrière que si l'effacement aléatoire n'est pas activé
    If EffAleaSz <> 1 Then
        '
        If PointeurSz > 0 Then
            '
            With ClsT(2)
                '
                .Largeur = LargeurSz
                .X = XSz
                .Y = ((PointeurSz * 1.5) * (HauteurSz - 3)) - ClsT(2).Hauteur + HauteurSz
                .AfficherTemps
                '
            End With
            '
        End If
        '
    End If
    '
    With ClsT(1)
        '
        .Largeur = LargeurSz
        .X = XSz
        .Y = ((HTrainee) * (HauteurSz - 3)) - ClsT(1).Hauteur
        .AfficherTemps
        '
    End With
    '
    'on vérifie si l'état de la dernière lettre = 0
    If ClsL(UBound(ClsL)).Etat = 0 Then
        '
        'oui, donc on change d'état
        EtatSz = 0
        ReDim ClsL(0 To 0)
        OkFinSz = False
        PointeurSz = 0
        NbrLettres = 0
        YSz = 0
        '
        'on enlève cette ligne de la liste
        ListeCooX.Remove "k" & XSz
        '
    End If
    '
End Sub
