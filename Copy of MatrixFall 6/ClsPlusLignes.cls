VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsPlusLignes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************************************
'*                                                                                *
'* CETTE CLASSE GERE L'AFFICHAGE D'UNE OU DE PLUSIEURS LIGNES                     *
'*                                                                                *
'* Thomas John (thomas.john@swing.be)                                             *
'*                                                                                *
'**********************************************************************************
'
'classes représentant une ligne (chaque type de ligne a sa spécialité)
Private ClsL() As ClsLigne
Private ClsL2() As ClsLigne2
Private ClsL3() As ClsLigne3
'
'limite gauche
Dim LimG As Single
Dim LimD As Single
'
'nombre max de lignes
Dim MaxSz As Long
'
'nombre courant de lignes
Dim NbrCourLignes As Long
'
'pause entre chaque apparition d'une ou de plusieurs lignes
Dim PauseApp As Long
'
'temps écoulé durant la pause
Dim PauseTmp As Long
'
'variable destinée à stocker l'index de chaque nouvelle ligne créée
Dim iTmp As Integer
'
'simples variables
Dim i As Integer
Dim i2 As Integer
'
'coordonnée X
Dim XTmp As Long
'
'variable utile dans certain cas
Dim lTmp As Long
'
'nous permet de savoir si le nombre de lignes nécessaire à été créé
Dim OkFinSz As Boolean
'
'état de la classe
Dim EtatSz As Long
'
'nombre de lignes à créer en même temps
Dim NbrLignesMTps As Long
'
'
'*************************************************************************
'* PROPRIETES
'*************************************************************************
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
'NOMBRE MAX DE LIGNES
Public Property Let Max(Valeur As Long)
    '
    MaxSz = Valeur
    '
End Property
'
Public Property Get Max() As Long
    '
    Max = MaxSz
    '
End Property
'
'LIMITE GAUCHE
Public Property Let LimiteGauche(Valeur As Single)
    '
    LimG = Valeur
    '
End Property
'
Public Property Get LimiteGauche() As Single
    '
    LimiteGauche = LimG
    '
End Property
'
'LIMITE DROITE
Public Property Let LimiteDroite(Valeur As Single)
    '
    LimD = Valeur
    '
End Property
'
Public Property Get LimiteDroite() As Single
    '
    LimiteDroite = LimD
    '
End Property
'
'*************************************************************************
'* FIN DES PROPRIETES
'*************************************************************************
'
Private Sub Class_Initialize()
    '
    'on initialise quelques variables
    ReDim ClsL(0 To 0)
    ReDim ClsL2(0 To 0)
    ReDim ClsL3(0 To 0)
    '
    XTmp = -HauteurLettreSz
    OkFinSz = False
    EtatSz = 0
    NbrLignesMTps = 1
    '
End Sub
'
'AFFICHAGE DES LIGNES (C'EST ICI QUE TOUT SE PASSE)
Public Sub Afficher()
    '
    'on affiche rien si l'état de la classe = 0
    If EtatSz = 0 Then Exit Sub
    '
    'on vérifie si le temps alloué pour la pause s'est écoulé
    If PauseApp <= PauseTmp Then
        '
        PauseTmp = 0
        '
        'on vérifie si le nombre de lignes créées ne dépasse pas une certaine limite (afin d'éviter d'utiliser trop de mémoire car les instances de classe créées ne sont pas déchargées même si on ne les utilises plus)
        'dépend de la capacité de l'odinateur (peut facilement aller au-delà de 20000)
        'on peut aussi l'utiliser afin de revenir à l'affichage des textes après un certain temps (ce qui donne une boucle)
        If NbrCourLignes > LimiteLignesAffSz Then
            '
            'si, on attend que la liste contenant la position des lignes actives soit vide
            If ListeCooX.Count = 0 Then
                '
                'elle l'est, on décharge toutes les instances de lignes créées pour libérer la mémoire
                ReDim ClsL(0 To 0)
                ReDim ClsL2(0 To 0)
                ReDim ClsL3(0 To 0)
                '
                'on change notre état afin que cette classe n'affiche plus de ligne
                EtatSz = 0
                '
                'on ré-initialise la classe s'occupant de l'affichage des textes
                ClsTxt.Init
                '
                'on réinitialise le nombre de lignes créées
                NbrCourLignes = 0
                '
            End If
            '
            'on ne crée plus aucune ligne, n en fait qu'afficher celles restantes
            GoTo AFF_PROC
            '
        End If
        '
        'on détermine la durée de la prochaine pause
        PauseApp = RndNbr(1, 2)
        '
        'on vérifie s'il est nécessaire de créer une nouvelle ligne
        If ListeCooX.Count * LargeurLettreSz >= DimL Then
            '
            'non, on passe directement à la procédure d'affichage
            GoTo AFF_PROC
            '
        End If
        '
        'on crée autant de lignes qu'il le faut pour ce cycle
        For i2 = 1 To NbrLignesMTps
            '
            'on incrémente le nombre courant de lignes de 1
            NbrCourLignes = NbrCourLignes + 1
            '
POSITIONNEMENT_LIGNE:
            '
            'on positionne notre ligne aléatoirement sur la coordonnée x
            XTmp = RndNbr(0, DimL \ LargeurLettreSz) * LargeurLettreSz
            '
            'on vérifie que cette coordonnée x n'est pas déjà utilisée
            For i = 1 To ListeCooX.Count
                '
                If CLng(ListeCooX.Item(i)) = XTmp Then
                    '
                    'on vérifie quand même si on est pas au bout de la liste
                    If i = ListeCooX.Count Then GoTo AFF_PROC
                    '
                    GoTo POSITIONNEMENT_LIGNE
                    '
                End If
                '
            Next
            '
            Select Case RndNbr(1, 4)
                '
                Case 1
                    '
                    'on détermine la position y de départ de la ligne
                    lTmp = 0
                    If RndNbr(1, 10) = 1 Then lTmp = RndNbr(1, DimH \ HauteurLettreSz) * HauteurLettreSz
                    '
                    'on crée une nouvelle ligne n°3
                    iTmp = ChargerLigne3(XTmp, lTmp, HauteurLettreSz, LargeurLettreSz, 0, RndNbr(1, 10), 10, 10, 20, RndNbr(5, 70), 128, 128)
                    '
                Case 2
                    '
                    'on crée une nouvelle ligne n°2
                    iTmp = ChargerLigne2(XTmp, HauteurLettreSz, LargeurLettreSz, 0, RndNbr(6, 15), RndNbr(20, 60), RndNbr(1, 2), RndNbr(0, 4))
                    '
                Case 3
                    '
                    'on crée une nouvelle ligne n°1
                    iTmp = ChargerLigne(XTmp, HauteurLettreSz, LargeurLettreSz, 0, 8, RndNbr(30, 65), RndNbr(5, 25))
                    '
                Case 4
                    '
                    'on crée une nouvelle ligne n°1
                    iTmp = ChargerLigne(XTmp, HauteurLettreSz, LargeurLettreSz, 0, 8, RndNbr(30, 65), RndNbr(5, 25))
                    '
                '
            End Select
            '
            ListeCooX.Add XTmp, "k" & XTmp
            '
        Next
        '
        'on détermine le nouveau nombre de ligne à créer par cycle
        NbrLignesMTps = RndNbr(1, 10)
        '
    Else
        '
        PauseTmp = PauseTmp + 1
        '
    End If
    '
AFF_PROC:
    '
    'on affiche toutes les lignes n°3
    For i = 1 To UBound(ClsL3)
        '
        With ClsL3(i)
            '
            'on affiche la ligne que si son état est <> 0
            If .Etat <> 0 Then .Afficher
            '
        End With
        '
    Next
    '
    'on affiche toutes les lignes n°2
    For i = 1 To UBound(ClsL2)
        '
        With ClsL2(i)
            '
            'on affiche la ligne que si son état est <> 0
            If .Etat <> 0 Then .Afficher
            '
        End With
        '
    Next
    '
    'on affiche toutes les lignes n°1
    For i = 1 To UBound(ClsL)
        '
        With ClsL(i)
            '
            'on affiche la ligne que si son état est <> 0
            If .Etat <> 0 Then .Afficher
            '
        End With
        '
    Next
    '
End Sub
'
'cette fontion charge une ligne n°1 et renvoie son index
Private Function ChargerLigne(XSz As Long, DimHauteurSz As Long, DimLargeurSz As Long, VitesseSz As Long, Trans1 As Long, Trans2 As Long, Trans3 As Long) As Integer
    '
    ReDim Preserve ClsL(LBound(ClsL) To UBound(ClsL) + 1)
    '
    Set ClsL(UBound(ClsL)) = New ClsLigne
    '
    With ClsL(UBound(ClsL))
        '
        'dimensions des lettres
        .Hauteur = DimHauteurSz 'RndNbr(10, 25)
        .Largeur = DimLargeurSz '.Hauteur * 0.65
        '
        'position de la ligne
        .X = XSz
        '
        'vitesse des lettres
        .Vitesse = VitesseSz
        '
        'durée de la transition vers le vert
        '.Transition1 = Trans1
        '
        'durée de la transition 2
        .Transition2 = Trans2
        '
        'durée de la transition vers le noir
        .Transition3 = Trans3
        '
    End With
    '
    ChargerLigne = UBound(ClsL)
    '
End Function
'
'cette fontion charge une ligne n°2 et renvoie son index
Private Function ChargerLigne2(XSz As Long, DimHauteurSz As Long, DimLargeurSz As Long, VitesseSz As Long, Trans3 As Long, NbrLettreAvt As Long, NbrLettresCycleSz As Long, EffacementAleatoire As Long) As Integer
    '
    ReDim Preserve ClsL2(LBound(ClsL2) To UBound(ClsL2) + 1)
    '
    Set ClsL2(UBound(ClsL2)) = New ClsLigne2
    '
    With ClsL2(UBound(ClsL2))
        '
        'dimensions des lettres
        .Hauteur = DimHauteurSz
        .Largeur = DimLargeurSz
        '
        'position de la ligne
        .X = XSz
        '
        'vitesse des lettres
        .Vitesse = VitesseSz
        '
        'durée de la transition vers le noir
        .Transition3 = Trans3
        '
        'on détermine le nombre de lettre que la ligne doit contenir avant de commencer à obscurcir les dernières lettres
        .NombreLettres = NbrLettreAvt
        '
        'on détermine le nombre de lettres à créer par cycle
        .NombreLettresCycle = NbrLettresCycleSz
        '
        .EffacementAlea = EffacementAleatoire
        '
    End With
    '
    ChargerLigne2 = UBound(ClsL2)
    '
End Function
'
'cette fontion charge une ligne n°3 et renvoie son index
Private Function ChargerLigne3(XSz As Long, YDepart As Long, DimHauteurSz As Long, DimLargeurSz As Long, VitesseSz As Long, Trans0 As Long, Trans1 As Long, Trans2 As Long, Trans3 As Long, DureeVie As Long, OpaciteTrainee1 As Long, OpaciteTrainee2 As Long) As Integer
    '
    ReDim Preserve ClsL3(LBound(ClsL3) To UBound(ClsL3) + 1)
    '
    Set ClsL3(UBound(ClsL3)) = New ClsLigne3
    '
    With ClsL3(UBound(ClsL3))
        '
        'dimensions des lettres
        .Hauteur = DimHauteurSz 'RndNbr(10, 25)
        .Largeur = DimLargeurSz '.Hauteur * 0.65
        '
        'position de la ligne
        .X = XSz
        .YDep = YDepart
        '
        'vitesse des lettres
        .Vitesse = VitesseSz
        '
        'durée de vie de la ligne
        .Vie = DureeVie
        '
        'durée de la transition (opacité 0 --> max)
        .Transition0 = Trans0
        '
        'durée de la transition vers le vert
        .Transition1 = Trans1
        '
        'durée de la transition 2
        .Transition2 = Trans2
        '
        'durée de la transition vers le noir
        .Transition3 = Trans3
        '
        'opacité des trainées
        .Trainee1Alpha = OpaciteTrainee1
        .Trainee2Alpha = OpaciteTrainee2
        '
    End With
    '
    ChargerLigne3 = UBound(ClsL3)
    '
End Function

Private Sub Class_Terminate()
    '
    'on libère de la mémoire (pas nécessaire, mais toujours utile pour être sûr)
    ReDim ClsL(0 To 0)
    ReDim ClsL2(0 To 0)
    ReDim ClsL3(0 To 0)
    '
End Sub
