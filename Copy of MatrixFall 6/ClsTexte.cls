VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsTexte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************************************
'*                                                                                *
'* CETTE CLASSE GERE L'AFFICHAGE D'UN OU DE PLUSIEURS TEXTES                      *
'*                                                                                *
'* Thomas John (thomas.john@swing.be)                                             *
'*                                                                                *
'**********************************************************************************
'
'classe de ligne n°4
Dim ClsLng4() As ClsLigne4
'
'temps courant écoulé pendant une pause
Dim TempsCourSz As Long
'
'texte courant
Dim TexteCourSz As Integer
'
'temps de pause entre l'affichage de chaque lettre
Dim PauseLettre As Long
'
'temps écoulé pour chaque lettre (durant la pause)
Dim TempsLettreSz As Long
'
'état de la classe
Dim EtatSz As Long
'
'nombre de lignes par phrase
Dim NbrLignes As Long
'
'index de la ligne courante
Dim LigneCourSz As Integer
'
'coordonnée x de la ligne courante
Dim XTmp As Long
'
'coordonnée x du début du texte
Dim XDeb As Long
'
'lettre à afficher
Dim LettreTmp As Long
'
'récupère l'index de la nouvelle ligne créée
Dim ResSz As Integer
'
'simple variables d'incrémentation
Dim i As Integer
'
'variable permettant de savoir s'il y a eu un texte avant afin de l'enlever
Dim TexteAvt As Boolean
'
'variable permettant de savoir s'il faut attendre que les lettres courantes s'effacent
Dim TexteAtt As Boolean
'
'cette liste sert à garder en mémoire soit les lignes qu'il reste à afficher, soit les lignes qu'il reste à effacer
'elle va pouvoir nous permettre d'afficher les lignes aléatoirement
Dim ListeLignes As New Collection
'
'nous permet de savoir si l'ensemble des textes a été affiché
Dim OkFinSz As Boolean
'
'position y des lettres
Dim PosMilieuY As Long
'
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
Private Sub Class_Initialize()
    '
    'on initialise le tableau
    ReDim ClsLng4(0 To 0)
    '
    'on initialise quelques variables
    TempsCourSz = 0
    EtatSz = 1
    TexteCourSz = 0
    TempsLettreSz = 0
    XTmp = 0
    TexteAvt = False
    TexteAtt = False
    OkFinSz = False
    '
    'on calcule la position y des lettres
    PosMilieuY = (DimH \ 2) - (HauteurLettreSz \ 2)
    '
End Sub
'
'procédure d'initialisation
Public Sub Init()
    '
    TempsCourSz = 0
    EtatSz = 1
    TexteCourSz = 0
    TempsLettreSz = 0
    XTmp = 0
    TexteAvt = False
    TexteAtt = False
    OkFinSz = False
    '
End Sub
'
'procédure d'affichage des textes
Public Sub Afficher()
    '
    'si l'état de la classe = 0 alors cette classe n'a plus rien à faire, donc on quitte la procédure
    If EtatSz = 0 Then Exit Sub
    '
    'on affiche les lignes s'il y en a
    For i = 1 To UBound(ClsLng4)
        '
        'on affiche si la ligne n'est pas désactivée (état = 0)
        If ClsLng4(i).Etat <> 0 Then ClsLng4(i).Afficher
        '
    Next
    '
    'on attend que toutes les lettres se soient effacées s'il le faut
    If TexteAtt = True Then
        '
        'on vérifie chaque ligne à propos de leur état
        For i = 1 To UBound(ClsLng4)
            '
            'si une seule ligne est encore en train de s'effacer, on quite la procédure
            If ClsLng4(i).Etat <> 0 Then Exit Sub
            '
        Next
        '
        'si on arrive jusqu'ici, c'est que toutes les lignes se sont effacées, donc on peut ré-initialiser le tableau contenant les lignes
        ReDim ClsLng4(0 To 0)
        '
        'il ne faut plus attendre
        TexteAtt = False
        '
        'on vérifie si toutes les phrase ont été affichées
        If OkFinSz = True Then
            '
            'oui donc on stoppe ici
            EtatSz = 0
            '
            'on active la classe qui va s'occuper de l'affichage des lignes
            ClsPlusL.Etat = 1
            '
            Exit Sub
            '
        End If
        '
    End If
    '
    'on vérifie que cet index ne dépasse pas la limite
    If TexteCourSz > UBound(TexteSz) Then
        '
        'oui, donc on change l'état de la classe et on stoppe la procédure ici
        'EtatSz = 0
        '
        TexteCourSz = 0
        '
        'Exit Sub
        '
    End If
    '
    'on attend le temps de la pause sur laquelle pointe "TexteCourSz"
    If PauseTexte(TexteCourSz) <= TempsCourSz Then
        '
        'le temps s'est écoulé, on enlève les lignes précédentes s'il y en a
        If TexteAvt = True Then
            '
            'on vérifie que le temps de pause se soit écoulé afin d'entamer la procédure de fade-out de la lettre courante
            If PauseLettre <= TempsLettreSz Then
                '
                TempsLettreSz = 0
                '
                'on calcule un autre temps de pause pour le suivant
                PauseLettre = RndNbr(1, 30)
                '
                LigneCourSz = LigneCourSz + 1
                '
                'on vérifie que l'index de la ligne courante ne soit pas trop élevé
                If LigneCourSz > UBound(ClsLng4) Then
                    '
                    'oui, donc on a fini avec cette phrase-ci
                    TexteAvt = False
                    '
                    'ceci signifie qu'il faut attendre que les lettres qui s'effacent avant d'entamer toute autre procédure
                    TexteAtt = True
                    '
                    LigneCourSz = 0
                    '
                Else
                    '
                    ClsLng4(LigneCourSz).Fin
                    '
                End If
                '
            Else
                '
                'le temps ne s'est pas encore écoulé, on incrémente notre variable
                TempsLettreSz = TempsLettreSz + 1
                '
            End If
            '
            'on sort de la procédure
            Exit Sub
            '
        End If
        '
        'si c'est la première fois qu'on s'occupe de cette phrase, on récupère quelques infos
        If NbrLignes = 0 Then
            '
            'le nombre de lettres
            NbrLignes = Len(TexteSz(TexteCourSz))
            '
            'on ajoute autant de ligne qu'il faut afficher à notre liste
            For i = 1 To NbrLignes
                '
                ListeLignes.Add i, "k" & i
                '
            Next
            '
            'la coordonnée x qui va nous indiquer là où le début de la phrase doit commencer
            XDeb = (DimL / 2) - ((Len(TexteSz(TexteCourSz)) * LargeurLettreSz) / 2)
            '
            'on détermine les limites gauche et droite de notre classe s'occupant de l'affichage des lignes
            'With ClsPlusL
                '
                '.LimiteGauche = XDeb
                '.LimiteDroite = XDeb + (NbrLignes * (LargeurLettreSz))
                '
            'End With
            '
        End If
        '
        'on vérifie que le temps de pause entre l'affichage de chaque lettre s'est écoulé
        If PauseLettre <= TempsLettreSz Then
            '
            TempsLettreSz = 0
            '
            'on calcule un autre temps de pause pour le suivant
            PauseLettre = RndNbr(1, 40)
            '
            'on vérifie que la liste contient encore un index
            If ListeLignes.Count = 0 Then
                '
                'non, donc on a fini avec cette phrase-ci
                NbrLignes = 0
                TexteAvt = True
                LigneCourSz = 0
                TempsCourSz = 0
                XTmp = 0
                '
                'on incrémente l'index du texte courant à afficher
                TexteCourSz = TexteCourSz + 1
                '
                'on vérifie que l'index du texte n'est pas trop élevé
                If TexteCourSz > UBound(TexteSz) Then
                    '
                    TexteCourSz = 0
                    '
                    'on a affiché toutes les phrase
                    OkFinSz = True
                    '
                End If
                '
                'on sort de la procédure
                Exit Sub
                '
            End If
            '
            'on récupère l'index de la ligne à afficher
            LigneCourSz = ListeLignes.Item(RndNbr(1, ListeLignes.Count))
            '
            'on enlève cet index de la liste
            ListeLignes.Remove "k" & LigneCourSz
            '
            TexteAvt = False
            '
            'on vérifie que le caractère ne soit pas un espace, sinon on affiche rien
            If Mid(TexteSz(TexteCourSz), LigneCourSz, 1) <> " " Then
                '
                'on calcule la coordonnée x de la lettre par rapport à son index
                XTmp = ((LigneCourSz - 1) * (LargeurLettreSz)) + XDeb
                '
                'on calcule la lettre par rapport à sa position dans le tableau des caractères de vb
                LettreTmp = Asc(Mid(TexteSz(TexteCourSz), LigneCourSz, 1)) - 96
                '
                ResSz = ChargerLng(XTmp, 0, PosMilieuY, LettreTmp, HauteurLettreSz, LargeurLettreSz, 0, RndNbr(10, 25), 255, 150)
                '
            End If
            '
        Else
            '
            'le temps ne s'est pas encore écoulé, on incrémente notre variable
            TempsLettreSz = TempsLettreSz + 1
            '
        End If
        '
    Else
        '
        'le temps ne s'est pas encore écoulé, on incrémente notre variable
        TempsCourSz = TempsCourSz + 1
        '
    End If
    
    '
End Sub
'
'cette fontion charge une ligne n°4 et renvoie son index
Private Function ChargerLng(XSz As Long, YDepart As Long, YArrivee As Long, LettreSz As Long, DimHauteurSz As Long, DimLargeurSz As Long, VitesseSz As Long, Trans3 As Long, OpaciteTrainee1 As Long, OpaciteTrainee2 As Long) As Integer
    '
    ReDim Preserve ClsLng4(LBound(ClsLng4) To UBound(ClsLng4) + 1)
    '
    Set ClsLng4(UBound(ClsLng4)) = New ClsLigne4
    '
    With ClsLng4(UBound(ClsLng4))
        '
        'lettre
        .Lettre = LettreSz
        '
        'dimensions des lettres
        .Hauteur = DimHauteurSz
        .Largeur = DimLargeurSz
        '
        'position de la ligne (x)
        .X = XSz
        '
        'point de départ de la ligne (y)
        .YDep = YDepart
        '
        'point d'arrivée de la lettre (y)
        .YArr = YArrivee
        '
        'vitesse des lettres
        .Vitesse = VitesseSz
        '
        'durée de la transition "fade-off"
        .Transition3 = Trans3
        '
        'opacité des trainées
        .Trainee1Alpha = OpaciteTrainee1
        .Trainee2Alpha = OpaciteTrainee2
        '
    End With
    '
    ChargerLng = UBound(ClsLng4)
    '
End Function
