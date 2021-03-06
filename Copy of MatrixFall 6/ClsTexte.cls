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
'classe de ligne n�4
Dim ClsLng4() As ClsLigne4
'
'temps courant �coul� pendant une pause
Dim TempsCourSz As Long
'
'texte courant
Dim TexteCourSz As Integer
'
'temps de pause entre l'affichage de chaque lettre
Dim PauseLettre As Long
'
'temps �coul� pour chaque lettre (durant la pause)
Dim TempsLettreSz As Long
'
'�tat de la classe
Dim EtatSz As Long
'
'nombre de lignes par phrase
Dim NbrLignes As Long
'
'index de la ligne courante
Dim LigneCourSz As Integer
'
'coordonn�e x de la ligne courante
Dim XTmp As Long
'
'coordonn�e x du d�but du texte
Dim XDeb As Long
'
'lettre � afficher
Dim LettreTmp As Long
'
'r�cup�re l'index de la nouvelle ligne cr��e
Dim ResSz As Integer
'
'simple variables d'incr�mentation
Dim i As Integer
'
'variable permettant de savoir s'il y a eu un texte avant afin de l'enlever
Dim TexteAvt As Boolean
'
'variable permettant de savoir s'il faut attendre que les lettres courantes s'effacent
Dim TexteAtt As Boolean
'
'cette liste sert � garder en m�moire soit les lignes qu'il reste � afficher, soit les lignes qu'il reste � effacer
'elle va pouvoir nous permettre d'afficher les lignes al�atoirement
Dim ListeLignes As New Collection
'
'nous permet de savoir si l'ensemble des textes a �t� affich�
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
'proc�dure d'initialisation
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
'proc�dure d'affichage des textes
Public Sub Afficher()
    '
    'si l'�tat de la classe = 0 alors cette classe n'a plus rien � faire, donc on quitte la proc�dure
    If EtatSz = 0 Then Exit Sub
    '
    'on affiche les lignes s'il y en a
    For i = 1 To UBound(ClsLng4)
        '
        'on affiche si la ligne n'est pas d�sactiv�e (�tat = 0)
        If ClsLng4(i).Etat <> 0 Then ClsLng4(i).Afficher
        '
    Next
    '
    'on attend que toutes les lettres se soient effac�es s'il le faut
    If TexteAtt = True Then
        '
        'on v�rifie chaque ligne � propos de leur �tat
        For i = 1 To UBound(ClsLng4)
            '
            'si une seule ligne est encore en train de s'effacer, on quite la proc�dure
            If ClsLng4(i).Etat <> 0 Then Exit Sub
            '
        Next
        '
        'si on arrive jusqu'ici, c'est que toutes les lignes se sont effac�es, donc on peut r�-initialiser le tableau contenant les lignes
        ReDim ClsLng4(0 To 0)
        '
        'il ne faut plus attendre
        TexteAtt = False
        '
        'on v�rifie si toutes les phrase ont �t� affich�es
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
    'on v�rifie que cet index ne d�passe pas la limite
    If TexteCourSz > UBound(TexteSz) Then
        '
        'oui, donc on change l'�tat de la classe et on stoppe la proc�dure ici
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
        'le temps s'est �coul�, on enl�ve les lignes pr�c�dentes s'il y en a
        If TexteAvt = True Then
            '
            'on v�rifie que le temps de pause se soit �coul� afin d'entamer la proc�dure de fade-out de la lettre courante
            If PauseLettre <= TempsLettreSz Then
                '
                TempsLettreSz = 0
                '
                'on calcule un autre temps de pause pour le suivant
                PauseLettre = RndNbr(1, 30)
                '
                LigneCourSz = LigneCourSz + 1
                '
                'on v�rifie que l'index de la ligne courante ne soit pas trop �lev�
                If LigneCourSz > UBound(ClsLng4) Then
                    '
                    'oui, donc on a fini avec cette phrase-ci
                    TexteAvt = False
                    '
                    'ceci signifie qu'il faut attendre que les lettres qui s'effacent avant d'entamer toute autre proc�dure
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
                'le temps ne s'est pas encore �coul�, on incr�mente notre variable
                TempsLettreSz = TempsLettreSz + 1
                '
            End If
            '
            'on sort de la proc�dure
            Exit Sub
            '
        End If
        '
        'si c'est la premi�re fois qu'on s'occupe de cette phrase, on r�cup�re quelques infos
        If NbrLignes = 0 Then
            '
            'le nombre de lettres
            NbrLignes = Len(TexteSz(TexteCourSz))
            '
            'on ajoute autant de ligne qu'il faut afficher � notre liste
            For i = 1 To NbrLignes
                '
                ListeLignes.Add i, "k" & i
                '
            Next
            '
            'la coordonn�e x qui va nous indiquer l� o� le d�but de la phrase doit commencer
            XDeb = (DimL / 2) - ((Len(TexteSz(TexteCourSz)) * LargeurLettreSz) / 2)
            '
            'on d�termine les limites gauche et droite de notre classe s'occupant de l'affichage des lignes
            'With ClsPlusL
                '
                '.LimiteGauche = XDeb
                '.LimiteDroite = XDeb + (NbrLignes * (LargeurLettreSz))
                '
            'End With
            '
        End If
        '
        'on v�rifie que le temps de pause entre l'affichage de chaque lettre s'est �coul�
        If PauseLettre <= TempsLettreSz Then
            '
            TempsLettreSz = 0
            '
            'on calcule un autre temps de pause pour le suivant
            PauseLettre = RndNbr(1, 40)
            '
            'on v�rifie que la liste contient encore un index
            If ListeLignes.Count = 0 Then
                '
                'non, donc on a fini avec cette phrase-ci
                NbrLignes = 0
                TexteAvt = True
                LigneCourSz = 0
                TempsCourSz = 0
                XTmp = 0
                '
                'on incr�mente l'index du texte courant � afficher
                TexteCourSz = TexteCourSz + 1
                '
                'on v�rifie que l'index du texte n'est pas trop �lev�
                If TexteCourSz > UBound(TexteSz) Then
                    '
                    TexteCourSz = 0
                    '
                    'on a affich� toutes les phrase
                    OkFinSz = True
                    '
                End If
                '
                'on sort de la proc�dure
                Exit Sub
                '
            End If
            '
            'on r�cup�re l'index de la ligne � afficher
            LigneCourSz = ListeLignes.Item(RndNbr(1, ListeLignes.Count))
            '
            'on enl�ve cet index de la liste
            ListeLignes.Remove "k" & LigneCourSz
            '
            TexteAvt = False
            '
            'on v�rifie que le caract�re ne soit pas un espace, sinon on affiche rien
            If Mid(TexteSz(TexteCourSz), LigneCourSz, 1) <> " " Then
                '
                'on calcule la coordonn�e x de la lettre par rapport � son index
                XTmp = ((LigneCourSz - 1) * (LargeurLettreSz)) + XDeb
                '
                'on calcule la lettre par rapport � sa position dans le tableau des caract�res de vb
                LettreTmp = Asc(Mid(TexteSz(TexteCourSz), LigneCourSz, 1)) - 96
                '
                ResSz = ChargerLng(XTmp, 0, PosMilieuY, LettreTmp, HauteurLettreSz, LargeurLettreSz, 0, RndNbr(10, 25), 255, 150)
                '
            End If
            '
        Else
            '
            'le temps ne s'est pas encore �coul�, on incr�mente notre variable
            TempsLettreSz = TempsLettreSz + 1
            '
        End If
        '
    Else
        '
        'le temps ne s'est pas encore �coul�, on incr�mente notre variable
        TempsCourSz = TempsCourSz + 1
        '
    End If
    
    '
End Sub
'
'cette fontion charge une ligne n�4 et renvoie son index
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
        'point de d�part de la ligne (y)
        .YDep = YDepart
        '
        'point d'arriv�e de la lettre (y)
        .YArr = YArrivee
        '
        'vitesse des lettres
        .Vitesse = VitesseSz
        '
        'dur�e de la transition "fade-off"
        .Transition3 = Trans3
        '
        'opacit� des train�es
        .Trainee1Alpha = OpaciteTrainee1
        .Trainee2Alpha = OpaciteTrainee2
        '
    End With
    '
    ChargerLng = UBound(ClsLng4)
    '
End Function
