Attribute VB_Name = "ModMain"
'c'est ici que tout commence
Sub Main()
    '
    'on v�rifie qu'aucune instance de ce programme ne tourne
    If App.PrevInstance Then Exit Sub
    '
    'on initialise certaines variables
    ReDim TexteSz(0 To 0)
    ReDim PauseTexte(0 To 0)
    '
    'on charge les options
    ChargerOptions
    '
    Select Case LCase(Left$(Command, 2))
        '
        'aper�u
        Case "/p"
            '
            'rien pour le moment
            '
        'mode plein �cran
        Case "/s"
            '
            'on v�rifie que certaines informations utiles sont pr�sentes (gr�ce au chargement des options)
            If AccMatSoft <> "HAL" And AccMatSoft <> "REF" Then
                '
                'des informations pr�cieuses sont manquantes, on lance le panneau de configuration
                FrmConfig.Show
                '
            Else
                '
                'on lance la feuille principale
                FrmMain.Show
                '
                'on lance la proc�dure principale
                FrmMain.LancerProcP
                '
            End If
            '
        'panneau de configuration
        Case Else
            '
            'on lance le panneau de configuration
            FrmConfig.Show
            '
        '
    End Select
    '
End Sub
