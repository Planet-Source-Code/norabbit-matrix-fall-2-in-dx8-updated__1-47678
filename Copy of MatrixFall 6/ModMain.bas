Attribute VB_Name = "ModMain"
'c'est ici que tout commence
Sub Main()
    '
    'on vérifie qu'aucune instance de ce programme ne tourne
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
        'aperçu
        Case "/p"
            '
            'rien pour le moment
            '
        'mode plein écran
        Case "/s"
            '
            'on vérifie que certaines informations utiles sont présentes (grâce au chargement des options)
            If AccMatSoft <> "HAL" And AccMatSoft <> "REF" Then
                '
                'des informations précieuses sont manquantes, on lance le panneau de configuration
                FrmConfig.Show
                '
            Else
                '
                'on lance la feuille principale
                FrmMain.Show
                '
                'on lance la procédure principale
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
