Option Explicit

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    ' Cette procédure s'exécute chaque fois qu'une nouvelle feuille est activée (cliquée).
    
    ' Vérifie si le mode de calcul est Manuel, car c'est nécessaire pour cette logique.
    If Application.Calculation = xlCalculationManual Then
        ' Force le recalcul des cellules DÉPENDANTES sur la feuille active (Sh).
        ' Ceci est plus rapide que de recalculer toute la feuille.
        Sh.Calculate
    End If
End Sub
