Sub Chiffre_Affaire()
'
' Chiffre_Affaire Macro
'

'
    Columns("J:J").Select
    Columns("K:K").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("K1").Select
    'Nom de la colonne
    ActiveCell.FormulaR1C1 = "Chiffre d'affaires"
    Range("K2").Select
    Application.CutCopyMode = False
    'Formule pour calculer le chiffre d'affaire par produit
    ActiveCell.FormulaR1C1 = "=RC[-2]*RC[-1]"
    Range("K2").Select
    Selection.AutoFill Destination:=Range("K2:K45")
    Range("K2:K45").Select
    ' Changement du format
    Selection.NumberFormat = "#,##0.00 $"
    ' Mettre le chiffre d'affaire en gras et en bleu
    With Selection.Font
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = -0
        .Bold = True
    End With
    ActiveWindow.SmallScroll Down:=18
    Range("K46").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-44]C:R[-1]C)"
    Range("K47").Select
End Sub
