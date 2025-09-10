Attribute VB_Name = "Clients_Facture"


Sub Final_Client_Facture()


Application.StatusBar = "Facture"


Sheets.Add.Name = "Facture"
    Cells.Select
    With Selection.Font
        .Name = "Times New Roman"
        .Size = 10
    End With
    With ActiveSheet.PageSetup
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.25)
        .BottomMargin = Application.InchesToPoints(0.25)
        .Order = xlOverThenDown
    End With
    Call Fiche_Client_Livraison
    Call Fiche_Client_Facture

Application.StatusBar = False
    
    
End Sub


Sub Fiche_Client_Facture()


    Call Mise_en_page_Client_Facture
Range("F11:I11").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlLeft
    'ActiveCell.FormulaR1C1 = "titre"
Range("F12:J12").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlLeft
    'ActiveCell.FormulaR1C1 = "nom prénom"
Range("F13:I13").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlLeft
    'ActiveCell.FormulaR1C1 = "adresse 1"
Range("F14:I14").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlLeft
    'ActiveCell.FormulaR1C1 = "adresse 2"
Range("F15:I15").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlLeft
    'ActiveCell.FormulaR1C1 = "code postale"
Range("F16:I16").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlLeft
    'ActiveCell.FormulaR1C1 = "pays"
Range("B21:C21").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlRight
    ActiveCell.FormulaR1C1 = "Facture :"
Range("D21:E21").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlLeft
    'ActiveCell.FormulaR1C1 = "1234"
Range("F21").Select
    Selection.HorizontalAlignment = xlLeft
    ActiveCell.FormulaR1C1 = "Genève, le"
Range("G21:I21").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlLeft
    Selection.NumberFormat = "d mmmm yyyy"
Range("B22:C22").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlRight
    'ActiveCell.FormulaR1C1 = "Référence :"
Range("D22:E22").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlLeft
    'ActiveCell.FormulaR1C1 = "5678"
Range("B26").Select
    Selection.HorizontalAlignment = xlCenter
    ActiveCell.FormulaR1C1 = "Quantité"
Range("D26").Select
   Selection.HorizontalAlignment = xlCenter
    ActiveCell.FormulaR1C1 = "Article"
Range("H26").Select
    Selection.HorizontalAlignment = xlCenter
    ActiveCell.FormulaR1C1 = "Unité"
Range("J26").Select
    Selection.HorizontalAlignment = xlCenter
    ActiveCell.FormulaR1C1 = "Prix"
Range("D30").Select
    Selection.HorizontalAlignment = xlCenter
Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("B32").Select
    Selection.HorizontalAlignment = xlCenter
Range("H32").Select
    Selection.NumberFormat = "#,##0.00"
Range("J32").Select
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=RC[-8]*RC[-2]"
Range("B34").Select
    Selection.HorizontalAlignment = xlCenter
Range("H34").Select
    Selection.NumberFormat = "#,##0.00"
Range("J34").Select
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=RC[-8]*RC[-2]"
Range("B36").Select
    Selection.HorizontalAlignment = xlCenter
Range("H36").Select
    Selection.NumberFormat = "#,##0.00"
Range("J36").Select
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=RC[-8]*RC[-2]"
Range("B38").Select
    Selection.HorizontalAlignment = xlCenter
Range("H38").Select
    Selection.NumberFormat = "#,##0.00"
Range("J38").Select
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=RC[-8]*RC[-2]"
Range("B40").Select
    Selection.HorizontalAlignment = xlCenter
Range("H40").Select
    Selection.NumberFormat = "#,##0.00"
Range("J40").Select
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=RC[-8]*RC[-2]"
Range("B42").Select
    Selection.HorizontalAlignment = xlCenter
Range("H42").Select
    Selection.NumberFormat = "#,##0.00"
Range("J42").Select
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=RC[-8]*RC[-2]"
Range("B44").Select
    Selection.HorizontalAlignment = xlCenter
Range("H44").Select
    Selection.NumberFormat = "#,##0.00"
Range("J44").Select
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=RC[-8]*RC[-2]"
Range("B46").Select
    Selection.HorizontalAlignment = xlCenter
Range("H46").Select
    Selection.NumberFormat = "#,##0.00"
Range("J46").Select
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=RC[-8]*RC[-2]"
Range("B48").Select
    Selection.HorizontalAlignment = xlCenter
Range("H48").Select
    Selection.NumberFormat = "#,##0.00"
Range("J48").Select
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=RC[-8]*RC[-2]"
Range("B50").Select
    Selection.HorizontalAlignment = xlCenter
Range("H50").Select
    Selection.NumberFormat = "#,##0.00"
Range("J50").Select
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=RC[-8]*RC[-2]"
Range("J51").Select
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("F53").Select
    Selection.HorizontalAlignment = xlRight
    ActiveCell.FormulaR1C1 = "Total"
Range("J53").Select
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=SUM(R[-25]C:R[-3]C)"
Columns("K").Select
    ActiveWindow.SelectedSheets.VPageBreaks.Add Before:=ActiveCell
Rows("64").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell


End Sub


Sub Fiche_Client_Livraison()


    Call Mise_en_page_Client_Livraison
Range("P11:S11").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlLeft
    'ActiveCell.FormulaR1C1 = "titre"
Range("P12:T12").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlLeft
    'ActiveCell.FormulaR1C1 = "nom prénom"
Range("P13:S13").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlLeft
    'ActiveCell.FormulaR1C1 = "adresse 1"
Range("P14:S14").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlLeft
    'ActiveCell.FormulaR1C1 = "adresse 2"
Range("P15:S15").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlLeft
    'ActiveCell.FormulaR1C1 = "code postale"
Range("P16:S16").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlLeft
    'ActiveCell.FormulaR1C1 = "pays"
Range("P18:S18").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlLeft
    'ActiveCell.FormulaR1C1 = "téléphone"
Range("L21:M21").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlRight
    ActiveCell.FormulaR1C1 = "Livraison :"
Range("N21:O21").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlLeft
    'ActiveCell.FormulaR1C1 = "1234"
Range("P21").Select
    Selection.HorizontalAlignment = xlLeft
    ActiveCell.FormulaR1C1 = "Genève, le"
Range("Q21:S21").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlLeft
    Selection.NumberFormat = "d mmmm yyyy"
Range("L22:M22").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlRight
    'ActiveCell.FormulaR1C1 = "Référence :"
Range("N22:O22").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlLeft
    'ActiveCell.FormulaR1C1 = "5678"
Range("L26").Select
    Selection.HorizontalAlignment = xlCenter
    ActiveCell.FormulaR1C1 = "Quantité"
Range("N26").Select
   Selection.HorizontalAlignment = xlCenter
    ActiveCell.FormulaR1C1 = "Article"
Range("N30").Select
    Selection.HorizontalAlignment = xlCenter
    ActiveCell.FormulaR1C1 = "=RC[-10]"
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("L32").Select
    Selection.HorizontalAlignment = xlCenter
    ActiveCell.FormulaR1C1 = "=RC[-10]"
Range("N32").Select
    ActiveCell.FormulaR1C1 = "=RC[-10]"
Range("L34").Select
    Selection.HorizontalAlignment = xlCenter
    ActiveCell.FormulaR1C1 = "=RC[-10]"
Range("N34").Select
    ActiveCell.FormulaR1C1 = "=RC[-10]"
Range("L36").Select
    Selection.HorizontalAlignment = xlCenter
    ActiveCell.FormulaR1C1 = "=RC[-10]"
Range("N36").Select
    ActiveCell.FormulaR1C1 = "=RC[-10]"
Range("L38").Select
    Selection.HorizontalAlignment = xlCenter
    ActiveCell.FormulaR1C1 = "=RC[-10]"
Range("N38").Select
    ActiveCell.FormulaR1C1 = "=RC[-10]"
Range("L40").Select
    Selection.HorizontalAlignment = xlCenter
    ActiveCell.FormulaR1C1 = "=RC[-10]"
Range("N40").Select
    ActiveCell.FormulaR1C1 = "=RC[-10]"
Range("L42").Select
    Selection.HorizontalAlignment = xlCenter
    ActiveCell.FormulaR1C1 = "=RC[-10]"
Range("N42").Select
    ActiveCell.FormulaR1C1 = "=RC[-10]"
Range("L44").Select
    Selection.HorizontalAlignment = xlCenter
    ActiveCell.FormulaR1C1 = "=RC[-10]"
Range("N44").Select
    ActiveCell.FormulaR1C1 = "=RC[-10]"
Range("L46").Select
    Selection.HorizontalAlignment = xlCenter
    ActiveCell.FormulaR1C1 = "=RC[-10]"
Range("N46").Select
    ActiveCell.FormulaR1C1 = "=RC[-10]"
Range("L48").Select
    Selection.HorizontalAlignment = xlCenter
    ActiveCell.FormulaR1C1 = "=RC[-10]"
Range("N48").Select
    ActiveCell.FormulaR1C1 = "=RC[-10]"
Range("L50").Select
    Selection.HorizontalAlignment = xlCenter
    ActiveCell.FormulaR1C1 = "=RC[-10]"
Range("N50").Select
    ActiveCell.FormulaR1C1 = "=RC[-10]"
Range("N53").Select
    ActiveCell.FormulaR1C1 = "date et signature"
  Columns("U").Select
    ActiveWindow.SelectedSheets.VPageBreaks.Add Before:=ActiveCell
 Rows("64").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
   

End Sub


Sub Mise_en_page_Client_Facture()


Columns("A:A").ColumnWidth = 4
Columns("B:B").ColumnWidth = 8.83
Columns("C:C").ColumnWidth = 0.82
Columns("D:D").ColumnWidth = 33
Columns("E:E").ColumnWidth = 2
Columns("F:F").ColumnWidth = 7.5
Columns("G:G").ColumnWidth = 3
Columns("H:H").ColumnWidth = 9.5
Columns("I:I").ColumnWidth = 2.83
Columns("J:J").ColumnWidth = 9.33


End Sub


Sub Mise_en_page_Client_Livraison()


Columns("K:K").ColumnWidth = 4
Columns("L:L").ColumnWidth = 8.83
Columns("M:M").ColumnWidth = 0.82
Columns("N:N").ColumnWidth = 33
Columns("O:O").ColumnWidth = 2
Columns("P:P").ColumnWidth = 7.5
Columns("Q:Q").ColumnWidth = 3
Columns("R:R").ColumnWidth = 9.5
Columns("S:S").ColumnWidth = 2.83
Columns("T:T").ColumnWidth = 9.33


End Sub


