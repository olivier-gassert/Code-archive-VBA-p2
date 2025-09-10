Attribute VB_Name = "Clients_Certificat"
Sub Final_Client_Certificat()


Application.StatusBar = "Certificat"

Sheets.Add.Name = "Certificat"
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
    Call Fiche_Client_Certificat
    

Application.StatusBar = False
    
    
End Sub



Sub Fiche_Client_Certificat()


    Call Mise_en_page_Client_Certificat
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
Range("D30:H30").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    ActiveCell.FormulaR1C1 = "Certificat Elisa Gassert"

Range("D34:H34").Select
    Selection.Merge
    'Selection.HorizontalAlignment = xlCenter
    ActiveCell.FormulaR1C1 = "Par le présent certificat nous attestons que tous les meubles ELISA GASSERT sont"
Range("D36:H36").Select
    Selection.Merge
    'Selection.HorizontalAlignment = xlCenter
    ActiveCell.FormulaR1C1 = "en hêtre massif. Ils sont fabriques en suisse et peint dans notre atelier à Genève."
 Range("D39:H39").Select
    Selection.Merge
    'Selection.HorizontalAlignment = xlCenter
    ActiveCell.FormulaR1C1 = "Nos meubles se caractérisent par leur robustesse et leur longévité."
Range("D41:H41").Select
    Selection.Merge
    'Selection.HorizontalAlignment = xlCenter
    ActiveCell.FormulaR1C1 = "La peinture assure une conservation durable des couleurs et des motifs."
Range("D43:H43").Select
    Selection.Merge
    'Selection.HorizontalAlignment = xlCenter
    ActiveCell.FormulaR1C1 = ""
Range("D44:H44").Select
    Selection.Merge
    'Selection.HorizontalAlignment = xlCenter
    ActiveCell.FormulaR1C1 = "Nous vous remercions d'avoir choisi les meubles ELISA GASSERT."


Columns("K").Select
    ActiveWindow.SelectedSheets.VPageBreaks.Add Before:=ActiveCell
Rows("64").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell


End Sub


Sub Mise_en_page_Client_Certificat()


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
