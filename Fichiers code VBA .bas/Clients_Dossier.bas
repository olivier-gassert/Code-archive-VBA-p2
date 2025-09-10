Attribute VB_Name = "Clients_Dossier"

Sub Final_Client_Dossier()


Application.StatusBar = "Dossier"

Sheets.Add.Name = "Dossier"
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
    End With
    Call Fiche_Client_Dossier

Application.StatusBar = False
    
    
   End Sub
    
 
Sub Fiche_Client_Dossier()
    
    
    Call Mise_en_page_Client_Dossier
ActiveCell.Offset(1, 1).Range("A1:S1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Clients"
ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(0, 17).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(1, -16).Range("A1:Q1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    ActiveCell.FormulaR1C1 = "Dossier"
ActiveCell.Offset(2, -1).Range("A1:K1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Nom et Prénom"
ActiveCell.Offset(0, 2).Range("A1:G1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "No Référance"
ActiveCell.Offset(1, -12).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(0, 9).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(0, 5).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(1, -16).Range("A1:I2").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Font.Size = 20
ActiveCell.Offset(0, 4).Range("A1:E1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
Columns("V").Select
    ActiveWindow.SelectedSheets.VPageBreaks.Add Before:=ActiveCell
Rows("64").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell

End Sub


Sub Mise_en_page_Client_Dossier()


ActiveCell.Columns("A:A").EntireColumn.ColumnWidth = 2
ActiveCell.Offset(0, 1).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 2).Columns("A:A").EntireColumn.ColumnWidth = 22.14
ActiveCell.Offset(0, 3).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 4).Columns("A:A").EntireColumn.ColumnWidth = 1
ActiveCell.Offset(0, 5).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 6).Columns("A:A").EntireColumn.ColumnWidth = 10.71
ActiveCell.Offset(0, 7).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 8).Columns("A:A").EntireColumn.ColumnWidth = 1
ActiveCell.Offset(0, 9).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 10).Columns("A:A").EntireColumn.ColumnWidth = 10.71
ActiveCell.Offset(0, 11).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 12).Columns("A:A").EntireColumn.ColumnWidth = 1
ActiveCell.Offset(0, 13).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 14).Columns("A:A").EntireColumn.ColumnWidth = 10.71
ActiveCell.Offset(0, 15).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 16).Columns("A:A").EntireColumn.ColumnWidth = 1
ActiveCell.Offset(0, 17).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 18).Columns("A:A").EntireColumn.ColumnWidth = 10.71
ActiveCell.Offset(0, 19).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 20).Columns("A:A").EntireColumn.ColumnWidth = 2


End Sub


