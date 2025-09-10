Attribute VB_Name = "Clients_"

Sub Bouton_Nouveau_Dossier_Client()


Application.ScreenUpdating = False
Application.DisplayAlerts = False

Workbooks.Add
    Call Final_Client_Visites
    Call Final_Client_Certificat
    Call Final_Client_Facture
    Call Final_Client_Adresse
    Call Final_Client_Famille
    Call Final_Client_Dossier
Sheets(Array("Feuil1")).Select
    ActiveWindow.SelectedSheets.Delete
Sheets("Dossier").Select
Range("C8:K8").Select
    ActiveCell.FormulaR1C1 = "=Adresse!R[7]C[4]"
Range("O8:S8").Select
    ActiveCell.FormulaR1C1 = "=Adresse!R[29]C[-8]"
Sheets("Adresse").Select
Range("C8:K8").Select
    ActiveCell.FormulaR1C1 = "=Adresse!R[7]C[4]"
Range("O8:S8").Select
    ActiveCell.FormulaR1C1 = "=Adresse!R[29]C[-8]"
Sheets("Famille").Select
Range("C8:K8").Select
    ActiveCell.FormulaR1C1 = "=Adresse!R[7]C[4]"
Range("O8:S8").Select
    ActiveCell.FormulaR1C1 = "=Adresse!R[29]C[-8]"
Sheets("Visites").Select
Range("C8:K8").Select
    ActiveCell.FormulaR1C1 = "=Adresse!R[7]C[4]"
Range("O8:S8").Select
    ActiveCell.FormulaR1C1 = "=Adresse!R[29]C[-8]"
'verifier l'enregistrement

'ActiveWorkbook.SaveAs Filename:= _
 '       "Macintosh HD:Users:bureaucentral:Documents:Elisa Gassert:Classeurs informatiques:D informatique:Ordonxls"


'ChDir "C:\Documents and Settings\HP_PropriŽtaire\Bureau\REACTIVATION\Back up 001\Exploitation\Classeurs\D\Client"
'ActiveWorkbook.SaveAs Filename:=Worksheets("Adresse").Range("G15")
'Sheets("Adresse").Select


'Worksheets("Adresse").Range ("G15")

End Sub


Sub Bouton_Switch_Facture_Devis_Facture_et_Livraison_1_Page()
    
    
    If Range("B21") = "Devis :" Then
        Range("B21:C21").Select
            ActiveCell.FormulaR1C1 = "Facture :"
    Else
        Range("B21:C21").Select
            ActiveCell.FormulaR1C1 = "Devis :"
    End If


End Sub

Sub Bouton_Switch_Facture_Devis_Facture_et_Livraison_2_Page()
    
    
    If Range("B21") = "Devis :" Then
        Range("B21:C21").Select
            ActiveCell.FormulaR1C1 = "Facture :"
    Else
        Range("B21:C21").Select
            ActiveCell.FormulaR1C1 = "Devis :"
    End If
    If Range("B84") = "Devis :" Then
        Range("B84:C84").Select
            ActiveCell.FormulaR1C1 = "Facture :"
    Else
        Range("B84:C84").Select
            ActiveCell.FormulaR1C1 = "Devis :"
    End If


End Sub

Sub Bouton_Switch_Facture_Devis_Facture_et_Livraison_3_Page()
    
    
    If Range("B21") = "Devis :" Then
        Range("B21:C21").Select
            ActiveCell.FormulaR1C1 = "Facture :"
    Else
        Range("B21:C21").Select
            ActiveCell.FormulaR1C1 = "Devis :"
    End If
    If Range("B84") = "Devis :" Then
        Range("B84:C84").Select
            ActiveCell.FormulaR1C1 = "Facture :"
    Else
        Range("B84:C84").Select
            ActiveCell.FormulaR1C1 = "Devis :"
    End If
    If Range("B147") = "Devis :" Then
        Range("B147:C147").Select
            ActiveCell.FormulaR1C1 = "Facture :"
    Else
        Range("B147:C147").Select
            ActiveCell.FormulaR1C1 = "Devis :"
    End If

End Sub


Sub Bouton_Mise_ˆ_Jour_Facture_et_Livraison_1_Page()


Application.ScreenUpdating = False

'facture
Sheets("Adresse").Range("G13").Copy Sheets("Facture").Range("F11:I11")
Sheets("Adresse").Range("G15").Copy Sheets("Facture").Range("F12:J12")
Sheets("Adresse").Range("G17").Copy Sheets("Facture").Range("F13:I13")
    If Sheets("Adresse").Range("G19") > 0 Then
        Sheets("Adresse").Range("G19").Copy Sheets("Facture").Range("F14:I14")
        Sheets("Adresse").Range("G21").Copy Sheets("Facture").Range("F15:I15")
        Sheets("Adresse").Range("G23").Copy Sheets("Facture").Range("F16:I16")
    Else
        Sheets("Adresse").Range("G21").Copy Sheets("Facture").Range("F14:I14")
    End If
Sheets("Adresse").Range("G35").Copy Sheets("Facture").Range("D21:E21")
'Sheets("Adresse").Range("G37").Copy Sheets("Facture").Range("D22:E22")
Sheets("Adresse").Range("G39").Copy Sheets("Facture").Range("G21:I21")
'livraison
Sheets("Adresse").Range("O13").Copy Sheets("Facture").Range("P11:S11")
Sheets("Adresse").Range("O15").Copy Sheets("Facture").Range("P12:T12")
Sheets("Adresse").Range("O17").Copy Sheets("Facture").Range("P13:S13")
    If Sheets("Adresse").Range("O19") > 0 Then
        Sheets("Adresse").Range("O19").Copy Sheets("Facture").Range("P14:S14")
        Sheets("Adresse").Range("O21").Copy Sheets("Facture").Range("P15:S15")
        Sheets("Adresse").Range("O23").Copy Sheets("Facture").Range("P16:S16")
    Else
        Sheets("Adresse").Range("O21").Copy Sheets("Facture").Range("P14:S14")
    End If
    If Sheets("Adresse").Range("O25") > 0 Then
        Sheets("Adresse").Range("O25").Copy Sheets("Facture").Range("P18:S18")
    Else
        Sheets("Adresse").Range("O27").Copy Sheets("Facture").Range("P18:S18")
    End If
Sheets("Adresse").Range("O35").Copy Sheets("Facture").Range("N21:O21")
'Sheets("Adresse").Range("O37").Copy Sheets("Facture").Range("N22:O22")
Sheets("Adresse").Range("O39").Copy Sheets("Facture").Range("Q21:S21")
'certificat
Sheets("Adresse").Range("G13").Copy Sheets("Certificat").Range("F11:I11")
Sheets("Adresse").Range("G15").Copy Sheets("Certificat").Range("F12:J12")
Sheets("Adresse").Range("G17").Copy Sheets("Certificat").Range("F13:I13")
    If Sheets("Adresse").Range("G19") > 0 Then
        Sheets("Adresse").Range("G19").Copy Sheets("Certificat").Range("F14:I14")
        Sheets("Adresse").Range("G21").Copy Sheets("Certificat").Range("F15:I15")
        Sheets("Adresse").Range("G23").Copy Sheets("Certificat").Range("F16:I16")
    Else
        Sheets("Adresse").Range("G21").Copy Sheets("Certificat").Range("F14:I14")
    End If
Sheets("Adresse").Range("G35").Copy Sheets("Certificat").Range("D21:E21")
'Sheets("Adresse").Range("G37").Copy Sheets("Certificat").Range("D22:E22")
Sheets("Adresse").Range("G39").Copy Sheets("Certificat").Range("G21:I21")




End Sub


Sub Bouton_Mise_ˆ_Jour_Facture_et_Livraison_2_Pages()


Application.ScreenUpdating = False

    Call Bouton_Mise_ˆ_Jour_Facture_et_Livraison_1_Page
'facture 2
Sheets("Adresse").Range("G13").Copy Sheets("Facture").Range("F74:I74")
Sheets("Adresse").Range("G15").Copy Sheets("Facture").Range("F75:J75")
Sheets("Adresse").Range("G17").Copy Sheets("Facture").Range("F76:I76")
    If Sheets("Adresse").Range("G19") > 0 Then
        Sheets("Adresse").Range("G19").Copy Sheets("Facture").Range("F77:I77")
        Sheets("Adresse").Range("G21").Copy Sheets("Facture").Range("F78:I78")
        Sheets("Adresse").Range("G23").Copy Sheets("Facture").Range("F79:I79")
    Else
        Sheets("Adresse").Range("G21").Copy Sheets("Facture").Range("F77:I77")
    End If
Sheets("Adresse").Range("G35").Copy Sheets("Facture").Range("D84:E84")
'Sheets("Adresse").Range("G37").Copy Sheets("Facture").Range("D85:E85")
Sheets("Adresse").Range("G39").Copy Sheets("Facture").Range("G84:I84")
'livraison 2
Sheets("Adresse").Range("O13").Copy Sheets("Facture").Range("P74:S74")
Sheets("Adresse").Range("O15").Copy Sheets("Facture").Range("P75:T75")
Sheets("Adresse").Range("O17").Copy Sheets("Facture").Range("P76:S76")
    If Sheets("Adresse").Range("O19") > 0 Then
        Sheets("Adresse").Range("O19").Copy Sheets("Facture").Range("P77:S77")
        Sheets("Adresse").Range("O21").Copy Sheets("Facture").Range("P78:S78")
        Sheets("Adresse").Range("O23").Copy Sheets("Facture").Range("P79:S79")
    Else
        Sheets("Adresse").Range("O21").Copy Sheets("Facture").Range("P77:S77")
    End If
    If Sheets("Adresse").Range("O25") > 0 Then
        Sheets("Adresse").Range("O25").Copy Sheets("Facture").Range("P81:S81")
    Else
        Sheets("Adresse").Range("O27").Copy Sheets("Facture").Range("P81:S81")
    End If
Sheets("Adresse").Range("O35").Copy Sheets("Facture").Range("N84:O84")
'Sheets("Adresse").Range("O37").Copy Sheets("Facture").Range("N85:O85")
Sheets("Adresse").Range("O39").Copy Sheets("Facture").Range("Q84:S84")


End Sub


Sub Bouton_Mise_ˆ_Jour_Facture_et_Livraison_3_Pages()


Application.ScreenUpdating = False

    Call Bouton_Mise_ˆ_Jour_Facture_et_Livraison_2_Pages
'facture 2
Sheets("Adresse").Range("G13").Copy Sheets("Facture").Range("F137:I137")
Sheets("Adresse").Range("G15").Copy Sheets("Facture").Range("F138:J138")
Sheets("Adresse").Range("G17").Copy Sheets("Facture").Range("F139:I139")
    If Sheets("Adresse").Range("G19") > 0 Then
        Sheets("Adresse").Range("G19").Copy Sheets("Facture").Range("F140:I140")
        Sheets("Adresse").Range("G21").Copy Sheets("Facture").Range("F141:I141")
        Sheets("Adresse").Range("G23").Copy Sheets("Facture").Range("F142:I142")
    Else
        Sheets("Adresse").Range("G21").Copy Sheets("Facture").Range("F140:I140")
    End If
Sheets("Adresse").Range("G35").Copy Sheets("Facture").Range("D147:E147")
'Sheets("Adresse").Range("G37").Copy Sheets("Facture").Range("D148:E148")
Sheets("Adresse").Range("G39").Copy Sheets("Facture").Range("G147:I147")
'livraison 2
Sheets("Adresse").Range("O13").Copy Sheets("Facture").Range("P137:S137")
Sheets("Adresse").Range("O15").Copy Sheets("Facture").Range("P138:T138")
Sheets("Adresse").Range("O17").Copy Sheets("Facture").Range("P139:S139")
    If Sheets("Adresse").Range("O19") > 0 Then
        Sheets("Adresse").Range("O19").Copy Sheets("Facture").Range("P140:S140")
        Sheets("Adresse").Range("O21").Copy Sheets("Facture").Range("P141:S141")
        Sheets("Adresse").Range("O23").Copy Sheets("Facture").Range("P142:S142")
    Else
        Sheets("Adresse").Range("O21").Copy Sheets("Facture").Range("P140:S140")
    End If
    If Sheets("Adresse").Range("O25") > 0 Then
        Sheets("Adresse").Range("O25").Copy Sheets("Facture").Range("P144:S144")
    Else
        Sheets("Adresse").Range("O27").Copy Sheets("Facture").Range("P144:S144")
    End If
Sheets("Adresse").Range("O35").Copy Sheets("Facture").Range("N147:O147")
'Sheets("Adresse").Range("O37").Copy Sheets("Facture").Range("N148:O148")
Sheets("Adresse").Range("O39").Copy Sheets("Facture").Range("Q147:S147")


End Sub




Sub finition_facture_acompte()

'verifier les connections


Range("F55").Select
    Selection.HorizontalAlignment = xlRight
    ActiveCell.FormulaR1C1 = "Acompte"
Range("J55").Select
    Selection.NumberFormat = "#,##0.00"
Range("J56").Select
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("D58").Select
    Selection.HorizontalAlignment = xlRight
    ActiveCell.FormulaR1C1 = "Payement ˆ la livraison."
Range("F58").Select
    Selection.HorizontalAlignment = xlRight
    ActiveCell.FormulaR1C1 = "Solde"
Range("J58").Select
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=R[-5]C-R[-3]C"
Range("J59").Select
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("D61").Select
    ActiveCell.FormulaR1C1 = "Merci pour votre confiance."
    
    
End Sub


Sub Bouton_Impression_Facture_et_Livraison_1_Page()
    
    'verifier le code
    
    ActiveWindow.SelectedSheets.PrintOut From:=1, To:=2, Copies:=1, Collate:=True
    
    
End Sub


Sub Bouton_Impression_Facture_et_Livraison_2_Pages()
    
    'verifier le code
    
    ActiveWindow.SelectedSheets.PrintOut From:=1, To:=4, Copies:=1, Collate:=True
    
    
End Sub


Sub Bouton_Impression_Facture_et_Livraison_3_Pages()
    
    'verifier le code
    
    ActiveWindow.SelectedSheets.PrintOut From:=1, To:=6, Copies:=1, Collate:=True
    
    
End Sub


Sub Bouton_Nouvelle_Facture_Livraison_Certificat_1_Page()


Application.ScreenUpdating = False

Sheets("Facture").Select
Rows("1:63").Select
    Selection.Insert Shift:=xlDown
    Call Fiche_Client_Livraison
    Call Fiche_Client_Facture
'Sheets("Certificat").Select
'Rows("1:63").Select
    'Selection.Insert Shift:=xlDown
    'Call Fiche_Client_Certificat


End Sub


Sub Bouton_Nouvelle_Facture_Livraison_Certificat_2_Pages()


Application.ScreenUpdating = False

Sheets("Facture").Select
Rows("1:63").Select
    Selection.Insert Shift:=xlDown
    Call Fiche_Client_Livraison
    Call Fiche_Client_Facture
Rows("1:63").Select
    Selection.Insert Shift:=xlDown
    Call Fiche_Client_Livraison
    Call Fiche_Client_Facture
Range("F30").Select
    ActiveCell.FormulaR1C1 = "Page 1"
Range("F53").Select
    ActiveCell.FormulaR1C1 = "Sous total 1"
Range("J54").Select
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("D93").Select
    ActiveCell.FormulaR1C1 = "=R[-63]C"
Range("F93").Select
    ActiveCell.FormulaR1C1 = "Page 2"
Range("B95").Select
    ActiveCell.FormulaR1C1 = "1"
Range("D95").Select
    ActiveCell.FormulaR1C1 = "sous total page 1"
Range("H95").Select
    ActiveCell.FormulaR1C1 = "=R[-42]C[2]"
Range("P30").Select
    ActiveCell.FormulaR1C1 = "Page 1"
Range("P93").Select
    ActiveCell.FormulaR1C1 = "Page 2"
Sheets("Certificat").Select
Rows("1:63").Select
    Selection.Insert Shift:=xlDown
Rows("1:63").Select
    Selection.Insert Shift:=xlDown
    Call Fiche_Client_Certificat


End Sub


Sub Bouton_Nouvelle_Facture_Livraison_Certificat_3_Pages()


Application.ScreenUpdating = False


Sheets("Facture").Select
Rows("1:63").Select
    Selection.Insert Shift:=xlDown
    Call Fiche_Client_Livraison
    Call Fiche_Client_Facture
Rows("1:63").Select
    Selection.Insert Shift:=xlDown
    Call Fiche_Client_Livraison
    Call Fiche_Client_Facture
Rows("1:63").Select
    Selection.Insert Shift:=xlDown
    Call Fiche_Client_Livraison
    Call Fiche_Client_Facture
Range("F30").Select
    ActiveCell.FormulaR1C1 = "Page 1"
Range("F53").Select
    ActiveCell.FormulaR1C1 = "Sous total 1"
Range("J54").Select
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("D93").Select
    ActiveCell.FormulaR1C1 = "=R[-63]C"
Range("F93").Select
    ActiveCell.FormulaR1C1 = "Page 2"
Range("B95").Select
    ActiveCell.FormulaR1C1 = "1"
Range("D95").Select
    ActiveCell.FormulaR1C1 = "sous total page 1"
Range("H95").Select
    ActiveCell.FormulaR1C1 = "=R[-42]C[2]"
Range("F116").Select
    ActiveCell.FormulaR1C1 = "sous total 2"
Range("J117").Select
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("D156").Select
    ActiveCell.FormulaR1C1 = "=R[-126]C"
Range("F156").Select
    ActiveCell.FormulaR1C1 = "Page 3"
Range("B158").Select
    ActiveCell.FormulaR1C1 = "1"
Range("D158").Select
    ActiveCell.FormulaR1C1 = "sous total page 2"
Range("H158").Select
    ActiveCell.FormulaR1C1 = "=R[-42]C[2]"
Range("P30").Select
    ActiveCell.FormulaR1C1 = "Page 1"
Range("P93").Select
    ActiveCell.FormulaR1C1 = "Page 2"
Range("P156").Select
    ActiveCell.FormulaR1C1 = "Page 3"
Sheets("Certificat").Select
Rows("1:63").Select
    Selection.Insert Shift:=xlDown
Rows("1:63").Select
    Selection.Insert Shift:=xlDown
Rows("1:63").Select
    Selection.Insert Shift:=xlDown
    Call Fiche_Client_Certificat


End Sub



