Attribute VB_Name = "Module2"
Sub Formatage_Eleve()
Attribute Formatage_Eleve.VB_ProcData.VB_Invoke_Func = "E\n14"
'
' Formatage_Eleve Macro
'
' Touche de raccourci du clavier: Ctrl+Shift+F
'
    Dim Decal As Integer, Nb_Eleve As Integer, Col_deb As Integer, Col_fin As Integer
    Set ref = Worksheets("ref")
    
    
    For i = 1 To Worksheets.Count
        If Worksheets(i).Name Like "Elève*" Then
            Worksheets(i).Select
            
            Decal = ref.Range("P3").Value
            Col_T1 = ref.Range("E3").Value
            Col_T2 = ref.Range("F3").Value
            Col_T3 = ref.Range("G3").Value
            Col_An = ref.Range("H3").Value
            Col = 3
            
            Range("2:2, 3:3, 4:4, 10:10, 13:13, 16:16, 20:20").UnMerge
            
            'Verouillage des bonnes cellules
            Set Plage = Cells
            Plage.Locked = False
            Plage.FormulaHidden = True
            Union(Range("A1"), Range("B1")).EntireColumn.Locked = True
            Union(Range("B4"), Range("B10"), Range("B13"), Range("B16"), Range("B20")).Locked = False
            
            ' Formatage du 1er et 2e élève
            For Indice_eleve = 1 To 2 Step 1
                Range(Cells(2, Col), Cells(2, Col + Decal - 1)).HorizontalAlignment = xlCenterAcrossSelection
                Range(Cells(3, Col_T1 - 1), Cells(3, Col_T1)).HorizontalAlignment = xlCenterAcrossSelection
                Range(Cells(3, Col_T2 - 1), Cells(3, Col_T2)).HorizontalAlignment = xlCenterAcrossSelection
                Range(Cells(3, Col_T3 - 1), Cells(3, Col_T3)).HorizontalAlignment = xlCenterAcrossSelection
                Range(Cells(3, Col_An - 1), Cells(3, Col_An)).HorizontalAlignment = xlCenterAcrossSelection
                Union(Range(Cells(1, Col), Cells(2, Col)), Cells(4, Col), Cells(10, Col), Cells(13, Col), Cells(16, Col), Cells(20, Col), Range(Cells(1, Col_T1 + 1), Cells(2, Col_T1 + 1)), Cells(4, Col_T1 + 1), Cells(10, Col_T1 + 1), Cells(13, Col_T1 + 1), Cells(16, Col_T1 + 1), Cells(20, Col_T1 + 1), Range(Cells(1, Col_T2 + 1), Cells(2, Col_T2 + 1)), Cells(4, Col_T2 + 1), Cells(10, Col_T2 + 1), Cells(13, Col_T2 + 1), Cells(16, Col_T2 + 1), Cells(20, Col_T2 + 1)).Locked = True
                Union(Range(Cells(1, Col_T1 - 1), Cells(1, Col_T1)), Range(Cells(1, Col_T2 - 1), Cells(1, Col_T2)), Range(Cells(1, Col_T3 - 1), Cells(1, Col_T3 + 2))).EntireColumn.Locked = True
                Union(Range(Cells(22, Col), Cells(22, Col_T1 - 2)), Range(Cells(22, Col_T1 + 1), Cells(22, Col_T2 - 2)), Range(Cells(22, Col_T2 + 1), Cells(22, Col_T3 - 2))).FormulaHidden = False
                            
                Col = Col + Decal
                Col_T1 = Col_T1 + Decal
                Col_T2 = Col_T2 + Decal
                Col_T3 = Col_T3 + Decal
                Col_An = Col_An + Decal
            Next Indice_eleve
            
            
            'Formatage de tous les élèves à partir du modèle du 2e
            Appliquer_Tous_Eleves
            
            Set Plage = Nothing
            Range("C1").Select
        End If
    Next i
      
End Sub

Sub Formatage_ref()
Attribute Formatage_ref.VB_ProcData.VB_Invoke_Func = "R\n14"
'
' Formatage_ref Macro
'
' Touche de raccourci du clavier: Ctrl+Shift+R

      
' Mise en forme
    If ActiveSheet.Name = "ref" Then
        Range("M:N,P:P").ColumnWidth = 12
        Columns("O:O").ColumnWidth = 4
        
        Range("M2:N5").Select
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        
        Range("M2:N2").Select
        With Selection
            .HorizontalAlignment = xlCenterAcrossSelection
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    
        Range("J2:K2").Select
        With Selection
            .HorizontalAlignment = xlCenterAcrossSelection
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 5287936
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
    ' Remplissage des cases / applications des formules
        Range("M2:N2").Select
        ActiveCell.FormulaR1C1 = "Evaluations par trimestre"
        Range("M3").Select
        ActiveCell.FormulaR1C1 = "1er tri"
        Range("M4").Select
        ActiveCell.FormulaR1C1 = "2e tri"
        Range("M5").Select
        ActiveCell.FormulaR1C1 = "3e tri"
        Range("N3").Select
        ActiveCell.FormulaR1C1 = "3"
        Range("N4").Select
        ActiveCell.FormulaR1C1 = "3"
        Range("N5").Select
        ActiveCell.FormulaR1C1 = "3"
        
        Range("E3").Select
        ActiveCell.FormulaR1C1 = "=RC[9]+4"
        Range("F3").Select
        ActiveCell.FormulaR1C1 = "=RC[-1]+R[1]C[8]+2"
        Range("G3").Select
        ActiveCell.FormulaR1C1 = "=RC[-1]+R[2]C[7]+2"
        Range("H3").Select
        ActiveCell.FormulaR1C1 = "=RC[-1]+2"
        Range("P3").Select
        ActiveCell.FormulaR1C1 = "=RC[-2]+R[1]C[-2]+R[2]C[-2]+8"
    Else
        MsgBox ("Veuillez vous mettre sur la feuille nommée 'ref' pour éxecuter cette macro.")
    End If
    
End Sub
