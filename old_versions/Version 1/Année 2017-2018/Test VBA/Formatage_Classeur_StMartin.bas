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
    
    Decal = ref.Range("P3").Value
    
    
    For Index = 1 To Sheets.Count Step 1
        If Worksheets(Index).Name Like "Elève*" Then
            Worksheets(Index).Select
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
                
                Col = Col + Decal
                Col_T1 = Col_T1 + Decal
                Col_T2 = Col_T2 + Decal
                Col_T3 = Col_T3 + Decal
                Col_An = Col_An + Decal
            Next Indice_eleve
                
            'Formatage de tous les élèves à partir du modèle du 2e
            Appliquer_Tous_Eleves
            
            Range("C1").Select
       End If
    Next Index
    
    Set Plage = Nothing
  
End Sub

Sub Formatage_ref()
Attribute Formatage_ref.VB_ProcData.VB_Invoke_Func = "R\n14"
'
' Formatage_ref Macro
'
' Touche de raccourci du clavier: Ctrl+Shift+R

      
' Mise en forme
    For Index = 1 To Sheets.Count Step 1
        If Worksheets(Index).Name Like "ref*" Then
            Worksheets(Index).Select
            Worksheets(Index).Name = "ref"
            
            Range("M:N,P:P").ColumnWidth = 12
            Columns("O:O").ColumnWidth = 4
            
            'Liste des domaines
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
            
            'Evaluation par trimestre (valeurs)
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
            
            ' Evaluations par trimestre (titre)
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
            
            ' Evaluation par trimestre (valeurs)
            With Range("N3:N5")
                .Locked = False
                .FormulaHidden = False
            End With
            
            'Decalage
            Set Zone = Range("P2,P3")
            With Zone
                .Borders(xlEdgeBottom).Weight = xlMedium
                .Borders(xlEdgeTop).Weight = xlMedium
                .Borders(xlEdgeLeft).Weight = xlMedium
                .Borders(xlEdgeRight).Weight = xlMedium
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            Range("P2").Interior.Color = 3628543
            Range("P2").Value = "Décalage"
            
            'Remplissage des cases / applications des formules
            Range("M2:N2").Select
            ActiveCell.FormulaR1C1 = "Evaluations par trimestre"
            Range("M3").Select
            ActiveCell.FormulaR1C1 = "1er tri"
            Range("M4").Select
            ActiveCell.FormulaR1C1 = "2e tri"
            Range("M5").Select
            ActiveCell.FormulaR1C1 = "3e tri"
            
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
            
            For Col = 5 To 8 Step 1
                For lig = 4 To 34 Step 1
                    Cells(lig, Col).Formula = "=" & Cells(lig - 1, Col).Address(False, False, , False) & "+ $P$3"
                Next lig
            Next Col
        
        End If
    Next Index
End Sub
