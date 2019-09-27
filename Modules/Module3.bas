Attribute VB_Name = "Module3"
' **********************************
' Page 3 (entrée notes) - Procédure & fonctions
' **********************************

Sub creerTableauNotes(nomClasse As String, indexClasse As Integer, nombreEleves As Integer)

    ' Creation page
    ActiveWorkbook.Unprotect strPassword
    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Name = "Notes (" & nomClasse & ")"
    ActiveWorkbook.Protect strPassword, True, True
    With Cells
        .Borders.ColorIndex = 2
        .Locked = True
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
    End With
    
    '**** COLONNE INFOS + LISTE ELEVE ****
    ' Taille ligne/colonne
    For ligne = 1 To nombreEleves + 5
        If ligne < 4 Then
            Rows(ligne).RowHeight = 20
        ElseIf ligne = 4 Or ligne = 5 Then
            Rows(ligne).RowHeight = 30
        Else
            Rows(ligne).RowHeight = 15
        End If
    Next ligne
    For colonne = 1 To 2
        Columns(colonne).ColumnWidth = 25
    Next colonne
    
    ' Bouton 'ajouter éval'
    Set buttonCell = Range("A1")
    Set Button = ActiveSheet.Buttons.Add(buttonCell.Left, buttonCell.Top, buttonCell.Width, buttonCell.Height)
    With Button
        .Caption = "Ajouter évaluation"
        .OnAction = "btnAjouterEvaluation_Click"
    End With
    
    ' Légende
    With Range("A5")
        .Value = nomClasse
        .Interior.ColorIndex = intColorClasse
        .Borders.ColorIndex = xlColorIndexAutomatic
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
    End With
    Range("B1").Value = "Nom de l'évaluation"
    Range("B2").Value = "Trimestre / Coeff"
    Range("B3").Value = "Domaines"
    Range("B4").Value = "Compétences"
    Range("B5").Value = "Coeff compétence"
    With Range("B1:B2,B3:B5")
        .Borders.ColorIndex = xlColorIndexAutomatic
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeRight).Weight = xlMedium
    End With
    Range("B1").Interior.ColorIndex = intColorEval
    Range("B3").Interior.ColorIndex = intColorDomaine
    Range("B4").Interior.ColorIndex = intColorDomaine2
    
    ' Liste élève
    For indexEleve = 1 To nombreEleves
        With Cells(5 + indexEleve, 1)
            .Value = Sheets(strPage2).Cells(3 + indexEleve, indexClasse * 2 - 1).Value
        End With
        Range(Cells(5 + indexEleve, 1), Cells(5 + indexEleve, 2)).MergeCells = True
    Next indexEleve
    With Range(Cells(6, 1), Cells(5 + nombreEleves, 2))
        .HorizontalAlignment = xlHAlignLeft
        .Borders.ColorIndex = xlColorIndexAutomatic
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    ' Figeage des volets
    With ActiveWindow
        .SplitColumn = 2
        .SplitRow = 5
        .FreezePanes = True
    End With
    
    '**** 1e EVALUATION ****
    colonneDepart = 3
    
    ajouterEvaluation (colonneDepart)
    
    
    ' Protection feuille
    ActiveSheet.EnableSelection = xlUnlockedCells
    ActiveSheet.Protect strPassword:=strPassword
    
End Sub

Sub btnAjouterEvaluation_Click()

    Dim nombreEleves As Integer, nombreDomaines As Integer, nombreCompetences As Integer, colonneDepart As Integer, indexEval As Integer, indexDomaine As Integer
    Dim totalCompetences As Integer
    nombreDomaines = getNombreDomaines
    nombreEleves = getNombreEleves(ActiveSheet.Cells(5, 1).Value)
        
    ' Determine la colonne où ajouter l'éval
    indexEval = ActiveSheet.Buttons.Count - 1
    nombreCompetences = getNombreCompetences
    colonneDepart = 3 + indexEval * (nombreCompetences + 1)
    
    ' Retrait protection feuille
    ActiveSheet.Unprotect strPassword
    
    ' Ajout de l'évaluation
    ajouterEvaluation (colonneDepart)
        
    ' Protection feuille
    ActiveSheet.EnableSelection = xlUnlockedCells
    ActiveSheet.Protect strPassword:=strPassword
    
    MsgBox ("Évaluation ajoutée")

End Sub
    
Sub ajouterEvaluation(colonneDepart As Integer)
    Dim nombreEleves As Integer, nombreDomaines As Integer, totalCompetences As Integer
    Dim indexDomaine As Integer, indexCompetences As Integer
    
    ' Calcul données nécessaires
    nombreDomaines = getNombreDomaines
    nombreEleves = getNombreEleves(ActiveSheet.Cells(5, 1).Value)
    
    ' Domaines/Compétences
    totalCompetences = 0
    For indexDomaine = 1 To nombreDomaines
        intNombreCompetences = getNombreCompetences(indexDomaine)
        For indexCompetence = 1 To intNombreCompetences
            totalCompetences = totalCompetences + 1
            Columns(colonneDepart + totalCompetences - 1).ColumnWidth = 3
            With Cells(4, colonneDepart + totalCompetences - 1)
                .Value = "D" & indexDomaine & "/" & indexCompetence
                .Orientation = xlUpward
                .Interior.ColorIndex = intColorDomaine2
            End With
        Next indexCompetence
        Cells(3, colonneDepart + totalCompetences - indexCompetence + 1).Value = "D" & indexDomaine
        Range(Cells(3, colonneDepart + totalCompetences - indexCompetence + 1), Cells(3, colonneDepart + totalCompetences - 1)).Interior.ColorIndex = intColorDomaine
        Range(Cells(3, colonneDepart + totalCompetences - indexCompetence + 1), Cells(3, colonneDepart + totalCompetences - 1)).MergeCells = True
        With Range(Cells(3, colonneDepart + totalCompetences - indexCompetence + 1), Cells(5 + nombreEleves, colonneDepart + totalCompetences - 1))
            .Borders.ColorIndex = xlColorIndexAutomatic
            .Borders.LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlMedium
            .Borders(xlEdgeTop).Weight = xlMedium
            .Borders(xlEdgeLeft).Weight = xlMedium
            .Borders(xlEdgeRight).Weight = xlMedium
            .Borders(xlInsideHorizontal).Weight = xlThin
            .Borders(xlInsideVertical).Weight = xlHairline
        End With
    Next indexDomaine
    
    ' Infos Eval
    moitieCompetences = (totalCompetences - totalCompetences Mod 2) / 2
    Range(Cells(1, colonneDepart), Cells(1, colonneDepart + totalCompetences - 1)).Interior.ColorIndex = intColorEval
    Range(Cells(1, colonneDepart), Cells(1, colonneDepart + totalCompetences - 1)).MergeCells = True
    Range(Cells(2, colonneDepart), Cells(2, colonneDepart + moitieCompetences - 1)).MergeCells = True
    Range(Cells(2, colonneDepart + moitieCompetences), Cells(2, colonneDepart + totalCompetences - 1)).MergeCells = True
    Set buttonCell = Range(Cells(1, colonneDepart + totalCompetences), Cells(2, colonneDepart + totalCompetences))
    Set Button = ActiveSheet.Buttons.Add(buttonCell.Left, buttonCell.Top, buttonCell.Width, buttonCell.Height)
    With Button
        .Caption = "Calcul note"
        .OnAction = "btnCalculNote_Click"
    End With
    With Range(Cells(3, colonneDepart + totalCompetences), Cells(5, colonneDepart + totalCompetences))
        .Interior.ColorIndex = intColorNote
        .MergeCells = True
        .Orientation = xlUpward
        .Value = "Note / 20"
        .Columns.ColumnWidth = 6
    End With
    Range(Cells(6, colonneDepart + totalCompetences), Cells(5 + nombreEleves, colonneDepart + totalCompetences)).Interior.ColorIndex = intColorNote2
    With Range(Cells(3, colonneDepart + totalCompetences), Cells(5 + nombreEleves, colonneDepart + totalCompetences))
        .Borders.ColorIndex = xlColorIndexAutomatic
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlInsideHorizontal).Weight = xlThin
    End With
    With Range(Cells(1, colonneDepart), Cells(2, colonneDepart + totalCompetences - 1))
        .Borders.ColorIndex = xlColorIndexAutomatic
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeRight).Weight = xlMedium
    End With
    Range(Cells(1, colonneDepart), Cells(5 + nombreEleves, colonneDepart + totalCompetences)).BorderAround xlDouble, xlThin, xlColorIndexAutomatic
    Range(Cells(6, 1), Cells(5 + nombreEleves, colonneDepart + totalCompetences)).BorderAround xlDouble, xlThin, xlColorIndexAutomatic
    Range(Cells(1, colonneDepart), Cells(2, colonneDepart + totalCompetences - 1)).Locked = False
    Range(Cells(5, colonneDepart), Cells(5 + nombreEleves, colonneDepart + totalCompetences - 1)).Locked = False
    
End Sub

Sub btnCalculNote_Click()

    ' Determiner éval à calculer
    indexEval = Val(Right(Application.Caller, 1)) - 1
    nombreCompetences = getNombreCompetences
    colonneDepart = 3 + (indexEval - 1) * (nombreCompetences + 1)
    
    ' Retrait protection feuille
    ActiveSheet.Unprotect strPassword
    
    ' Calcul note éval
    calculNote (colonneDepart)
        
    ' Protection feuille
    ActiveSheet.EnableSelection = xlUnlockedCells
    ActiveSheet.Protect strPassword
    
End Sub

Sub calculNote(colonneDepart As Integer)
    Dim lettre As String

    ' Calcul données nécessaires
    nombreEleves = getNombreEleves(ActiveSheet.Cells(5, 1).Value)
    nombreCompetences = getNombreCompetences
    
    ' Calcul note éval
    For indexEleve = 1 To nombreEleves
        diviseur = 0
        somme = 0
        For indexCompetence = 1 To nombreCompetences
            lettre = ActiveSheet.Cells(5 + indexEleve, colonneDepart + indexCompetence - 1).Value
            coeffCompetence = ActiveSheet.Cells(5, colonneDepart + indexCompetence - 1).Value
            If StrComp(lettre, "") <> 0 And IsEmpty(coeffCompetence) = False Then
                somme = somme + lettreToValeur(lettre) * coeffCompetence
                diviseur = diviseur + coeffCompetence
            End If
        Next indexCompetence
        If diviseur <> 0 Then
            Cells(5 + indexEleve, colonneDepart + nombreCompetences).Value = Format(5 * somme / diviseur, "Standard")
        ElseIf somme = 0 And diviseur = 0 Then
            Cells(5 + indexEleve, colonneDepart + nombreCompetences).Value = ""
        End If
    Next indexEleve
End Sub

Function lettreToValeur(lettre As String)
    asciiLettre = Asc(lettre)
    If asciiLettre > 64 And asciiLettre < 70 Then
        lettreToValeur = 69 - asciiLettre
    Else
        lettreToValeur = 0
    End If
End Function

Function valeurToLettre(valeur As Integer)
    If valeur >= 0 And valeur <= 4 Then
        Select Case valeur
            Case Is > 3.3
                valeurToLettre = "A"
            Case Is > 2.3
                valeurToLettre = "B"
            Case Is > 1
                valeurToLettre = "C"
            Case Is > 0
                valeurToLettre = "D"
            Case Else
                valeurToLettre = "E"
        End Select
    Else
        valeurToLettre = "Z"
    End If
End Function
