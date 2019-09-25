Attribute VB_Name = "Module4"
' **********************************
' Page 4 (résultats élèves) - Procédure & fonctions
' **********************************

Sub creerTableauBilan(nomClasse As String, indexClasse As Integer, nombreEleves As Integer)

    ' Creation page
    ActiveWorkbook.Unprotect Password
    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Name = "Bilan (" & nomClasse & ")"
    ActiveWorkbook.Protect Password, True, True
    With Cells
        .Borders.ColorIndex = 2
        .Locked = True
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
    End With
    
    ' Figeage des volets
    With ActiveWindow
        .SplitColumn = 1
        .SplitRow = 3
        .FreezePanes = True
    End With
    
    '**** COLONNE INFOS + LISTE ELEVE ****
    ' Taille ligne/colonne
    For ligne = 1 To nombreEleves + 3
        If ligne < 4 Then
            Rows(ligne).RowHeight = 25
        Else
            Rows(ligne).RowHeight = 15
        End If
    Next ligne
    Columns.ColumnWidth = 6
    Columns(1).ColumnWidth = 40
    
    ' Bouton 'actualiser résultats'
    Set buttonCell = Range("A1")
    Set Button = ActiveSheet.Buttons.Add(buttonCell.Left, buttonCell.Top, buttonCell.Width, buttonCell.Height)
    With Button
        .Caption = "Actualiser résultats"
        .OnAction = "btnActualiserResultats_Click"
    End With
    
    ' Légende
    With Range("A3")
        .Value = nomClasse
        .Interior.ColorIndex = colorindexClasse
        .Borders.ColorIndex = xlColorIndexAutomatic
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
    End With
    
    ' Liste élève
    For indexEleve = 1 To nombreEleves
        With Cells(3 + indexEleve, 1)
            .Value = Sheets(Page2).Cells(3 + indexEleve, indexClasse * 2 - 1).Value
        End With
    Next indexEleve
    With Range(Cells(4, 1), Cells(3 + nombreEleves, 1))
        .HorizontalAlignment = xlHAlignLeft
        .Borders.ColorIndex = xlColorIndexAutomatic
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    '**** LIGNE EN-TETE + CONTENU ****
    nombreDomaines = Sheets(Page1).Cells(10, 3).Value
    With Range(Cells(1, 2), Cells(1, 1 + 4 * (nombreDomaines + 1)))
        .Interior.ColorIndex = colorindexBilan
        .MergeCells = True
        .Value = "Bilan trimestriel & annuel"
    End With
    For indexDomaine = 1 To nombreDomaines + 1
        If indexDomaine <= nombreDomaines Then
            With Range(Cells(2, 2 + 4 * (indexDomaine - 1)), Cells(2, 5 + 4 * (indexDomaine - 1)))
                .Interior.ColorIndex = colorindexDomaine
                .MergeCells = True
                .Value = "D" & indexDomaine
            End With
            Range(Cells(3, 5 + 4 * (indexDomaine - 1)), Cells(3 + nombreEleves, 5 + 4 * (indexDomaine - 1))).Interior.ColorIndex = colorindexDomaine2
            
        Else
            With Range(Cells(2, 2 + 4 * (indexDomaine - 1)), Cells(2, 5 + 4 * (indexDomaine - 1)))
                .Interior.ColorIndex = colorindexNote
                .MergeCells = True
                .Value = "Note globale"
            End With
            Range(Cells(3, 5 + 4 * (indexDomaine - 1)), Cells(3 + nombreEleves, 5 + 4 * (indexDomaine - 1))).Interior.ColorIndex = colorindexNote2
        End If
        Cells(3, 2 + 4 * (indexDomaine - 1)).Value = "1e tri"
        Cells(3, 3 + 4 * (indexDomaine - 1)).Value = "2e tri"
        Cells(3, 4 + 4 * (indexDomaine - 1)).Value = "3e tri"
        Cells(3, 5 + 4 * (indexDomaine - 1)).Value = "Année"
        With Range(Cells(2, 2 + 4 * (indexDomaine - 1)), Cells(3 + nombreEleves, 5 + 4 * (indexDomaine - 1)))
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
    Range(Cells(1, 2), Cells(3 + nombreEleves, 1 + 4 * (nombreDomaines + 1))).BorderAround xlDouble, xlThin, xlColorIndexAutomatic
    Range(Cells(4, 1), Cells(3 + nombreEleves, 1 + 4 * (nombreDomaines + 1))).BorderAround xlDouble, xlThin, xlColorIndexAutomatic
    
    Range(Cells(4, 2), Cells(3 + nombreEleves, 1 + 4 * (nombreDomaines + 1))).Cells.Locked = False
    
    ' Protection feuille
    ActiveSheet.EnableSelection = xlUnlockedCells
    ActiveSheet.Protect Password
    
End Sub

Sub btnActualiserResultats_Click()
    Dim indexDomaine As Integer, indexTrimestre As Integer

    ' Valeurs nécessaires
    nomClasse = Range("A3").Value
    nombreDomaines = getNombreDomaines
    
    ' Retrait protection page notes
    Sheets("Bilan (" & nomClasse & ")").Unprotect Password
    
    For indexTrimestre = 1 To 4
        For indexDomaine = 1 To nombreDomaines
            calculMoyenneDomaine indexDomaine, indexTrimestre
        Next indexDomaine
        calculMoyenneTrimestre indexTrimestre
    Next indexTrimestre
    
    ' Protection page notres
    Sheets("Bilan (" & nomClasse & ")").Protect Password
    
    MsgBox ("Données mises à jour")
End Sub

' Calcul de la moyenne trimestrielle/annuelle pour chaque domaine
' indexTrimestre = 4 pour indiquer l'année
Sub calculMoyenneDomaine(indexDomaine As Integer, indexTrimestre As Integer)
    Dim nomClasse As String, lettre As String

    ' Valeurs nécessaires
    nomClasse = Range("A3").Value
    nombreEvals = Sheets("Notes (" & nomClasse & ")").Buttons.Count - 1
    nombreCompetences = getNombreCompetences(indexDomaine)
    nombreTotalCompetences = getNombreCompetences
    nombreEleves = getNombreEleves(nomClasse)
    nombreDomaines = getNombreDomaines
    moitieCompetences = (nombreTotalCompetences - nombreTotalCompetences Mod 2) / 2
    indexReference = 1
    
    ' Vérfication des entrées
    If indexDomaine <= nombreDomaines And (indexTrimestre > 0 Or indexTrimestre < 5) Then
        
        ' Calcul indexReference = colonne du domaine concerné
        For domaine = 1 To indexDomaine
            If domaine <> indexDomaine Then
                indexReference = indexReference + Sheets(Page1).Cells(12 + domaine, 3).Value
            End If
        Next domaine
        
        ' Calcul de la moyenne
        For indexEleve = 1 To nombreEleves
            somme = 0
            diviseur = 0
            For indexEval = 1 To nombreEvals
                If indexTrimestre = 4 Or Sheets("Notes (" & nomClasse & ")").Cells(2, 3 + (indexEval - 1) * (nombreTotalCompetences + 1)).Value = indexTrimestre Then
                    For indexCompetence = indexReference To indexReference + nombreCompetences - 1
                        lettre = Sheets("Notes (" & nomClasse & ")").Cells(5 + indexEleve, 2 + (indexEval - 1) * (nombreTotalCompetences + 1) + indexCompetence).Value
                        coeffCompetence = Sheets("Notes (" & nomClasse & ")").Cells(5, 2 + (indexEval - 1) * (nombreTotalCompetences + 1) + indexCompetence).Value
                        If StrComp(lettre, "") <> 0 And IsEmpty(coeffCompetence) = False Then
                            somme = somme + coeffCompetence * lettreToValeur(lettre)
                            diviseur = diviseur + coeffCompetence
                        End If
                    Next indexCompetence
                End If
            Next indexEval
            If somme <> 0 Then
                Cells(3 + indexEleve, 1 + 4 * (indexDomaine - 1) + indexTrimestre).Value = valeurToLettre(somme / diviseur)
            ElseIf somme = 0 And diviseur = 0 Then
                Cells(3 + indexEleve, 1 + 4 * (indexDomaine - 1) + indexTrimestre).Value = ""
            End If
        Next indexEleve
    End If
    
End Sub

' Calcul la moyenne des notes du trimestre
Sub calculMoyenneTrimestre(indexTrimestre As Integer)
    Dim nomClasse As String

    ' Valeurs nécessaires
    nomClasse = Range("A3").Value
    nombreEvals = Sheets("Notes (" & nomClasse & ")").Buttons.Count - 1
    nombreTotalCompetences = getNombreCompetences
    nombreEleves = getNombreEleves(nomClasse)
    nombreDomaines = getNombreDomaines
    moitieCompetences = (nombreTotalCompetences - nombreTotalCompetences Mod 2) / 2
    
    ' Calcul de la moyenne
    For indexEleve = 1 To nombreEleves
        somme = 0
        diviseur = 0
        For indexEval = 1 To nombreEvals
            If indexTrimestre = 4 Or Sheets("Notes (" & nomClasse & ")").Cells(2, 3 + (indexEval - 1) * (nombreTotalCompetences + 1)).Value = indexTrimestre Then
                note = Sheets("Notes (" & nomClasse & ")").Cells(5 + indexEleve, 2 + (indexEval) * (nombreTotalCompetences + 1)).Value
                If IsEmpty(note) = False Then
                    coeffEval = Sheets("Notes (" & nomClasse & ")").Cells(2, 3 + moitieCompetences + (indexEval - 1) * (nombreTotalCompetences + 1)).Value
                    somme = somme + coeffEval * note
                    diviseur = diviseur + coeffEval
                End If
            End If
        Next indexEval
        If somme <> 0 Then
            Cells(3 + indexEleve, 1 + 4 * nombreDomaines + indexTrimestre).Value = Format(somme / diviseur, "Standard")
        ElseIf somme = 0 And diviseur = 0 Then
            Cells(3 + indexEleve, 1 + 4 * nombreDomaines + indexTrimestre).Value = ""
        End If
    Next indexEleve
    
End Sub
