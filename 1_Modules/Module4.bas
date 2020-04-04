Attribute VB_Name = "Module4"
' ##################################
' PAGE 4 (résultats élèves)
' ##################################

Option Explicit

' **********************************
' FONCTIONS
' **********************************

' Aucune fonction

' **********************************
' PROCÉDURES
' **********************************

Sub creerTableauBilan(intIndiceClasse As Integer, intNombreEleves As Integer)
    Dim intIndiceLigne As Integer
    Dim rngCelluleBouton As Range, btnBouton As Variant
    Dim strNomClasse As String
    Dim intIndiceEleve As Integer
    Dim intNombreDomaines As Integer, intIndiceDomaine As Integer

    ' Données nécessaires
    strNomClasse = getNomClasse(intIndiceClasse)

    Application.ScreenUpdating = False
    
    ' Creation page
    ActiveWorkbook.Unprotect strPassword
    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Name = "Bilan (" & strNomClasse & ")"
    ActiveWorkbook.Protect strPassword, True, True
    With Cells
        .Borders.ColorIndex = 2
        .Locked = True
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
    End With
    
    ' Figeage des volets
    freezePanes ActiveWindow, 1, 3
    
    '**** COLONNE INFOS + LISTE ELEVE ****
    ' Taille ligne/colonne
    For intIndiceLigne = 1 To intNombreEleves + 3
        If intIndiceLigne < 4 Then
            Rows(intIndiceLigne).RowHeight = 25
        Else
            Rows(intIndiceLigne).RowHeight = 15
        End If
    Next intIndiceLigne
    Columns.ColumnWidth = 6
    Columns(1).ColumnWidth = 40
    
    ' Bouton 'actualiser résultats'
    Set rngCelluleBouton = Range("A1")
    Set btnBouton = ActiveSheet.Buttons.Add(rngCelluleBouton.Left, rngCelluleBouton.Top, rngCelluleBouton.Width, rngCelluleBouton.Height)
    With btnBouton
        .Caption = "Actualiser résultats"
        .OnAction = "btnActualiserResultats_Click"
    End With
    
    ' Légende
    With Range("A3")
        .Value = strNomClasse
        .Interior.ColorIndex = intColorClasse
        .Borders.ColorIndex = xlColorIndexAutomatic
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
    End With
    
    ' Liste élève
    For intIndiceEleve = 1 To intNombreEleves
        With Cells(3 + intIndiceEleve, 1)
            .Value = Sheets(strPage2).Cells(3 + intIndiceEleve, intIndiceClasse * 2 - 1).Value
        End With
    Next intIndiceEleve
    With Range(Cells(4, 1), Cells(3 + intNombreEleves, 1))
        .HorizontalAlignment = xlHAlignLeft
        .Borders.ColorIndex = xlColorIndexAutomatic
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    '**** LIGNE EN-TETE + CONTENU ****
    intNombreDomaines = getNombreDomaines
    With Range(Cells(1, 2), Cells(1, 1 + 4 * (intNombreDomaines + 1)))
        .Interior.ColorIndex = intColorBilan
        .MergeCells = True
        .Value = "Bilan trimestriel & annuel"
    End With
    For intIndiceDomaine = 1 To intNombreDomaines + 1
        If intIndiceDomaine <= intNombreDomaines Then
            With Range(Cells(2, 2 + 4 * (intIndiceDomaine - 1)), Cells(2, 5 + 4 * (intIndiceDomaine - 1)))
                .Interior.ColorIndex = intColorDomaine
                .MergeCells = True
                .Value = "D" & intIndiceDomaine
            End With
            Range(Cells(3, 5 + 4 * (intIndiceDomaine - 1)), Cells(3 + intNombreEleves, 5 + 4 * (intIndiceDomaine - 1))).Interior.ColorIndex = intColorDomaine2
            
        Else
            With Range(Cells(2, 2 + 4 * (intIndiceDomaine - 1)), Cells(2, 5 + 4 * (intIndiceDomaine - 1)))
                .Interior.ColorIndex = intColorNote
                .MergeCells = True
                .Value = "Note globale"
            End With
            Range(Cells(3, 5 + 4 * (intIndiceDomaine - 1)), Cells(3 + intNombreEleves, 5 + 4 * (intIndiceDomaine - 1))).Interior.ColorIndex = intColorNote2
        End If
        Cells(3, 2 + 4 * (intIndiceDomaine - 1)).Value = "1e tri"
        Cells(3, 3 + 4 * (intIndiceDomaine - 1)).Value = "2e tri"
        Cells(3, 4 + 4 * (intIndiceDomaine - 1)).Value = "3e tri"
        Cells(3, 5 + 4 * (intIndiceDomaine - 1)).Value = "Année"
        With Range(Cells(2, 2 + 4 * (intIndiceDomaine - 1)), Cells(3 + intNombreEleves, 5 + 4 * (intIndiceDomaine - 1)))
            .Borders.ColorIndex = xlColorIndexAutomatic
            .Borders.LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlMedium
            .Borders(xlEdgeTop).Weight = xlMedium
            .Borders(xlEdgeLeft).Weight = xlMedium
            .Borders(xlEdgeRight).Weight = xlMedium
            .Borders(xlInsideHorizontal).Weight = xlThin
            .Borders(xlInsideVertical).Weight = xlHairline
        End With
    Next intIndiceDomaine
    Range(Cells(1, 2), Cells(3 + intNombreEleves, 1 + 4 * (intNombreDomaines + 1))).BorderAround xlDouble, xlThin, xlColorIndexAutomatic
    Range(Cells(4, 1), Cells(3 + intNombreEleves, 1 + 4 * (intNombreDomaines + 1))).BorderAround xlDouble, xlThin, xlColorIndexAutomatic
    
    Range(Cells(4, 2), Cells(3 + intNombreEleves, 1 + 4 * (intNombreDomaines + 1))).Cells.Locked = False
    
    ' Protection feuille
    ActiveSheet.EnableSelection = xlUnlockedCells
    ActiveSheet.Protect strPassword
    Application.ScreenUpdating = True
    
End Sub

Sub btnActualiserResultats_Click()
    Dim strNomClasse As String
    Dim intNombreDomaines As Integer, intIndiceDomaine As Integer
    Dim intIndiceTrimestre As Integer
    Dim intNombreEvals As Integer, intIndiceEval As Integer
    Dim shtPage3 As Worksheet, shtPage4 As Worksheet
    Dim intAvancementTotal As Integer, intAvancementActuel As Integer

    ' Valeurs nécessaires
    strNomClasse = Range("A3").Value
    Set shtPage3 = Sheets("Notes (" & strNomClasse & ")")
    Set shtPage4 = Sheets("Bilan (" & strNomClasse & ")")
    intNombreDomaines = getNombreDomaines
    intNombreEvals = getNombreEvals(strNomClasse)
    
    ' UserForm5: Avancement de l'opération
    intAvancementActuel = 0
    intAvancementTotal = intNombreEvals + 4 * intNombreDomaines
    UserForm5.Show vbModeless

    ' Retrait protection page notes
    Application.ScreenUpdating = False
    shtPage3.Unprotect strPassword
    shtPage4.Unprotect strPassword

    ' Calcul des résultats
    For intIndiceEval = 1 To intNombreEvals
        calculNote strNomClasse, intIndiceEval
        intAvancementActuel = intAvancementActuel + 1
        UserForm5.updateAvancement intAvancementActuel, intAvancementTotal
    Next intIndiceEval
    For intIndiceTrimestre = 1 To 4
        For intIndiceDomaine = 1 To intNombreDomaines
            calculMoyenneDomaine strNomClasse, intIndiceDomaine, intIndiceTrimestre
            intAvancementActuel = intAvancementActuel + 1
            UserForm5.updateAvancement intAvancementActuel, intAvancementTotal
        Next intIndiceDomaine
        calculMoyenneTrimestre strNomClasse, intIndiceTrimestre
    Next intIndiceTrimestre
    
    ' Protection page notes
    shtPage3.Protect strPassword
    shtPage4.Protect strPassword
    Application.ScreenUpdating = True
    
    ' Opérations de fin de procédure
    UserForm5.Hide
    shtPage4.Activate
    MsgBox ("Données mises à jour.")
End Sub

' Calcul de la moyenne trimestrielle/annuelle pour chaque domaine
' intIndiceTrimestre = 4 pour indiquer l'année
Sub calculMoyenneDomaine(strNomClasse As String, intNumeroDomaine As Integer, intIndiceTrimestre As Integer)
    Dim intNombreEleves As Integer, intIndiceEleve As Integer
    Dim intNombreDomaines As Integer, intIndiceDomaine As Integer
    Dim intTotalCompetences As Integer, intMoitieTotalCompetences As Integer, intNumeroTotalCompetences As Integer
    Dim intNombreCompetences As Integer, intIndiceCompetence As Integer
    Dim intNombreEvals As Integer, intIndiceEval As Integer
    Dim strLettre As String
    Dim dblSomme As Double, dblDiviseur As Double, dblCoeffCompetence As Double
    Dim shtPage3 As Worksheet, shtPage4 As Worksheet

    ' Valeurs nécessaires
    Set shtPage3 = Sheets("Notes (" & strNomClasse & ")")
    Set shtPage4 = Sheets("Bilan (" & strNomClasse & ")")
    intNombreEvals = getNombreEvals(strNomClasse)
    intNombreCompetences = getNombreCompetences(intNumeroDomaine)
    intTotalCompetences = getNombreCompetences
    intNombreEleves = getNombreEleves(strNomClasse)
    intNombreDomaines = getNombreDomaines
    intMoitieTotalCompetences = (intTotalCompetences - intTotalCompetences Mod 2) / 2
    intNumeroTotalCompetences = 1
    
    ' Vérfication des entrées
    If intNumeroDomaine <= intNombreDomaines And (intIndiceTrimestre >= 1 And intIndiceTrimestre <= 4) Then
        
        ' Calcul intNumeroTotalCompetences = intIndiceColonneDepart du domaine concerné
        For intIndiceDomaine = 1 To intNumeroDomaine
            If intIndiceDomaine <> intNumeroDomaine Then
                intNumeroTotalCompetences = intNumeroTotalCompetences + getNombreCompetences(intIndiceDomaine)
            End If
        Next intIndiceDomaine
        
        ' Calcul de la moyenne
        For intIndiceEleve = 1 To intNombreEleves
            dblSomme = 0
            dblDiviseur = 0
            For intIndiceEval = 1 To intNombreEvals
                If intIndiceTrimestre = 4 Or shtPage3.Cells(2, 3 + (intIndiceEval - 1) * (intTotalCompetences + 1)).Value = intIndiceTrimestre Then
                    For intIndiceCompetence = intNumeroTotalCompetences To intNumeroTotalCompetences + intNombreCompetences - 1
                        strLettre = shtPage3.Cells(5 + intIndiceEleve, 2 + (intIndiceEval - 1) * (intTotalCompetences + 1) + intIndiceCompetence).Value
                        dblCoeffCompetence = shtPage3.Cells(5, 2 + (intIndiceEval - 1) * (intTotalCompetences + 1) + intIndiceCompetence).Value
                        If StrComp(strLettre, vbNullString) <> 0 And IsEmpty(dblCoeffCompetence) = False Then
                            dblSomme = dblSomme + dblCoeffCompetence * lettreToValeur(strLettre)
                            dblDiviseur = dblDiviseur + dblCoeffCompetence
                        End If
                    Next intIndiceCompetence
                End If
            Next intIndiceEval
            If dblSomme <> 0 Then
                shtPage4.Cells(3 + intIndiceEleve, 1 + 4 * (intNumeroDomaine - 1) + intIndiceTrimestre).Value = valeurToLettre(dblSomme / dblDiviseur)
            ElseIf dblSomme = 0 And dblDiviseur = 0 Then
                shtPage4.Cells(3 + intIndiceEleve, 1 + 4 * (intNumeroDomaine - 1) + intIndiceTrimestre).Value = vbNullString
            End If
        Next intIndiceEleve
    End If
    
    Set shtPage3 = Nothing
    Set shtPage4 = Nothing
    
End Sub

' Calcul la moyenne des notes du trimestre
Sub calculMoyenneTrimestre(strNomClasse As String, intIndiceTrimestre As Integer)
    Dim intNombreEleves As Integer, intIndiceEleve As Integer
    Dim intNombreDomaines As Integer
    Dim intTotalCompetences As Integer, intMoitieTotalCompetences As Integer
    Dim intNombreEvals As Integer, intIndiceEval As Integer
    Dim varNote As Variant, dblSomme As Double, dblDiviseur As Double, dblCoeffEval As Double
    Dim shtPage3 As Worksheet, shtPage4 As Worksheet

    ' Valeurs nécessaires
    Set shtPage3 = Sheets("Notes (" & strNomClasse & ")")
    Set shtPage4 = Sheets("Bilan (" & strNomClasse & ")")
    intNombreEvals = getNombreEvals(strNomClasse)
    intTotalCompetences = getNombreCompetences
    intNombreEleves = getNombreEleves(strNomClasse)
    intNombreDomaines = getNombreDomaines
    intMoitieTotalCompetences = (intTotalCompetences - intTotalCompetences Mod 2) / 2
    
    ' Calcul de la moyenne
    For intIndiceEleve = 1 To intNombreEleves
        dblSomme = 0
        dblDiviseur = 0
        For intIndiceEval = 1 To intNombreEvals
            If intIndiceTrimestre = 4 Or shtPage3.Cells(2, 3 + (intIndiceEval - 1) * (intTotalCompetences + 1)).Value = intIndiceTrimestre Then
                varNote = shtPage3.Cells(5 + intIndiceEleve, 2 + (intIndiceEval) * (intTotalCompetences + 1)).Value
                If Not IsEmpty(varNote) Then
                    dblCoeffEval = shtPage3.Cells(2, 3 + intMoitieTotalCompetences + (intIndiceEval - 1) * (intTotalCompetences + 1)).Value
                    dblSomme = dblSomme + dblCoeffEval * CDbl(varNote)
                    dblDiviseur = dblDiviseur + dblCoeffEval
                End If
            End If
        Next intIndiceEval
        If dblSomme <> 0 Then
            shtPage4.Cells(3 + intIndiceEleve, 1 + 4 * intNombreDomaines + intIndiceTrimestre).Value = Format(dblSomme / dblDiviseur, "Standard")
        ElseIf dblSomme = 0 And dblDiviseur = 0 Then
            shtPage4.Cells(3 + intIndiceEleve, 1 + 4 * intNombreDomaines + intIndiceTrimestre).Value = vbNullString
        End If
    Next intIndiceEleve
    
    Set shtPage3 = Nothing
    Set shtPage4 = Nothing
    
End Sub

