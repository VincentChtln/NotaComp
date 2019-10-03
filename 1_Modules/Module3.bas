Attribute VB_Name = "Module3"
' ##################################
' PAGE 3 (entrée notes)
' ##################################

Option Explicit

' **********************************
' FONCTIONS
' lettreToValeur(lettre As String) As Integer
' valeurToLettre(valeur As Integer) As String
' getNombreEvals(strNomClasse As String) As Integer
' **********************************

Function lettreToValeur(strLettre As String) As Integer
    Dim intAsciiLettre As String
    intAsciiLettre = Asc(strLettre)
    If intAsciiLettre > 64 And intAsciiLettre < 70 Then
        lettreToValeur = 69 - intAsciiLettre
    Else
        lettreToValeur = 0
    End If
End Function

Function valeurToLettre(intValeur As Integer) As String
    If intValeur >= 0 And intValeur <= 4 Then
        Select Case intValeur
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

Function getNombreEvals(strNomClasse As String) As Integer
    getNombreEvals = Sheets("Notes (" & strNomClasse & ")").Buttons.Count - 1
End Function

' **********************************
' PROCÉDURES
' **********************************

Sub creerTableauNotes(intIndiceClasse As Integer, intNombreEleves As Integer)
    Dim intIndiceLigne As Integer, intIndiceColonne As Integer, intIndiceColonneDepart As Integer
    Dim rngCelluleBouton As Range, btnBouton As Variant
    Dim intIndiceEleve As Integer
    Dim strNomClasse As String
    
    strNomClasse = getNomClasse(intIndiceClasse)
    
    ' Deverouillage
    Application.ScreenUpdating = False
    ActiveWorkbook.Unprotect strPassword
    
    ' Creation page
    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Name = "Notes (" & strNomClasse & ")"
    ActiveWorkbook.Protect strPassword, True, True
    With Cells
        .Borders.ColorIndex = 2
        .Locked = True
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
    End With
    
    '**** COLONNE INFOS + LISTE ELEVE ****
    ' Taille ligne/colonne
    For intIndiceLigne = 1 To intNombreEleves + 5
        If intIndiceLigne < 4 Then
            Rows(intIndiceLigne).RowHeight = 20
        ElseIf intIndiceLigne = 4 Or intIndiceLigne = 5 Then
            Rows(intIndiceLigne).RowHeight = 30
        Else
            Rows(intIndiceLigne).RowHeight = 15
        End If
    Next intIndiceLigne
    For intIndiceColonne = 1 To 2
        Columns(intIndiceColonne).ColumnWidth = 25
    Next intIndiceColonne
    
    ' Bouton 'ajouter éval'
    Set rngCelluleBouton = Range("A1")
    Set btnBouton = ActiveSheet.Buttons.Add(rngCelluleBouton.Left, rngCelluleBouton.Top, rngCelluleBouton.Width, rngCelluleBouton.Height)
    With btnBouton
        .Caption = "Ajouter évaluation"
        .OnAction = "btnAjouterEvaluation_Click"
    End With
    
    ' Légende
    With Range("A5")
        .Value = strNomClasse
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
    For intIndiceEleve = 1 To intNombreEleves
        With Cells(5 + intIndiceEleve, 1)
            .Value = Sheets(strPage2).Cells(3 + intIndiceEleve, intIndiceClasse * 2 - 1).Value
        End With
        Range(Cells(5 + intIndiceEleve, 1), Cells(5 + intIndiceEleve, 2)).MergeCells = True
    Next intIndiceEleve
    With Range(Cells(6, 1), Cells(5 + intNombreEleves, 2))
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
    intIndiceColonneDepart = 3
    
    ajouterEvaluation (intIndiceColonneDepart)
    
    
    ' Verouillage
    ActiveSheet.EnableSelection = xlUnlockedCells
    ActiveSheet.Protect strPassword
    Application.ScreenUpdating = True
    
    
End Sub

Sub btnAjouterEvaluation_Click()
    Dim intNombreEleves As Integer
    Dim intNombreDomaines As Integer
    Dim intNombreCompetences As Integer
    Dim intIndiceColonneDepart As Integer
    Dim intIndiceEval As Integer
    
    intNombreDomaines = getNombreDomaines
    intNombreEleves = getNombreEleves(ActiveSheet.Cells(5, 1).Value)
        
    ' Determine la colonne où ajouter l'éval
    intIndiceEval = ActiveSheet.Buttons.Count - 1
    intNombreCompetences = getNombreCompetences
    intIndiceColonneDepart = 3 + intIndiceEval * (intNombreCompetences + 1)
    
    ' Retrait protection feuille
    Application.ScreenUpdating = False
    ActiveSheet.Unprotect strPassword
    
    ' Ajout de l'évaluation
    ajouterEvaluation (intIndiceColonneDepart)
        
    ' Protection feuille
    ActiveSheet.EnableSelection = xlUnlockedCells
    ActiveSheet.Protect strPassword
    Application.ScreenUpdating = True
    
    MsgBox ("Évaluation ajoutée")

End Sub

Sub ajouterEvaluation(intIndiceColonneDepart As Integer)
    Dim intNombreEleves As Integer
    Dim intNombreDomaines As Integer, intIndiceDomaine As Integer
    Dim intTotalCompetences As Integer, intMoitieTotalCompetences As Integer, intIndiceCompetence As Integer, intNombreCompetences As Integer
    Dim rngCelluleBouton As Range, btnBouton As Variant
    
    ' Calcul données nécessaires
    intNombreDomaines = getNombreDomaines
    intNombreEleves = getNombreEleves(ActiveSheet.Cells(5, 1).Value)
    
    ' Domaines/Compétences
    intTotalCompetences = 0
    For intIndiceDomaine = 1 To intNombreDomaines
        intNombreCompetences = getNombreCompetences(intIndiceDomaine)
        For intIndiceCompetence = 1 To intNombreCompetences
            intTotalCompetences = intTotalCompetences + 1
            Columns(intIndiceColonneDepart + intTotalCompetences - 1).ColumnWidth = 3
            With Cells(4, intIndiceColonneDepart + intTotalCompetences - 1)
                .Value = "D" & intIndiceDomaine & "/" & intIndiceCompetence
                .Orientation = xlUpward
                .Interior.ColorIndex = intColorDomaine2
            End With
        Next intIndiceCompetence
        Cells(3, intIndiceColonneDepart + intTotalCompetences - intIndiceCompetence + 1).Value = "D" & intIndiceDomaine
        Range(Cells(3, intIndiceColonneDepart + intTotalCompetences - intIndiceCompetence + 1), Cells(3, intIndiceColonneDepart + intTotalCompetences - 1)).Interior.ColorIndex = intColorDomaine
        Range(Cells(3, intIndiceColonneDepart + intTotalCompetences - intIndiceCompetence + 1), Cells(3, intIndiceColonneDepart + intTotalCompetences - 1)).MergeCells = True
        With Range(Cells(3, intIndiceColonneDepart + intTotalCompetences - intIndiceCompetence + 1), Cells(5 + intNombreEleves, intIndiceColonneDepart + intTotalCompetences - 1))
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
    
    ' Infos Eval
    intMoitieTotalCompetences = (intTotalCompetences - intTotalCompetences Mod 2) / 2
    Range(Cells(1, intIndiceColonneDepart), Cells(1, intIndiceColonneDepart + intTotalCompetences - 1)).Interior.ColorIndex = intColorEval
    Range(Cells(1, intIndiceColonneDepart), Cells(1, intIndiceColonneDepart + intTotalCompetences - 1)).MergeCells = True
    Range(Cells(2, intIndiceColonneDepart), Cells(2, intIndiceColonneDepart + intMoitieTotalCompetences - 1)).MergeCells = True
    Range(Cells(2, intIndiceColonneDepart + intMoitieTotalCompetences), Cells(2, intIndiceColonneDepart + intTotalCompetences - 1)).MergeCells = True
    Set rngCelluleBouton = Range(Cells(1, intIndiceColonneDepart + intTotalCompetences), Cells(2, intIndiceColonneDepart + intTotalCompetences))
    Set btnBouton = ActiveSheet.Buttons.Add(rngCelluleBouton.Left, rngCelluleBouton.Top, rngCelluleBouton.Width, rngCelluleBouton.Height)
    With btnBouton
        .Caption = "Calcul note"
        .OnAction = "btnCalculNote_Click"
    End With
    With Range(Cells(3, intIndiceColonneDepart + intTotalCompetences), Cells(5, intIndiceColonneDepart + intTotalCompetences))
        .Interior.ColorIndex = intColorNote
        .MergeCells = True
        .Orientation = xlUpward
        .Value = "Note / 20"
        .Columns.ColumnWidth = 6
    End With
    Range(Cells(6, intIndiceColonneDepart + intTotalCompetences), Cells(5 + intNombreEleves, intIndiceColonneDepart + intTotalCompetences)).Interior.ColorIndex = intColorNote2
    With Range(Cells(3, intIndiceColonneDepart + intTotalCompetences), Cells(5 + intNombreEleves, intIndiceColonneDepart + intTotalCompetences))
        .Borders.ColorIndex = xlColorIndexAutomatic
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlInsideHorizontal).Weight = xlThin
    End With
    With Range(Cells(1, intIndiceColonneDepart), Cells(2, intIndiceColonneDepart + intTotalCompetences - 1))
        .Borders.ColorIndex = xlColorIndexAutomatic
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeRight).Weight = xlMedium
    End With
    Range(Cells(1, intIndiceColonneDepart), Cells(5 + intNombreEleves, intIndiceColonneDepart + intTotalCompetences)).BorderAround xlDouble, xlThin, xlColorIndexAutomatic
    Range(Cells(6, 1), Cells(5 + intNombreEleves, intIndiceColonneDepart + intTotalCompetences)).BorderAround xlDouble, xlThin, xlColorIndexAutomatic
    Range(Cells(1, intIndiceColonneDepart), Cells(2, intIndiceColonneDepart + intTotalCompetences - 1)).Locked = False
    Range(Cells(5, intIndiceColonneDepart), Cells(5 + intNombreEleves, intIndiceColonneDepart + intTotalCompetences - 1)).Locked = False
    
End Sub

Sub btnCalculNote_Click()
    Dim intIndiceEval As Integer
    Dim intNombreCompetences As Integer
    Dim strNomClasse As String

    ' Determiner éval à calculer
    intIndiceEval = Val(Right(Application.Caller, 1)) - 1
    intNombreCompetences = getNombreCompetences
    strNomClasse = ActiveSheet.Range("A5").Value
    
    ' Retrait protection feuille
    Application.ScreenUpdating = False
    ActiveSheet.Unprotect strPassword
    
    ' Calcul note éval
    calculNote strNomClasse, intIndiceEval
        
    ' Protection feuille
    ActiveSheet.EnableSelection = xlUnlockedCells
    ActiveSheet.Protect strPassword
    Application.ScreenUpdating = True
    
End Sub

Sub calculNote(strNomClasse As String, intIndiceEval As Integer)
    Dim shtPage3 As Worksheet
    Dim strLettre As String
    Dim intNombreEleves As Integer, intIndiceEleve As Integer
    Dim intNombreCompetences As Integer, intIndiceCompetence As Integer, dblCoeffCompetence As Double
    Dim dblSomme As Double, dblDiviseur As Double
    Dim intIndiceColonneEval As Integer

    ' Calcul données nécessaires
    '@Ignore ImplicitActiveWorkbookReference
    Set shtPage3 = Sheets("Notes (" & strNomClasse & ")")
    intNombreEleves = getNombreEleves(strNomClasse)
    intNombreCompetences = getNombreCompetences
    intIndiceColonneEval = 3 + (intIndiceEval - 1) * (intNombreCompetences + 1)
    
    ' Calcul note éval
    For intIndiceEleve = 1 To intNombreEleves
        dblDiviseur = 0
        dblSomme = 0
        For intIndiceCompetence = 1 To intNombreCompetences
            strLettre = shtPage3.Cells(5 + intIndiceEleve, intIndiceColonneEval + intIndiceCompetence - 1).Value
            dblCoeffCompetence = shtPage3.Cells(5, intIndiceColonneEval + intIndiceCompetence - 1).Value
            If StrComp(strLettre, vbNullString) <> 0 And Not IsEmpty(dblCoeffCompetence) Then
                dblSomme = dblSomme + lettreToValeur(strLettre) * dblCoeffCompetence
                dblDiviseur = dblDiviseur + dblCoeffCompetence
            End If
        Next intIndiceCompetence
        If dblDiviseur <> 0 Then
            shtPage3.Cells(5 + intIndiceEleve, intIndiceColonneEval + intNombreCompetences).Value = Format(5# * dblSomme / dblDiviseur, "Standard")
        ElseIf dblSomme = 0 And dblDiviseur = 0 Then
            shtPage3.Cells(5 + intIndiceEleve, intIndiceColonneEval + intNombreCompetences).Value = vbNullString
        End If
    Next intIndiceEleve
    
    Set shtPage3 = Nothing
End Sub

