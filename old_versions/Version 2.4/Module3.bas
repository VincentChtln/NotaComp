Attribute VB_Name = "Module3"
' *******************************************************************************
'
'                               Page 3 - Notes
'
'   Fonctions
'       lettreToValeur(lettre As String) As Integer
'       valeurToLettre(valeur As Integer) As String
'       getNombreEvals(strNomClasse As String) As Integer
'
'   Procédures
'       creerTableauNotes(intIndiceClasse As Integer, intNbEleves As Integer)
'       btnAjouterEvaluation_Click()
'       ajouterEvaluation(intIndiceColEval As Integer)
'       btnCalculNote_Click()
'       calculNote(strNomClasse As String, intIndiceEval As Integer)
'
' *******************************************************************************

Option Explicit

' *******************************************************************************
'                                   Fonctions
' *******************************************************************************

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
    ' *** CALCUL ***
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
        Case Is = 0
            valeurToLettre = "E"
        End Select
    Else
        valeurToLettre = "Z"
    End If
End Function

Function getNombreEvals(intIndiceClasse As Integer) As Integer
    getNombreEvals = Worksheets(getNomPage3(intIndiceClasse)).Buttons.Count - 1
End Function

' *******************************************************************************
'                                   Procédures
' *******************************************************************************

Sub creerTableauNotes(intIndiceClasse As Integer, intNbEleves As Integer)
    ' *** DECLARATION VARIABLES ***
    Dim strNomClasse As String
    Dim intIndiceEleve As Integer
    Dim intIndiceLigne As Integer
    Dim intIndiceColonne As Integer
    Dim intIndiceColEval As Integer
    Dim rngBoutonAjoutEval As Range
    Dim btnAjoutEval As Variant
    Dim strPage3 As String
    
    ' *** AFFECTATION VARIABLES ***
    strNomClasse = getNomClasse(intIndiceClasse)
    strPage3 = getNomPage3(intIndiceClasse)
    
    ' *** CREATION PAGE ***
    Worksheets.Add After:=Worksheets(Worksheets.Count)
    Worksheets(Worksheets.Count).Name = strPage3
    With Worksheets(strPage3)
        With .Cells
            .Borders.ColorIndex = 2
            .Locked = True
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
        End With
        
        ' *** FORMATAGE TAILLE LIGNES + COLONNES ***
        For intIndiceLigne = 1 To intLigListePage3 + intNbEleves
            If intIndiceLigne < 4 Then
                .Rows(intIndiceLigne).RowHeight = 20
            ElseIf intIndiceLigne = 4 Or intIndiceLigne = 5 Then
                .Rows(intIndiceLigne).RowHeight = 30
            Else
                .Rows(intIndiceLigne).RowHeight = 15
            End If
        Next intIndiceLigne
        For intIndiceColonne = 1 To 2
            .Columns(intIndiceColonne).ColumnWidth = 25
        Next intIndiceColonne
        
        ' *** BOUTON 'AJOUTER EVAL' ***
        Set rngBoutonAjoutEval = .Range("A1:A2")
        Set btnAjoutEval = .Buttons.Add(rngBoutonAjoutEval.Left, rngBoutonAjoutEval.Top, rngBoutonAjoutEval.Width, rngBoutonAjoutEval.Height)
        With btnAjoutEval
            .Caption = "Ajouter évaluation"
            .OnAction = "btnAjouterEvaluation_Click"
        End With
        
        ' *** LEGENDE ***
        With .Range("A3:A5")
            .Interior.ColorIndex = intColorClasse
            .Borders.ColorIndex = xlColorIndexAutomatic
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlMedium
            .MergeCells = True
            .Value = strNomClasse
        End With
        .Range("B1").Value = "Nom de l'évaluation"
        .Range("B2").Value = "Trimestre / Coeff"
        .Range("B3").Value = "Domaines"
        .Range("B4").Value = "Compétences"
        .Range("B5").Value = "Coeff compétence"
        With .Range("B1:B2,B3:B5")
            .Borders.ColorIndex = xlColorIndexAutomatic
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders(xlEdgeBottom).Weight = xlMedium
            .Borders(xlEdgeTop).Weight = xlMedium
            .Borders(xlEdgeLeft).Weight = xlMedium
            .Borders(xlEdgeRight).Weight = xlMedium
        End With
        .Range("B1").Interior.ColorIndex = intColorEval
        .Range("B3").Interior.ColorIndex = intColorDomaine
        .Range("B4").Interior.ColorIndex = intColorDomaine2
        
        ' *** LISTE ELEVES ***
        For intIndiceEleve = 1 To intNbEleves
            .Cells(intLigListePage3 + intIndiceEleve, 1).Value = Worksheets(strPage2).Cells(intLigListePage2 + intIndiceEleve, intIndiceClasse * 2 - 1).Value
            .Range(.Cells(intLigListePage3 + intIndiceEleve, 1), .Cells(intLigListePage3 + intIndiceEleve, 2)).MergeCells = True
        Next intIndiceEleve
        With .Range(.Cells(intLigListePage3 + 1, 1), .Cells(intLigListePage3 + intNbEleves, 2))
            .HorizontalAlignment = xlHAlignLeft
            .Borders.ColorIndex = xlColorIndexAutomatic
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
    End With
    
    ' *** FIGEAGE VOLETS ***
    freezePanes ActiveWindow, intLigListePage3, 2
    
    ' *** AJOUT 1e EVALUATION ***
    intIndiceColEval = 3
    ajouterEvaluation intIndiceClasse, intIndiceColEval
    
End Sub

Sub btnAjouterEvaluation_Click()
    ' *** DECLARATION VARIABLES ***
    Dim intNbCompetences As Integer
    Dim intIndiceEval As Integer
    Dim intIndiceColEval As Integer
    Dim intIndiceClasse As Integer
        
    ' *** AFFECTATION VARIABLES ***
    intIndiceClasse = getIndiceClasse(ActiveSheet.Name)
    intIndiceEval = ActiveSheet.Buttons.Count - 1
    intNbCompetences = getNombreCompetences
    intIndiceColEval = 3 + intIndiceEval * (intNbCompetences + 1)
    
    ' *** PROTECTION + REFRESH ECRAN OFF ***
    Application.ScreenUpdating = False
    unprotectWorksheet intIndiceClasse
    
    ' *** AJOUT EVALUATION ***
    ajouterEvaluation intIndiceClasse, intIndiceColEval
        
    ' *** PROTECTION + REFRESH ECRAN ON ***
    protectWorksheet intIndiceClasse
    Application.ScreenUpdating = True
    
    ' *** MESSAGE INFORMATION ***
    MsgBox ("Évaluation ajoutée")

End Sub

Sub ajouterEvaluation(intIndiceClasse As Integer, intIndiceColEval As Integer)
    ' *** DECLARATION VARIABLES ***
    Dim intNbEleves As Integer
    Dim intNbDomaines As Integer
    Dim intIndiceDomaine As Integer
    Dim intSommeCompetences As Integer
    Dim intMoitieTotalCompetences As Integer
    Dim intIndiceCompetence As Integer
    Dim intNbCompetences As Integer
    Dim rngCelluleBouton As Range
    Dim btnBouton As Variant
    Dim strPage3 As String
    
    
    ' *** AFFECTATION VARIABLES ***
    intNbDomaines = getNombreDomaines
    strPage3 = getNomPage3(intIndiceClasse)
    intSommeCompetences = 0
    
    With Worksheets(strPage3)
        intNbEleves = getNombreEleves(intIndiceClasse)
        
        ' *** FORMATAGE ZONE DOMAINES/COMPETENCES ***
        For intIndiceDomaine = 1 To intNbDomaines
            intNbCompetences = getNombreCompetences(intIndiceDomaine)
            For intIndiceCompetence = 1 To intNbCompetences
                intSommeCompetences = intSommeCompetences + 1
                .Columns(intIndiceColEval + intSommeCompetences - 1).ColumnWidth = 3
                With .Cells(4, intIndiceColEval + intSommeCompetences - 1)
                    .Value = "D" & intIndiceDomaine & "/" & intIndiceCompetence
                    .Orientation = xlUpward
                    .Interior.ColorIndex = intColorDomaine2
                End With
            Next intIndiceCompetence
            .Cells(3, intIndiceColEval + intSommeCompetences - intIndiceCompetence + 1).Value = "D" & intIndiceDomaine
            .Range(.Cells(3, intIndiceColEval + intSommeCompetences - intIndiceCompetence + 1), .Cells(3, intIndiceColEval + intSommeCompetences - 1)).Interior.ColorIndex = intColorDomaine
            .Range(.Cells(3, intIndiceColEval + intSommeCompetences - intIndiceCompetence + 1), .Cells(3, intIndiceColEval + intSommeCompetences - 1)).MergeCells = True
            With .Range(.Cells(3, intIndiceColEval + intSommeCompetences - intIndiceCompetence + 1), .Cells(5 + intNbEleves, intIndiceColEval + intSommeCompetences - 1))
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
        
        ' *** FORMATAGE ZONE INFOS EVAL ***
        intMoitieTotalCompetences = (intSommeCompetences - intSommeCompetences Mod 2) / 2
        .Range(.Cells(1, intIndiceColEval), .Cells(1, intIndiceColEval + intSommeCompetences - 1)).Interior.ColorIndex = intColorEval
        .Range(.Cells(1, intIndiceColEval), .Cells(1, intIndiceColEval + intSommeCompetences - 1)).MergeCells = True
        .Range(.Cells(2, intIndiceColEval), .Cells(2, intIndiceColEval + intMoitieTotalCompetences - 1)).MergeCells = True
        .Range(.Cells(2, intIndiceColEval + intMoitieTotalCompetences), .Cells(2, intIndiceColEval + intSommeCompetences - 1)).MergeCells = True
        
        ' *** FORMATAGE ZONE MOYENNE EVAL ***
        Set rngCelluleBouton = .Range(.Cells(1, intIndiceColEval + intSommeCompetences), .Cells(2, intIndiceColEval + intSommeCompetences))
        Set btnBouton = .Buttons.Add(rngCelluleBouton.Left, rngCelluleBouton.Top, rngCelluleBouton.Width, rngCelluleBouton.Height)
        With btnBouton
            .Caption = "Calcul note"
            .OnAction = "btnCalculNote_Click"
        End With
        With .Range(.Cells(3, intIndiceColEval + intSommeCompetences), .Cells(4, intIndiceColEval + intSommeCompetences))
            .Interior.ColorIndex = intColorNote
            .MergeCells = True
            .Orientation = xlUpward
            .Value = "Note / 20"
            .Columns.ColumnWidth = 6
        End With
        .Range(.Cells(intLigListePage3, intIndiceColEval + intSommeCompetences), .Cells(intLigListePage3 + intNbEleves, intIndiceColEval + intSommeCompetences)).Interior.ColorIndex = intColorNote2
        With .Range(.Cells(3, intIndiceColEval + intSommeCompetences), .Cells(intLigListePage3 + intNbEleves, intIndiceColEval + intSommeCompetences))
            .Borders.ColorIndex = xlColorIndexAutomatic
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders(xlEdgeBottom).Weight = xlMedium
            .Borders(xlEdgeTop).Weight = xlMedium
            .Borders(xlEdgeLeft).Weight = xlMedium
            .Borders(xlEdgeRight).Weight = xlMedium
            .Borders(xlInsideHorizontal).Weight = xlThin
        End With
        
        ' *** FORMATAGE ZONE RESULTATS ***
        With .Range(.Cells(1, intIndiceColEval), .Cells(2, intIndiceColEval + intSommeCompetences - 1))
            .Borders.ColorIndex = xlColorIndexAutomatic
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders(xlEdgeBottom).Weight = xlMedium
            .Borders(xlEdgeTop).Weight = xlMedium
            .Borders(xlEdgeLeft).Weight = xlMedium
            .Borders(xlEdgeRight).Weight = xlMedium
        End With
        .Range(.Cells(1, intIndiceColEval), .Cells(intLigListePage3 + intNbEleves, intIndiceColEval + intSommeCompetences)).BorderAround xlDouble, xlThin, xlColorIndexAutomatic
        .Range(.Cells(intLigListePage3 + 1, 1), .Cells(intLigListePage3 + intNbEleves, intIndiceColEval + intSommeCompetences)).BorderAround xlDouble, xlThin, xlColorIndexAutomatic
        .Range(.Cells(1, intIndiceColEval), .Cells(2, intIndiceColEval + intSommeCompetences - 1)).Locked = False
        .Range(.Cells(intLigListePage3, intIndiceColEval), .Cells(intLigListePage3 + intNbEleves, intIndiceColEval + intSommeCompetences - 1)).Locked = False
    End With
    
    ' *** LIMITATION ZONE SCROLL ***
    'limitScrollArea Worksheets(strPage3)
    
    
End Sub

Sub btnCalculNote_Click()
    ' *** DECLARATION VARIABLES ***
    Dim intIndiceEval As Integer
    Dim intIndiceClasse As Integer

    ' *** AFFECTATION VARIABLES ***
    intIndiceEval = Val(Split(Application.Caller, " ")(1)) - 1
    intIndiceClasse = getIndiceClasse(ActiveSheet.Name)
    
    ' *** PROTECTION + REFRESH ECRAN OFF ***
    UserForm5.Show vbModeless
    Application.ScreenUpdating = False
    unprotectWorksheet intIndiceClasse
    
    ' *** CALCUL NOTE ***
    calculNote intIndiceClasse, intIndiceEval
        
    ' *** PROTECTION + REFRESH ECRAN ON ***
    protectWorksheet intIndiceClasse
    Application.ScreenUpdating = True
    UserForm5.Hide
    
End Sub

Sub calculNote(intIndiceClasse As Integer, intIndiceEval As Integer)
    ' *** DECLARATION VARIABLES ***
    Dim strPage3 As String
    Dim strLettre As String
    Dim intColDebutEval As Integer
    Dim intIndiceEleve As Integer
    Dim intNbEleves As Integer
    Dim intIndiceCompetence As Integer
    Dim intNbCompetences As Integer
    Dim dblCoeffCompetence As Double
    Dim dblSommeEleve As Double
    Dim dblDiviseurEleve As Double
    Dim dblSommeClasse As Double
    Dim dblDiviseurClasse As Double

    ' *** AFFECTATION VARIABLES ***
    strPage3 = getNomPage3(intIndiceClasse)
    intNbEleves = getNombreEleves(intIndiceClasse)
    intNbCompetences = getNombreCompetences
    intColDebutEval = 3 + (intIndiceEval - 1) * (intNbCompetences + 1)
    dblSommeClasse = 0
    dblDiviseurClasse = 0
    
    ' *** CALCUL NOTE EVAL ***
    With Worksheets(strPage3)
        For intIndiceEleve = 1 To intNbEleves
            dblDiviseurEleve = 0
            dblSommeEleve = 0
            
            ' *** CALCUL AVEC DONNES PAGE 3 ***
            For intIndiceCompetence = 1 To intNbCompetences
                strLettre = .Cells(intLigListePage3 + intIndiceEleve, intColDebutEval + intIndiceCompetence - 1).Value
                dblCoeffCompetence = .Cells(intLigListePage3, intColDebutEval + intIndiceCompetence - 1).Value
                If StrComp(strLettre, vbNullString) <> 0 And Not IsEmpty(dblCoeffCompetence) Then
                    dblSommeEleve = dblSommeEleve + lettreToValeur(strLettre) * dblCoeffCompetence
                    dblDiviseurEleve = dblDiviseurEleve + dblCoeffCompetence
                End If
            Next intIndiceCompetence
            
            ' *** AFFECTATION RESULTAT PAGE 3 ***
            If dblDiviseurEleve <> 0 Then
                .Cells(intLigListePage3 + intIndiceEleve, intColDebutEval + intNbCompetences).Value = Format(5# * dblSommeEleve / dblDiviseurEleve, "Standard")
                dblSommeClasse = dblSommeClasse + 5# * dblSommeEleve / dblDiviseurEleve
                dblDiviseurClasse = dblDiviseurClasse + 1#
            ElseIf dblSommeEleve = 0 And dblDiviseurEleve = 0 Then
                .Cells(intLigListePage3 + intIndiceEleve, intColDebutEval + intNbCompetences).Value = vbNullString
            End If
            
            ' *** MAJ USERFORM AVANCEMENT ***
            UserForm5.updateAvancement intIndiceEleve, intNbEleves
        Next intIndiceEleve
        
        ' *** AFFECTATION MOYENNE PAGE 3 ***
        With .Cells(intLigListePage3, intColDebutEval + intNbCompetences)
            If dblDiviseurClasse <> 0 Then
                .Value = Format(dblSommeClasse / dblDiviseurClasse, "Standard")
            Else
                .Value = vbNullString
            End If
        End With
    End With
End Sub


