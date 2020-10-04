Attribute VB_Name = "Module4"
' ##################################
' PAGE 4 (résultats élèves)
' ##################################

Option Explicit

' ##################################
' FONCTIONS
' ##################################

' Aucune fonction

' ##################################
' PROCÉDURES
' ##################################
' creerTableauBilan(intIndiceClasse As Integer, intNbEleves As Integer)
' btnActualiserResultats_Click()
' calculMoyenneDomaine(intIndiceClasse As Integer, intIndiceDomaine As Integer, intIndiceTrimestre As Integer)
' calculMoyenneGlobale(intIndiceClasse, intIndiceTrimestre As Integer)
' ##################################

Sub creerTableauBilan(intIndiceClasse As Integer, intNbEleves As Integer)
    ' *** DECLARATION VARIABLES ***
    Dim ws4 As Worksheet
    Dim intIndiceLigne As Integer
    Dim intIndiceEleve As Integer
    Dim intIndiceDomaine As Integer
    Dim intNbDomaines As Integer
    Dim rngBtnActualiserResultats As Range
    Dim btnActualiserResultats As Variant

    ' *** AJOUT PAGE 4 - BILAN ***
    Worksheets.Add After:=Worksheets(Worksheets.Count)
    Worksheets(Worksheets.Count).Name = getNomPage4(intIndiceClasse)
    freezePanes ActiveWindow, intLigListePage4, 1
    
    ' *** AFFECTATION VARIABLES ***
    intNbDomaines = getNombreDomaines
    Set ws4 = Worksheets(getNomPage4(intIndiceClasse))
    
    With ws4
        ' *** FORMATAGE PAGE ***
        With .Cells
            .Borders.ColorIndex = 2
            .Locked = True
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
        End With
        
        ' *** FORMATAGE LIGNES + COLONNES ***
        For intIndiceLigne = 1 To intLigListePage4 + intNbEleves
            If intIndiceLigne <= intLigListePage4 Then
                .Rows(intIndiceLigne).RowHeight = 25
            Else
                .Rows(intIndiceLigne).RowHeight = 15
            End If
        Next intIndiceLigne
        .Columns.ColumnWidth = 6
        .Columns(1).ColumnWidth = 40
        
        ' *** CREATION BOUTON 'ACTUALISER RESULTATS' ***
        Set rngBtnActualiserResultats = .Range("A1")
        Set btnActualiserResultats = .Buttons.Add(rngBtnActualiserResultats.Left, rngBtnActualiserResultats.Top, rngBtnActualiserResultats.Width, rngBtnActualiserResultats.Height)
        With btnActualiserResultats
            .Caption = "Actualiser résultats"
            .OnAction = "btnActualiserResultats_Click"
        End With
        
        ' *** CELLULE AVEC NOM CLASSE ***
        With .Range("A2")
            .Value = getNomClasse(intIndiceClasse)
            .Interior.ColorIndex = intColorClasse
            .Borders.ColorIndex = xlColorIndexAutomatic
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlMedium
        End With
        .Range("A2:A3").MergeCells = True
        
        ' *** LISTE ELEVE ***
        For intIndiceEleve = 1 To intNbEleves
            .Cells(intLigListePage4 + intIndiceEleve, 1).Value = Worksheets(strPage2).Cells(intLigListePage2 + intIndiceEleve, intIndiceClasse * 2 - 1).Value
        Next intIndiceEleve
        With .Range(.Cells(intLigListePage4 + 1, 1), .Cells(intLigListePage4 + intNbEleves, 1))
            .HorizontalAlignment = xlHAlignLeft
            .Borders.ColorIndex = xlColorIndexAutomatic
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
        
        '**** LIGNE EN-TETE + CONTENU ****
        With .Range(.Cells(1, 2), .Cells(1, 1 + 4 * (intNbDomaines + 1)))
            .Interior.ColorIndex = intColorBilan
            .MergeCells = True
            .Value = "Bilan trimestriel & annuel"
        End With
        For intIndiceDomaine = 1 To intNbDomaines + 1
        
            ' *** FORMATAGE DOMAINES ***
            If intIndiceDomaine <= intNbDomaines Then
                With .Range(.Cells(2, 2 + 4 * (intIndiceDomaine - 1)), .Cells(2, 5 + 4 * (intIndiceDomaine - 1)))
                    .Interior.ColorIndex = intColorDomaine
                    .MergeCells = True
                    .Value = "D" & intIndiceDomaine
                End With
                .Range(.Cells(intLigListePage4, 5 + 4 * (intIndiceDomaine - 1)), .Cells(intLigListePage4 + intNbEleves, 5 + 4 * (intIndiceDomaine - 1))).Interior.ColorIndex = intColorDomaine2
            
            ' *** FORMATAGE MOYENNE GLOBALE ***
            Else
                With .Range(.Cells(2, 2 + 4 * (intIndiceDomaine - 1)), .Cells(2, 5 + 4 * (intIndiceDomaine - 1)))
                    .Interior.ColorIndex = intColorNote
                    .MergeCells = True
                    .Value = "Note globale"
                End With
                .Range(.Cells(intLigListePage4, 5 + 4 * (intIndiceDomaine - 1)), .Cells(intLigListePage4 + intNbEleves, 5 + 4 * (intIndiceDomaine - 1))).Interior.ColorIndex = intColorNote2
            End If
            
            ' *** LEGENDE LIGNE EN-TETE ***
            .Cells(intLigListePage4, 2 + 4 * (intIndiceDomaine - 1)).Value = "1e tri"
            .Cells(intLigListePage4, 3 + 4 * (intIndiceDomaine - 1)).Value = "2e tri"
            .Cells(intLigListePage4, 4 + 4 * (intIndiceDomaine - 1)).Value = "3e tri"
            .Cells(intLigListePage4, 5 + 4 * (intIndiceDomaine - 1)).Value = "Année"
            
            ' *** FORMATAGE TABLEAU RESULTATS
            With .Range(.Cells(2, 2 + 4 * (intIndiceDomaine - 1)), .Cells(intLigListePage4 + intNbEleves, 5 + 4 * (intIndiceDomaine - 1)))
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
        
        .Range(.Cells(1, 2), .Cells(intLigListePage4 + intNbEleves, 1 + 4 * (intNbDomaines + 1))).BorderAround xlDouble, xlThin, xlColorIndexAutomatic
        .Range(.Cells(intLigListePage4 + 1, 1), .Cells(intLigListePage4 + intNbEleves, 1 + 4 * (intNbDomaines + 1))).BorderAround xlDouble, xlThin, xlColorIndexAutomatic
        .Range(.Cells(intLigListePage4 + 1, 2), .Cells(intLigListePage4 + intNbEleves, 1 + 4 * (intNbDomaines + 1))).Cells.Locked = False
    End With
    
    ' *** LIMITATION ZONE SCROLL ***
    'limitScrollArea ws4
    
End Sub

Sub btnActualiserResultats_Click()
    ' *** DECLARATION VARIABLES ***
    Dim intIndiceClasse As Integer
    Dim intIndiceDomaine As Integer
    Dim intNbDomaines As Integer
    Dim intIndiceTrimestre As Integer
    Dim intIndiceEval As Integer
    Dim intNbEvals As Integer
    Dim strPage4 As String
    Dim intAvancementTotal As Integer
    Dim intAvancementActuel As Integer

    ' *** AFFECTATION VARIABLES ***
    intIndiceClasse = getIndiceClasse(ActiveSheet.Name)
    strPage4 = getNomPage4(intIndiceClasse)
    intNbDomaines = getNombreDomaines
    intNbEvals = getNombreEvals(intIndiceClasse)
    
    ' *** USERFORM 5 - AFFICHAGE AVANCEMENT ***
    intAvancementActuel = 0
    intAvancementTotal = intNbEvals + 4 * intNbDomaines
    UserForm5.Show vbModeless

    ' *** PROTECTION + REFRESH ECRAN OFF ***
    Application.ScreenUpdating = False
    unprotectWorksheet intIndiceClasse

    ' *** RECALCUL NOTES EVAL ***
    For intIndiceEval = 1 To intNbEvals
        calculNote intIndiceClasse, intIndiceEval
        intAvancementActuel = intAvancementActuel + 1
        UserForm5.updateAvancement intAvancementActuel, intAvancementTotal
    Next intIndiceEval
    
    ' *** CALCUL MOYENNES PAR DOMAINE ET PAR TRIMESTRE ***
    For intIndiceTrimestre = 1 To 4
        For intIndiceDomaine = 1 To intNbDomaines
            calculMoyenneDomaine intIndiceClasse, intIndiceDomaine, intIndiceTrimestre
            intAvancementActuel = intAvancementActuel + 1
            UserForm5.updateAvancement intAvancementActuel, intAvancementTotal
        Next intIndiceDomaine
        calculMoyenneGlobale intIndiceClasse, intIndiceTrimestre
    Next intIndiceTrimestre
    
    ' *** PROTECTION + REFRESH ECRAN ON ***
    UserForm5.Hide
    protectWorksheet intIndiceClasse
    Application.ScreenUpdating = True
    
    ' *** MESSAGE INFORMATION ***
    Worksheets(strPage4).Activate
    MsgBox ("Données mises à jour.")
End Sub

' Calcul de la moyenne trimestrielle/annuelle pour chaque domaine
' intIndiceTrimestre = 4 pour indiquer l'année
Sub calculMoyenneDomaine(intIndiceClasse As Integer, intIndiceDomaine As Integer, intIndiceTrimestre As Integer)
    ' *** DECLARATION VARIABLES ***
    Dim strNomClasse As String                  ' Nom de la classe
    Dim intIndiceEleve As Integer               ' Indice de l'élève actuel
    Dim intNbEleves As Integer                  ' Nombre total d'élèves dans la classe
    Dim intIndiceEval As Integer                ' Indice de l'éval actuelle
    Dim intNbEvals As Integer                   ' Nombre total d'évals effectuées par la classe
    Dim intIndiceBisDomaine As Integer          ' Indice du domaine actuel - utilisé pour la recherche intColDebutDomaine
    Dim intIndiceCompetence As Integer          ' Indice de la compétence actuelle
    Dim intNbCompetencesDomaine As Integer      '
    Dim intNbCompetences As Integer             '
    Dim intColDebutDomaine As Integer           ' Colonne de début de domaine (sur 1 éval)
    Dim intColFinDomaine As Integer             ' Colonne de fin de domaine (sur 1 éval)
    Dim strLettre As String                     ' Lettre de notation
    Dim dblSomme As Double                      ' Somme pondérée des notes d'un élève
    Dim dblDiviseur As Double                   ' Somme des coeff des évals
    Dim dblCoeffCompetence As Double            ' Coeff d'une compétence
    Dim strPage3 As String                      ' Nom de la page 3
    Dim strPage4 As String                      ' Nom de la page 4

    ' *** AFFECTATION VARIABLES ***
    strNomClasse = getNomClasse(intIndiceClasse)
    strPage3 = "Notes (" & strNomClasse & ")"
    strPage4 = "Bilan (" & strNomClasse & ")"
    intNbEvals = getNombreEvals(intIndiceClasse)
    intNbCompetencesDomaine = getNombreCompetences(intIndiceDomaine)
    intNbCompetences = getNombreCompetences
    intNbEleves = getNombreEleves(intIndiceClasse)
    intColDebutDomaine = 1
    
    ' *** CALCUL COLONNE DEBUT DOMAINE DANS EVAL ***
    intColDebutDomaine = 1
    For intIndiceBisDomaine = 1 To intIndiceDomaine
        If intIndiceBisDomaine <> intIndiceDomaine Then
            intColDebutDomaine = intColDebutDomaine + getNombreCompetences(intIndiceBisDomaine)
        End If
    Next intIndiceBisDomaine
    intColFinDomaine = intColDebutDomaine + intNbCompetencesDomaine - 1
    
    ' *** CALCUL MOYENNE DOMAINE TRIMESTRE ***
    For intIndiceEleve = 1 To intNbEleves
        dblSomme = 0
        dblDiviseur = 0
        
        ' *** CALCUL DONNEES PAGE 3 ***
        With Worksheets(strPage3)
            For intIndiceEval = 1 To intNbEvals
                ' *** SELECTION EVALUATION TRIMESTRE CONCERNE (OU ANNEE SI intIndiceTrimestre = 4) ***
                If intIndiceTrimestre = 4 Or .Cells(2, 3 + (intIndiceEval - 1) * (intNbCompetences + 1)).Value = intIndiceTrimestre Then
                    ' *** CALCUL SOMME PONDEREE ***
                    For intIndiceCompetence = intColDebutDomaine To intColFinDomaine
                        strLettre = .Cells(intLigListePage3 + intIndiceEleve, 2 + (intIndiceEval - 1) * (intNbCompetences + 1) + intIndiceCompetence).Value
                        dblCoeffCompetence = .Cells(intLigListePage3, 2 + (intIndiceEval - 1) * (intNbCompetences + 1) + intIndiceCompetence).Value
                        If StrComp(strLettre, vbNullString) <> 0 And IsEmpty(dblCoeffCompetence) = False Then
                            dblSomme = dblSomme + dblCoeffCompetence * lettreToValeur(strLettre)
                            dblDiviseur = dblDiviseur + dblCoeffCompetence
                        End If
                    Next intIndiceCompetence
                End If
            Next intIndiceEval
        End With
        
        ' *** AFFECTATION RESULTAT PAGE 4 ***
        With Worksheets(strPage4)
            If dblSomme <> 0 Then
                .Cells(3 + intIndiceEleve, 1 + 4 * (intIndiceDomaine - 1) + intIndiceTrimestre).Value = valeurToLettre(dblSomme / dblDiviseur)
            ElseIf dblSomme = 0 And dblDiviseur = 0 Then
                .Cells(3 + intIndiceEleve, 1 + 4 * (intIndiceDomaine - 1) + intIndiceTrimestre).Value = vbNullString
            End If
        End With
    Next intIndiceEleve
    
End Sub

' Calcul la moyenne globale trimestre
Sub calculMoyenneGlobale(intIndiceClasse As Integer, intIndiceTrimestre As Integer)
    ' *** DECLARATION VARIABLES ***
    Dim strNomClasse As String                  ' Nom de la classe
    Dim intIndiceEleve As Integer               ' Indice de l'élève actuel
    Dim intNbEleves As Integer                  ' Nombre total d'élèves dans la classe
    Dim intIndiceEval As Integer                ' Indice de l'éval actuelle
    Dim intNbEvals As Integer                   ' Nombre total d'évals
    Dim intNbDomaines As Integer                ' Nombre de domaines
    Dim intMoitieTotalCompetences As Integer    ' Nombre moitié des compétences - pour la séparation IndiceTrimestre/CoeffEval
    Dim intTotalCompetences As Integer          ' Nombre total compétences
    Dim varNote As Variant                      ' Note d'un élève à une éval
    Dim dblSomme As Double                      ' Somme pondérée des notes d'un élève
    Dim dblDiviseur As Double                   ' Somme des coeff des évals
    Dim dblCoeffEval As Double                  ' Coeff d'une éval
    Dim strPage3 As String                      ' Nom de la page 3
    Dim strPage4 As String                      ' Nom de la page 4

    ' *** AFFECTATION VARIABLES ***
    strNomClasse = getNomClasse(intIndiceClasse)
    strPage3 = "Notes (" & strNomClasse & ")"
    strPage4 = "Bilan (" & strNomClasse & ")"
    intNbEleves = getNombreEleves(intIndiceClasse)
    intNbEvals = getNombreEvals(intIndiceClasse)
    intNbDomaines = getNombreDomaines
    intTotalCompetences = getNombreCompetences
    intMoitieTotalCompetences = (intTotalCompetences - intTotalCompetences Mod 2) / 2
    
    ' *** CALCUL MOYENNE GLOBALE TRIMESTRE ***
    For intIndiceEleve = 1 To intNbEleves
        dblSomme = 0
        dblDiviseur = 0
        
        ' *** CALCUL AVEC DONNEES PAGE 3 ***
        With Worksheets(strPage3)
            For intIndiceEval = 1 To intNbEvals
                If intIndiceTrimestre = 4 Or .Cells(2, 3 + (intIndiceEval - 1) * (intTotalCompetences + 1)).Value = intIndiceTrimestre Then
                    varNote = .Cells(intLigListePage3 + intIndiceEleve, 2 + (intIndiceEval) * (intTotalCompetences + 1)).Value
                    If Not IsEmpty(varNote) Then        ' Calcul si varNote /= Null
                        dblCoeffEval = .Cells(2, 3 + intMoitieTotalCompetences + (intIndiceEval - 1) * (intTotalCompetences + 1)).Value
                        dblSomme = dblSomme + dblCoeffEval * CDbl(varNote)
                        dblDiviseur = dblDiviseur + dblCoeffEval
                    End If
                End If
            Next intIndiceEval
        End With
        
        ' *** AFFECTATION RESULTAT DANS PAGE 4 ***
        With Worksheets(strPage4)
            If dblSomme <> 0 Then
                .Cells(3 + intIndiceEleve, 1 + 4 * intNbDomaines + intIndiceTrimestre).Value = Format(dblSomme / dblDiviseur, "Standard")
            ElseIf dblSomme = 0 And dblDiviseur = 0 Then
                .Cells(3 + intIndiceEleve, 1 + 4 * intNbDomaines + intIndiceTrimestre).Value = vbNullString
            End If
        End With
    Next intIndiceEleve
End Sub


