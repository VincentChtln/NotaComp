Attribute VB_Name = "Module2"
' ##################################
' PAGE 2 (listes élèves)
' ##################################

Option Explicit

' ##################################
' FONCTIONS
' ##################################
' getIndiceEleve(strNomComplet As String, intIndiceClasse As Integer, bValeurExacte As Boolean) As Integer
' ##################################

' Retourne l'index de l'élève s'il est dans la liste de la classe donnée en argument, -1 sinon
' valeurExacte = True -> on cherche la place de l'élève donné en argument (supposant qu'il fait partie de la classe)
' valeurExacte = False -> on cherche où intégrer l'élève pour respecter l'ordre alphabéthique
Function getIndiceEleve(strNomComplet As String, intIndiceClasse As Integer, bValeurExacte As Boolean) As Integer
    ' *** DECLARATION VARIABLES ***
    Dim intNbEleves As Integer
    Dim intIndiceEleve As Integer
    
    ' *** AFFECTATION VARIABLES ***
    intNbEleves = getNombreEleves(intIndiceClasse)
    getIndiceEleve = -1
    
    ' *** RECHERCHE INDICE EXACT ***
    If bValeurExacte Then
        For intIndiceEleve = 1 To intNbEleves
            If StrComp(strNomComplet, Worksheets(strPage2).Cells(intLigListePage2 + intIndiceEleve, intIndiceClasse * 2 - 1).Value) = 0 Then
                getIndiceEleve = intIndiceEleve
                Exit For
            End If
        Next intIndiceEleve
        
    ' *** RECHERCHE INDICE POUR INSERTION ***
    Else
        For intIndiceEleve = 1 To intNbEleves
            If StrComp(strNomComplet, Worksheets(strPage2).Cells(intLigListePage2 + intIndiceEleve, intIndiceClasse * 2 - 1).Value) = -1 Then
                getIndiceEleve = intIndiceEleve
                Exit For
            ElseIf intIndiceEleve = intNbEleves Then getIndiceEleve = intNbEleves + 1
            End If
        Next intIndiceEleve
    End If
End Function

' ##################################
' PROCÉDURES
' ##################################
' creerListeEleve()
' btnCreerTableaux_Click()
' btnModifierListe_Click()
'
' btnAjouterEleve_Click()
' ajouterEleve(intIndiceClasse As Integer, intIndiceEleve As Integer, strNomComplet As String)
'
' btnSupprimerEleve_Click()
' supprimerEleve(intIndiceClasse As Integer, intIndiceEleve As Integer)
'
' transfererEleve(intClasseSource As Integer, intIndiceSourceEleve As Integer, intClasseDest As Integer, intIndiceDestEleve As Integer, strNomComplet As String)
' copierNotesEleve(intIndiceClasseSource As Integer, intIndiceEleveSource As Integer, intIndiceClasseDest As Integer, intIndiceEleveDest As Integer)
' ##################################

Sub creerListeEleve()
    ' *** DECLARATION VARIABLES ***
    Dim intNbClasses As Integer
    Dim intIndiceClasse As Integer
    Dim intNbEleves As Integer
    Dim intMaxEleves As Integer
    Dim intColonne As Integer
    Dim rngBouton As Range
    Dim btnBouton As Variant
    
    ' *** PROTECTION + REFRESH ECRAN OFF ***
    Application.ScreenUpdating = False
    unprotectWorkbook
    unprotectWorksheet
    
    ' *** AJOUT PAGE 2 - LISTE ELEVE ***
    Worksheets.Add After:=Worksheets(Worksheets.Count)
    Worksheets(Worksheets.Count).Name = strPage2
    
    With Worksheets(strPage2)
        ' *** FORMATAGE PAGE ***
        freezePanes ActiveWindow, intLigListePage2, 0
        .Cells.Borders.ColorIndex = 2
        .Cells.Locked = True
        
        ' *** AFFECTATION VARIABLES ***
        intNbClasses = getNombreClasses
    
        ' *** CREATION LISTE VIDE ***
        For intColonne = 1 To (2 * intNbClasses)
            If intColonne Mod 2 = 1 Then
                ' *** AFFECTATION VARIABLES ***
                intIndiceClasse = (intColonne + 1) / 2
                
                ' *** FORMATAGE COLONNE IMPAIRE ***
                intNbEleves = getNombreEleves(intIndiceClasse)
                If intMaxEleves < intNbEleves Then intMaxEleves = intNbEleves
                .Columns(intColonne).ColumnWidth = 40
                With .Cells(1, intColonne)
                    .Borders.ColorIndex = 1
                    .Borders.LineStyle = xlContinuous
                    .Borders.Weight = xlMedium
                    .Interior.ColorIndex = intColorClasse
                    .Value = Worksheets(strPage1).Cells(12 + intIndiceClasse, 6).Value
                    .HorizontalAlignment = xlHAlignCenter
                    .VerticalAlignment = xlVAlignCenter
                    .Locked = True
                End With
                With .Range(.Cells(intLigListePage2 + 1, intColonne), .Cells(intLigListePage2 + intNbEleves, intColonne))
                    .Borders.ColorIndex = 1
                    .Borders.LineStyle = xlContinuous
                    .Borders.Weight = xlThin
                    .VerticalAlignment = xlVAlignCenter
                    .Locked = False
                End With
            Else
                ' *** FORMATAGE COLONNE PAIRE ***
                .Columns(intColonne).ColumnWidth = 5
            End If
        Next intColonne
        
        .Columns(intColonne).ColumnWidth = 30
        
        ' *** CREATION BOUTON 'CREER TABLEAUX' ***
        Set rngBouton = .Cells(intLigListePage2, intColonne)
        Set btnBouton = .Buttons.Add(rngBouton.Left, rngBouton.Top, rngBouton.Width, rngBouton.Height)
        With btnBouton
            .Caption = "Modifier listes"
            .OnAction = "btnModifierListe_Click"
        End With
        
        ' *** CREATION BOUTON 'CREER TABLEAUX' ***
        Set rngBouton = .Cells(intLigListePage2 + 2, intColonne)
        Set btnBouton = .Buttons.Add(rngBouton.Left, rngBouton.Top, rngBouton.Width, rngBouton.Height)
        With btnBouton
            .Caption = "Créer Tableaux"
            .OnAction = "btnCreerTableaux_Click"
        End With
        With .Cells(intLigListePage2 + 3, intColonne)
            .Value = "Après avoir rempli les listes"
            .Interior.ColorIndex = 3
            .Borders.ColorIndex = xlColorIndexAutomatic
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
        End With
    End With
    
    ' *** LIMITATION ZONE SCROLL ***
'    With Worksheets(strPage2)
'        .ScrollArea = .Range(.Cells("A1"), .Cells(intLigListePage2 + intMaxEleves + 5, 2 * intNbClasses + 2))
'    End With

    ' *** PROTECTION + REFRESH ECRAN ON ***
    protectWorksheet
    protectWorkbook
    Application.ScreenUpdating = True
    
End Sub

Sub btnModifierListe_Click()
    UserForm1.Show
End Sub

' *** Origine: bouton "Créer Tableaux"
' *** Action: crée la feuille de listes de classes et tous les tableaux 'Classes' et 'Eval'
Sub btnCreerTableaux_Click()
    ' *** DEMANDE CONFIRMATION ***
    If MsgBox("Êtes-vous sûr(e) de valider ces listes ? Vous pourrez toujours les modifier (ajout, suppression ou transfert d'élève) mais il sera impossible de les recréer intégralement.", vbYesNo) = vbNo Then
        MsgBox ("Opération annulée.")
        Exit Sub
    End If
    
    ' *** DECLARATION VARIABLES ***
    Dim intNbClasses As Integer
    Dim intIndiceClasse As Integer
    Dim intNbEleves As Integer
    Dim intAvancementActuel As Integer
    Dim intAvancementTotal As Integer
    
    ' *** PROTECTION + REFRESH ECRAN OFF ***
    UserForm5.Show vbModeless
    Application.ScreenUpdating = False
    unprotectWorksheet
    unprotectWorkbook
    
    ' *** AFFECTATION VARIABLES ***
    intNbClasses = getNombreClasses
    intAvancementActuel = 0
    intAvancementTotal = 2 * intNbClasses
    
    ' *** AJOUT PAGE 3 + PAGE 4 ***
    For intIndiceClasse = 1 To intNbClasses
        intNbEleves = getNombreEleves(intIndiceClasse)
        
        ' *** AJOUT PAGE 3 + MAJ CHARGEMENT ***
        creerTableauNotes intIndiceClasse, intNbEleves
        intAvancementActuel = intAvancementActuel + 1
        UserForm5.updateAvancement intAvancementActuel, intAvancementTotal
        
        ' *** AJOUT PAGE 4 + MAJ CHARGEMENT ***
        creerTableauBilan intIndiceClasse, intNbEleves
        intAvancementActuel = intAvancementActuel + 1
        UserForm5.updateAvancement intAvancementActuel, intAvancementTotal
    Next intIndiceClasse
    
    ' *** FORMATAGE PAGE 2 ***
    With Worksheets(strPage2)
        .Buttons(.Buttons.Count).Delete
        .Cells(intLigListePage2 + 3, 2 * intNbClasses + 1).Delete xlShiftUp
        .Cells.Locked = True
    End With
    
    ' *** PROTECTION + REFRESH ECRAN ON ***
    UserForm5.Hide
    protectWorksheet
    protectWorkbook
    Application.ScreenUpdating = True
    
    ' *** MESSAGE INFORMATION ***
    MsgBox ("Tableaux 'Notes' et 'Bilan' créés avec succès !")
End Sub

Sub ajouterEleve(intIndiceClasse As Integer, intIndiceEleve As Integer, strNomComplet As String)
    ' *** DECLARATION VARIABLES ***
    Dim strNomClasse As String              ' Nom de la classe d'ajout
    Dim strPage3 As String                  ' Nom de la page "Notes" liée à la classe
    Dim strPage4 As String                  ' Nom de la page "Bilan" liée à la classe
    Dim intNbCompetences As Integer         ' Nombre de compétences pour le Workbook
    Dim intNbEleves As Integer              ' Nombre d'élèves de la classe
    Dim intNbEval As Integer                ' Nombre d'éval de la classe

    ' *** PROTECTION + REFRESH ECRAN OFF ***
    Application.ScreenUpdating = False
    unprotectWorksheet intIndiceClasse

    ' *** MODIFICATION PAGE 1 - ACCUEIL ***
    intNbEleves = getNombreEleves(intIndiceClasse)
    setNombreEleves intIndiceClasse, intNbEleves + 1

    ' *** MODIFICATION PAGE 2 - LISTE ***
    With Worksheets(strPage2)
        Select Case intIndiceEleve
        Case Is = 1                     ' Cas 1: élève en début de liste
            .Cells(intLigListePage2 + intIndiceEleve + 1, 2 * intIndiceClasse - 1).Insert xlShiftDown, xlFormatFromRightOrBelow
            .Cells(intLigListePage2 + intIndiceEleve + 1, 2 * intIndiceClasse - 1).Value = .Cells(intLigListePage2 + intIndiceEleve, 2 * intIndiceClasse - 1).Value
        Case 2 To intNbEleves           ' Cas 2: élève au milieu de la liste
            .Cells(intLigListePage2 + intIndiceEleve, 2 * intIndiceClasse - 1).Insert xlShiftDown, xlFormatFromLeftOrAbove
        Case Is = intNbEleves + 1       ' Cas 3: élève en fin de liste
            .Cells(intLigListePage2 + intIndiceEleve - 1, 2 * intIndiceClasse - 1).Insert xlShiftDown, xlFormatFromLeftOrAbove
            .Cells(intLigListePage2 + intIndiceEleve - 1, 2 * intIndiceClasse - 1).Value = .Cells(intLigListePage2 + intIndiceEleve, 2 * intIndiceClasse - 1).Value
        End Select
        .Cells(intLigListePage2 + intIndiceEleve, 2 * intIndiceClasse - 1).Value = strNomComplet
        .Cells.Locked = True
    End With

    If Worksheets.Count > 3 Then
        ' *** AFFECTATION VARIABLES ***
        strNomClasse = getNomClasse(intIndiceClasse)
        strPage3 = "Notes (" & strNomClasse & ")"
        strPage4 = "Bilan (" & strNomClasse & ")"
        intNbCompetences = getNombreCompetences
        intNbEval = Worksheets(strPage3).Buttons.Count - 1
        
        ' *** MODIFICATION PAGE 3 - NOTES ***
        With Worksheets(strPage3)
            Select Case intIndiceEleve
            Case Is = 1                     ' Cas 1: élève en début de liste
                .Range(.Cells(intLigListePage3 + intIndiceEleve + 1, 1), .Cells(intLigListePage3 + intIndiceEleve + 1, 2 + (intNbCompetences + 1) * intNbEval)).Insert xlDown, xlFormatFromRightOrBelow
                .Range(.Cells(intLigListePage3 + intIndiceEleve + 1, 1), .Cells(intLigListePage3 + intIndiceEleve + 1, 2)).MergeCells = True
                .Range(.Cells(intLigListePage3 + intIndiceEleve + 1, 1), .Cells(intLigListePage3 + intIndiceEleve + 1, 2 + (intNbCompetences + 1) * intNbEval)).Value = .Range(.Cells(intLigListePage3 + intIndiceEleve, 1), .Cells(intLigListePage3 + intIndiceEleve, 2 + (intNbCompetences + 1) * intNbEval)).Value
                .Range(.Cells(intLigListePage3 + intIndiceEleve, 1), .Cells(intLigListePage3 + intIndiceEleve, 2 + (intNbCompetences + 1) * intNbEval)).ClearContents
            Case 2 To intNbEleves           ' Cas 2: élève au milieu de la liste
                .Range(.Cells(intLigListePage3 + intIndiceEleve, 1), .Cells(intLigListePage3 + intIndiceEleve, 2 + (intNbCompetences + 1) * intNbEval)).Insert xlDown, xlFormatFromLeftOrAbove
                .Range(.Cells(intLigListePage3 + intIndiceEleve, 1), .Cells(intLigListePage3 + intIndiceEleve, 2)).MergeCells = True
                .Cells(intLigListePage3 + intIndiceEleve, 1).Value = strNomComplet
            Case Is = intNbEleves + 1       ' Cas 3: élève en fin de liste
                .Range(.Cells(intLigListePage3 + intIndiceEleve - 1, 1), .Cells(intLigListePage3 + intIndiceEleve - 1, 2 + (intNbCompetences + 1) * intNbEval)).Insert xlDown, xlFormatFromLeftOrAbove
                .Range(.Cells(intLigListePage3 + intIndiceEleve - 1, 1), .Cells(intLigListePage3 + intIndiceEleve - 1, 2)).MergeCells = True
                .Range(.Cells(intLigListePage3 + intIndiceEleve - 1, 1), .Cells(intLigListePage3 + intIndiceEleve - 1, 2 + (intNbCompetences + 1) * intNbEval)).Value = .Range(.Cells(intLigListePage3 + intIndiceEleve, 1), .Cells(intLigListePage3 + intIndiceEleve, 2 + (intNbCompetences + 1) * intNbEval)).Value
                .Range(.Cells(intLigListePage3 + intIndiceEleve, 1), .Cells(intLigListePage3 + intIndiceEleve, 2 + (intNbCompetences + 1) * intNbEval)).ClearContents
            End Select
            .Cells(intLigListePage3 + intIndiceEleve, 1).Value = strNomComplet
        End With

        ' *** MODIFICATION PAGE 4 - BILAN ***
        With Worksheets(strPage4)
            Select Case intIndiceEleve
            Case Is = 1                     ' Cas 1: élève en début de liste
                .Range(.Cells(intLigListePage4 + intIndiceEleve + 1, 1), .Cells(intLigListePage4 + intIndiceEleve + 1, 1 + 4 * (intNbCompetences + 1))).Insert xlDown, xlFormatFromRightOrBelow
                .Range(.Cells(intLigListePage4 + intIndiceEleve + 1, 1), .Cells(intLigListePage4 + intIndiceEleve + 1, 1 + 4 * (intNbCompetences + 1))).Value = .Range(.Cells(intLigListePage4 + intIndiceEleve, 1), .Cells(intLigListePage4 + intIndiceEleve, 1 + 4 * (intNbCompetences + 1))).Value
                .Range(.Cells(intLigListePage4 + intIndiceEleve, 1), .Cells(intLigListePage4 + intIndiceEleve, 1 + 4 * (intNbCompetences + 1))).ClearContents
            Case 2 To intNbEleves           ' Cas 2: élève au milieu de la liste
                .Range(.Cells(intLigListePage4 + intIndiceEleve, 1), .Cells(intLigListePage4 + intIndiceEleve, 1 + 4 * (intNbCompetences + 1))).Insert xlDown, xlFormatFromLeftOrAbove
            Case Is = intNbEleves + 1       ' Cas 3: élève en fin de liste
                .Range(.Cells(intLigListePage4 + intIndiceEleve - 1, 1), .Cells(intLigListePage4 + intIndiceEleve - 1, 1 + 4 * (intNbCompetences + 1))).Insert xlDown, xlFormatFromLeftOrAbove
                .Range(.Cells(intLigListePage4 + intIndiceEleve - 1, 1), .Cells(intLigListePage4 + intIndiceEleve - 1, 1 + 4 * (intNbCompetences + 1))).Value = .Range(.Cells(intLigListePage4 + intIndiceEleve, 1), .Cells(intLigListePage4 + intIndiceEleve, 1 + 4 * (intNbCompetences + 1))).Value
                .Range(.Cells(intLigListePage4 + intIndiceEleve, 1), .Cells(intLigListePage4 + intIndiceEleve, 1 + 4 * (intNbCompetences + 1))).ClearContents
            End Select
            .Cells(intLigListePage4 + intIndiceEleve, 1).Value = strNomComplet
        End With
    End If

    ' *** PROTECTION + REFRESH ECRAN ON ***
    protectWorksheet intIndiceClasse
    Application.ScreenUpdating = True
    
    ' *** MESSAGE INFORMATION ***
    MsgBox "Élève ajouté.", vbInformation, "Ajout d'élève"
End Sub

Sub supprimerEleve(intIndiceClasse As Integer, intIndiceEleve As Integer)
    ' *** DECLARATION VARIABLES ***
    Dim strNomClasse As String              ' Nom de la classe d'ajout
    Dim strPage3 As String                  ' Nom de la page "Notes" liée à la classe
    Dim strPage4 As String                  ' Nom de la page "Bilan" liée à la classe
    Dim intNbCompetences As Integer         ' Nombre de compétences pour le Workbook
    Dim intNbEleves As Integer              ' Nombre d'élèves de la classe
    Dim intNbEval As Integer                ' Nombre d'éval de la classe

    ' *** PROTECTION + REFRESH ECRAN OFF ***
    Application.ScreenUpdating = False
    unprotectWorksheet intIndiceClasse

    ' *** MODIFICATION PAGE 1 - ACCUEIL ***
    intNbEleves = getNombreEleves(intIndiceClasse)
    setNombreEleves intIndiceClasse, intNbEleves - 1

    ' *** MODIFICATION PAGE 2 - LISTE ***
    With Worksheets(strPage2)
        Select Case intIndiceEleve
        Case Is = 1                         ' Cas 1: élève en début de liste
            .Cells(intLigListePage2 + intIndiceEleve, 2 * intIndiceClasse - 1).Value = .Cells(intLigListePage2 + intIndiceEleve + 1, 2 * intIndiceClasse - 1).Value
            .Cells(intLigListePage2 + intIndiceEleve + 1, 2 * intIndiceClasse - 1).Delete xlShiftUp
        Case 2 To intNbEleves - 1           ' Cas 2: élève au milieu de la liste (classique)
            .Cells(intLigListePage2 + intIndiceEleve, 2 * intIndiceClasse - 1).Delete xlShiftUp
        Case Is = intNbEleves               ' Cas 3: élève en fin de liste
            .Cells(intLigListePage2 + intIndiceEleve, 2 * intIndiceClasse - 1).Value = .Cells(intLigListePage2 + intIndiceEleve - 1, 2 * intIndiceClasse - 1).Value
            .Cells(intLigListePage2 + intIndiceEleve - 1, 2 * intIndiceClasse - 1).Delete xlShiftUp
        End Select
        .Cells.Locked = True
    End With

    If Worksheets.Count > 3 Then
        ' *** AFFECTATION VARIABLES ***
        strNomClasse = getNomClasse(intIndiceClasse)
        strPage3 = "Notes (" & strNomClasse & ")"
        strPage4 = "Bilan (" & strNomClasse & ")"
        intNbCompetences = getNombreCompetences
        intNbEval = Worksheets(strPage3).Buttons.Count - 1

        ' *** MODIFICATION PAGE 3 - NOTES ***
        With Worksheets(strPage3)
            Select Case intIndiceEleve
            Case Is = 1                         ' Cas 1: élève en début de liste
                .Range(.Cells(intLigListePage3 + intIndiceEleve, 1), .Cells(intLigListePage3 + intIndiceEleve, 2 + (intNbCompetences + 1) * intNbEval)).Value = .Range(.Cells(intLigListePage3 + intIndiceEleve + 1, 1), .Cells(intLigListePage3 + intIndiceEleve + 1, 2 + (intNbCompetences + 1) * intNbEval)).Value
                .Range(.Cells(intLigListePage3 + intIndiceEleve + 1, 1), .Cells(intLigListePage3 + intIndiceEleve + 1, 2 + (intNbCompetences + 1) * intNbEval)).Delete xlShiftUp
            Case 2 To intNbEleves - 1           ' Cas 2: élève au milieu de la liste (classique)
                .Range(.Cells(intLigListePage3 + intIndiceEleve, 1), .Cells(intLigListePage3 + intIndiceEleve, 2 + (intNbCompetences + 1) * intNbEval)).Delete xlShiftUp
            Case Is = intNbEleves               ' Cas 3: élève en fin de liste
                .Range(.Cells(intLigListePage3 + intIndiceEleve, 1), .Cells(intLigListePage3 + intIndiceEleve, 2 + (intNbCompetences + 1) * intNbEval)).Value = .Range(.Cells(intLigListePage3 + intIndiceEleve - 1, 1), .Cells(intLigListePage3 + intIndiceEleve - 1, 2 + (intNbCompetences + 1) * intNbEval)).Value
                .Range(.Cells(intLigListePage3 + intIndiceEleve - 1, 1), .Cells(intLigListePage3 + intIndiceEleve - 1, 2 + (intNbCompetences + 1) * intNbEval)).Delete xlShiftUp
            End Select
        End With

        ' *** MODIFICATION PAGE 4 - BILAN ***
        With Worksheets(strPage4)
            Select Case intIndiceEleve
            Case Is = 1                         ' Cas 1: élève en début de liste
                .Range(.Cells(intLigListePage4 + intIndiceEleve, 1), .Cells(intLigListePage4 + intIndiceEleve, 1 + 4 * (intNbCompetences + 1))).Value = .Range(.Cells(intLigListePage4 + intIndiceEleve + 1, 1), .Cells(intLigListePage4 + intIndiceEleve + 1, 1 + 4 * (intNbCompetences + 1))).Value
                .Range(.Cells(intLigListePage4 + intIndiceEleve + 1, 1), .Cells(intLigListePage4 + intIndiceEleve + 1, 1 + 4 * (intNbCompetences + 1))).Delete xlShiftUp
            Case 2 To intNbEleves - 1           ' Cas 2: élève au milieu de la liste (classique)
                .Range(.Cells(intLigListePage4 + intIndiceEleve, 1), .Cells(intLigListePage4 + intIndiceEleve, 1 + (intNbCompetences + 1) * 4)).Delete xlShiftUp
            Case Is = intNbEleves               ' Cas 3: élève en fin de liste
                .Range(.Cells(intLigListePage4 + intIndiceEleve, 1), .Cells(intLigListePage4 + intIndiceEleve, 1 + 4 * (intNbCompetences + 1))).Value = .Range(.Cells(intLigListePage4 + intIndiceEleve - 1, 1), .Cells(intLigListePage4 + intIndiceEleve - 1, 1 + 4 * (intNbCompetences + 1))).Value
                .Range(.Cells(intLigListePage4 + intIndiceEleve - 1, 1), .Cells(intLigListePage4 + intIndiceEleve - 1, 1 + 4 * (intNbCompetences + 1))).Delete xlShiftUp
            End Select
        End With
    End If

    ' *** PROTECTION + REFRESH ECRAN ON ***
    protectWorksheet intIndiceClasse
    Application.ScreenUpdating = True
    
    ' *** MESSAGE INFORMATION ***
    MsgBox "Élève supprimé.", vbInformation, "Suppression d'élève"
End Sub

' Procédure en 2 opérations:
' 1: Ajout de l'élève dans la classe de destination
' 2: Suppression de l'élève dans la classe source

Sub transfererEleve(intIndiceClasseSource As Integer, intIndiceEleveSource As Integer, intIndiceClasseDest As Integer, intIndiceEleveDest As Integer, strNomComplet As String)
     ' *** OPERATION 1 : AJOUT DANS NOUVELLE CLASSE ***
     ajouterEleve intIndiceClasseDest, intIndiceEleveDest, strNomComplet
     
     ' Il n'y a pas de transfert de notes car rien ne garanti l'homogénéïté des évaluations entre plusieurs classes.
    
     ' *** OPERATION 2 : SUPPRESSION DANS ANCIENNE CLASSE ***
     supprimerEleve intIndiceClasseSource, intIndiceEleveSource
End Sub

'Sub copierNotesEleve(intIndiceClasseSource As Integer, intIndiceEleveSource As Integer, intIndiceClasseDest As Integer, intIndiceEleveDest As Integer)
'    ' *** DECLARATION VARIABLES ***
'    Dim ws3Source As Worksheet
'    Dim ws3Dest As Worksheet
'    Dim ws4Source As Worksheet
'    Dim ws4Dest As Worksheet
'    Dim intNbEvalSource As Integer
'    Dim intNbEvalDest As Integer
'    Dim intNbDomaines As Integer
'    Dim intSommeCompetences As Integer
'
'    ' *** PROTECTION + REFRESH ECRAN OFF ***
'    Application.ScreenUpdating = False
'    unprotectWorksheet
'
'    ' *** AFFECTATION VARIABLES ***
'    Set ws3Source = Worksheets(getNomPage3(intIndiceClasseSource))
'    Set ws3Dest = Worksheets(getNomPage3(intIndiceClasseDest))
'    Set ws4Source = Worksheets(getNomPage4(intIndiceClasseSource))
'    Set ws4Dest = Worksheets(getNomPage4(intIndiceClasseDest))
'    intNbEvalSource = getNombreEvals(intIndiceClasseSource)
'
'    ' *** COPIE DES NOTES ***
'    Worksheets (ws3Dest.Name)
'
'    MsgBox "Copie des notes de l'élève"
'
'    ' *** PROTECTION + REFRESH ECRAN ON ***
'    protectWorksheet
'    Application.ScreenUpdating = True
'End Sub
'
