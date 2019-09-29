Attribute VB_Name = "Module2"
' ##################################
' PAGE 2 (listes élèves)
' ##################################

Option Explicit

' **********************************
' FONCTIONS
' **********************************

' Retourne l'index de l'élève s'il est dans la liste de la classe donnée en argument, -1 sinon
' valeurExacte = True -> on cherche la place de l'élève donné en argument (supposant qu'il fait partie de la classe)
' valeurExacte = False -> on cherche où intégrer l'élève pour respecter l'ordre alphabéthique
Function getIndiceEleve(strNomComplet As String, intIndiceClasse As Integer, bValExacte As Boolean) As Integer
    Dim intNombreEleves As Integer, intIndiceEleve As Integer
    
    intNombreEleves = getNombreEleves(intIndiceClasse)
    getIndiceEleve = -1
    For intIndiceEleve = 1 To intNombreEleves
        If Not (bValExacte) Then
            If StrComp(strNomComplet, Sheets(strPage2).Cells(3 + intIndiceEleve, intIndiceClasse * 2 - 1).Value) = -1 Then
                getIndiceEleve = intIndiceEleve
                Exit For
            ElseIf intIndiceEleve = intNombreEleves Then getIndiceEleve = intNombreEleves + 1
            End If
        Else
            If StrComp(strNomComplet, Sheets(strPage2).Cells(3 + intIndiceEleve, intIndiceClasse * 2 - 1).Value) = 0 Then getIndiceEleve = intIndiceEleve
        End If
    Next intIndiceEleve
End Function

' **********************************
' PROCÉDURES
' **********************************

Sub creerListeEleve()
    Dim intNombreClasses As Integer, intNombreEleves As Integer
    Dim intColonne As Integer, intLigBouton As Integer
    Dim rngBouton As Range
    Dim btnBouton As Variant
    
    ' Données nécessaires
    intNombreClasses = getNombreClasses

    Application.ScreenUpdating = False
    
    ' Creation page
    ActiveWorkbook.Unprotect strPassword
    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Name = strPage2
    ActiveWorkbook.Protect strPassword, True, True
    
    Cells.Borders.ColorIndex = 2
    Cells.Locked = True
    
    ' Figeage volets
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 3
        .FreezePanes = True
    End With
    
    ' Creation listes vides
    For intColonne = 1 To (2 * intNombreClasses)
        If intColonne Mod 2 = 1 Then
            ' Formatage intColonne paire
            intNombreEleves = getNombreEleves((intColonne + 1) / 2)
            Columns(intColonne).ColumnWidth = 40
            For intLigBouton = 1 To 2
                Set rngBouton = Cells(intLigBouton, intColonne)
                Set btnBouton = ActiveSheet.Buttons.Add(rngBouton.Left, rngBouton.Top, rngBouton.Width, rngBouton.Height)
                If intLigBouton = 1 Then
                    With btnBouton
                        .Caption = "Ajouter élève"
                        .OnAction = "btnAjouterEleve_Click"
                    End With
                Else
                    With btnBouton
                        .Caption = "Supprimer élève"
                        .OnAction = "btnSupprimerEleve_Click"
                    End With
                End If
            Next intLigBouton
            With Cells(3, intColonne)
                .Borders.ColorIndex = 1
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlMedium
                .Interior.ColorIndex = intColorClasse
                .Value = Sheets(strPage1).Cells(12 + (intColonne + 1) / 2, 6).Value
                .HorizontalAlignment = xlHAlignCenter
                .VerticalAlignment = xlVAlignCenter
                .Locked = True
            End With
            With Range(Cells(4, intColonne), Cells(3 + intNombreEleves, intColonne))
                .Borders.ColorIndex = 1
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
                .VerticalAlignment = xlVAlignCenter
                .Locked = False
            End With
        Else
            ' Formatage intColonne impaire
            Columns(intColonne).ColumnWidth = 5
        End If
    Next intColonne
    
    ' Creation bouton "Créer Tableaux"
    Columns(intColonne).ColumnWidth = 30
    Set rngBouton = Cells(1, intColonne)
    Set btnBouton = ActiveSheet.Buttons.Add(rngBouton.Left, rngBouton.Top, rngBouton.Width, rngBouton.Height)
    With btnBouton
        .Caption = "Créer Tableaux"
        .OnAction = "btnCreerTableaux_Click"
    End With
    With Cells(2, intColonne)
        .Value = "Après avoir rempli les listes"
        .Interior.ColorIndex = 3
        .Borders.ColorIndex = xlColorIndexAutomatic
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
    End With
    
    Set rngBouton = Nothing
    Set btnBouton = Nothing
    
    ' Protection page
    With Sheets(strPage2)
        .EnableSelection = xlUnlockedCells
        .Protect strPassword
    End With
    Application.ScreenUpdating = True
    
    
End Sub

' *** Origine: bouton "Créer Tableaux"
' *** Action: crée la feuille de listes de classes et tous les tableaux 'Classes' et 'Eval'
Sub btnCreerTableaux_Click()
    Dim intNombreClasses As Integer, intIndiceClasse As Integer
    Dim intNombreEleves As Integer
    
    intNombreClasses = getNombreClasses
    
    ' Confirmation
    If MsgBox("Êtes-vous sûr(e) de valider ces listes ? Vous pourrez toujours ajouter des élèves mais il sera impossible de recréer les tableaux.", vbYesNo) = vbYes Then
        
        Application.ScreenUpdating = False
    
        ' Creation des pages 'Notes' et 'Bilan'
        For intIndiceClasse = 1 To intNombreClasses
            intNombreEleves = getNombreEleves(intIndiceClasse)
            creerTableauNotes intIndiceClasse, intNombreEleves
            creerTableauBilan intIndiceClasse, intNombreEleves
        Next intIndiceClasse
        
        ' Verouillage des listes
        With Sheets(strPage2)
            .Unprotect strPassword
            .Buttons(.Buttons.Count).Delete
            .Cells(2, 2 * intNombreClasses + 1).Delete xlShiftUp
            .Cells.Locked = True
            .EnableSelection = xlUnlockedCells
            .Protect strPassword
        End With
        
        Application.ScreenUpdating = True
        
        MsgBox ("Tableaux de notes et de bilan créés avec succès !")
    Else
        MsgBox ("Opération annulée.")
    End If
    
End Sub

' Procédure d'ajout d'un élève
Sub btnAjouterEleve_Click()
    Dim intIndiceClasse As Integer, strNomClasse As String
    Dim intIndiceEleve As Integer, strNomEleve As String, strPrenomEleve As String, strNomComplet As String
    
    
    ' Classe
    intIndiceClasse = WorksheetFunction.RoundUp(Val(Right(Application.Caller, 1)) / 2, 0)
    strNomClasse = getNomClasse(intIndiceClasse)

    ' Eleve
    strNomEleve = InputBox("Nom de l'élève à ajouter :")
    strPrenomEleve = InputBox("Prénom de l'élève à ajouter :")
    strNomComplet = StrConv(strNomEleve, vbUpperCase) & " " & StrConv(strPrenomEleve, vbProperCase)
    intIndiceEleve = getIndiceEleve(strNomComplet, intIndiceClasse, False)

    'Confirmation
    If MsgBox("Voulez vous ajouter l'élève '" & strNomComplet & "' à la classe '" & strNomClasse & "' ?", vbYesNo) = vbYes Then
        ajouterEleve intIndiceClasse, intIndiceEleve, strNomComplet
        MsgBox ("Élève ajouté !")
    Else
        MsgBox ("Opération annulée.")
    End If

End Sub

Sub ajouterEleve(intIndiceClasse As Integer, intIndiceEleve As Integer, strNomComplet As String)
    Dim strNomClasse As String
    Dim strPage3 As String, strPage4 As String
    Dim intNombreCompetences As Integer
    Dim intNombreEleves As Integer
    Dim intNombreEval
    
    ' Données initiales
    strNomClasse = getNomClasse(intIndiceClasse)
    strPage3 = "Notes (" & strNomClasse & ")"
    strPage4 = "Bilan (" & strNomClasse & ")"
    intNombreCompetences = getNombreCompetences
    intNombreEleves = getNombreEleves(strNomClasse)
    
    If Not (intIndiceEleve > intNombreMinEleves And intIndiceEleve <= intNombreEleves + 1) Then
        MsgBox ("L'indice de l'élève n'est pas compris dans la classe.")
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' Ajout page 1 (accueil)
    Sheets(strPage1).Unprotect strPassword
    Sheets(strPage1).Cells(12 + intIndiceClasse, 7).Value = intNombreEleves + 1
    Sheets(strPage1).Protect strPassword
    
    ' Ajout page 2 (liste)
    With Sheets(strPage2)
        .Unprotect strPassword
        If intIndiceEleve > 2 And intIndiceEleve < intNombreEleves + 1 Then
            .Cells(3 + intIndiceEleve, 2 * intIndiceClasse - 1).Insert xlShiftDown, xlFormatFromLeftOrAbove
        ElseIf intIndiceEleve = intNombreEleves + 1 Then
            .Cells(3 + intIndiceEleve - 1, 2 * intIndiceClasse - 1).Insert xlShiftDown, xlFormatFromLeftOrAbove
            .Cells(3 + intIndiceEleve - 1, 2 * intIndiceClasse - 1).Value = .Cells(3 + intIndiceEleve, 2 * intIndiceClasse - 1).Value
        Else
            .Cells(3 + intIndiceEleve + 1, 2 * intIndiceClasse - 1).Insert xlShiftDown, xlFormatFromRightOrBelow
            .Cells(3 + intIndiceEleve + 1, 2 * intIndiceClasse - 1).Value = .Cells(3 + intIndiceEleve, 2 * intIndiceClasse - 1).Value
        End If
        .Cells(3 + intIndiceEleve, 2 * intIndiceClasse - 1).Value = strNomComplet
        .Cells.Locked = True
        .EnableSelection = xlUnlockedCells
        .Protect strPassword
    End With
    
    If Sheets.Count > 3 Then
        intNombreEval = Sheets(strPage3).Buttons.Count - 1
        ' Ajout page 3 (notes)
        With Sheets(strPage3)
            .Unprotect strPassword
            If intIndiceEleve > 2 And intIndiceEleve < intNombreEleves + 1 Then
                .Range(.Cells(5 + intIndiceEleve, 1), .Cells(5 + intIndiceEleve, 2 + (intNombreCompetences + 1) * intNombreEval)).Insert xlDown, xlFormatFromLeftOrAbove
            ElseIf intIndiceEleve = intNombreEleves + 1 Then
                .Range(.Cells(5 + intIndiceEleve - 1, 1), .Cells(5 + intIndiceEleve - 1, 2 + (intNombreCompetences + 1) * intNombreEval)).Insert xlDown, xlFormatFromLeftOrAbove
                .Range(.Cells(5 + intIndiceEleve - 1, 1), .Cells(5 + intIndiceEleve - 1, 2 + (intNombreCompetences + 1) * intNombreEval)).Value = .Range(.Cells(5 + intIndiceEleve, 1), .Cells(5 + intIndiceEleve, 2 + (intNombreCompetences + 1) * intNombreEval)).Value
                .Range(.Cells(5 + intIndiceEleve, 1), .Cells(5 + intIndiceEleve, 2 + (intNombreCompetences + 1) * intNombreEval)).ClearContents
                .Range(.Cells(5 + intIndiceEleve - 1, 1), .Cells(5 + intIndiceEleve - 1, 2)).MergeCells = True
            Else
                .Range(.Cells(5 + intIndiceEleve + 1, 1), .Cells(5 + intIndiceEleve + 1, 2 + (intNombreCompetences + 1) * intNombreEval)).Insert xlDown, xlFormatFromRightOrBelow
                .Range(.Cells(5 + intIndiceEleve + 1, 1), .Cells(5 + intIndiceEleve + 1, 2 + (intNombreCompetences + 1) * intNombreEval)).Value = .Range(.Cells(5 + intIndiceEleve, 1), .Cells(5 + intIndiceEleve, 2 + (intNombreCompetences + 1) * intNombreEval)).Value
                .Range(.Cells(5 + intIndiceEleve, 1), .Cells(5 + intIndiceEleve, 2 + (intNombreCompetences + 1) * intNombreEval)).ClearContents
                .Range(.Cells(5 + intIndiceEleve + 1, 1), .Cells(5 + intIndiceEleve + 1, 2)).MergeCells = True
            End If
            .Range(.Cells(5 + intIndiceEleve, 1), .Cells(5 + intIndiceEleve, 2)).MergeCells = True
            .Cells(5 + intIndiceEleve, 1).Value = strNomComplet
            .EnableSelection = xlUnlockedCells
            .Protect strPassword
        End With
        
        ' Ajout page 4 (résultats)
        With Sheets(strPage4)
            .Unprotect strPassword
            If intIndiceEleve > 2 And intIndiceEleve < intNombreEleves + 1 Then
                .Range(.Cells(3 + intIndiceEleve, 1), .Cells(3 + intIndiceEleve, 1 + 4 * (intNombreCompetences + 1))).Insert xlDown, xlFormatFromLeftOrAbove
            ElseIf intIndiceEleve = intNombreEleves + 1 Then
                .Range(.Cells(3 + intIndiceEleve - 1, 1), .Cells(3 + intIndiceEleve - 1, 1 + 4 * (intNombreCompetences + 1))).Insert xlDown, xlFormatFromLeftOrAbove
                .Range(.Cells(3 + intIndiceEleve - 1, 1), .Cells(3 + intIndiceEleve - 1, 1 + 4 * (intNombreCompetences + 1))).Value = .Range(.Cells(3 + intIndiceEleve, 1), .Cells(3 + intIndiceEleve, 1 + 4 * (intNombreCompetences + 1))).Value
                .Range(.Cells(3 + intIndiceEleve, 1), .Cells(3 + intIndiceEleve, 1 + 4 * (intNombreCompetences + 1))).ClearContents
            Else
                .Range(.Cells(3 + intIndiceEleve + 1, 1), .Cells(3 + intIndiceEleve + 1, 1 + 4 * (intNombreCompetences + 1))).Insert xlDown, xlFormatFromRightOrBelow
                .Range(.Cells(3 + intIndiceEleve + 1, 1), .Cells(3 + intIndiceEleve + 1, 1 + 4 * (intNombreCompetences + 1))).Value = .Range(.Cells(3 + intIndiceEleve, 1), .Cells(3 + intIndiceEleve, 1 + 4 * (intNombreCompetences + 1))).Value
                .Range(.Cells(3 + intIndiceEleve, 1), .Cells(3 + intIndiceEleve, 1 + 4 * (intNombreCompetences + 1))).ClearContents
            End If
            .Cells(3 + intIndiceEleve, 1).Value = strNomComplet
            .EnableSelection = xlUnlockedCells
            .Protect strPassword
        End With
    End If
    
    Application.ScreenUpdating = True
    
End Sub

Sub btnSupprimerEleve_Click()
    Dim intIndiceClasse As Integer, strNomClasse As String
    Dim strNomEleve As String, strPrenomEleve As String, strNomComplet As String
    
    ' Classe
    intIndiceClasse = WorksheetFunction.RoundUp(Val(Right(Application.Caller, 1)) / 2, 0)
    strNomClasse = getNomClasse(intIndiceClasse)

    ' Eleve
    strNomEleve = InputBox("Nom de l'élève à supprimer (comme spécifié dans la liste, écrire en minuscule si accent) :")
    strPrenomEleve = InputBox("Prénom de l'élève à supprimer (comme spécifié dans la liste, écrire en minuscule si accent) :")
    strNomComplet = StrConv(strNomEleve, vbUpperCase) & " " & StrConv(strPrenomEleve, vbProperCase)

    'Confirmation
    If MsgBox("Voulez vous supprimer l'élève '" & strNomComplet & "' de la classe '" & strNomClasse & "' ?", vbYesNo) = vbYes Then
        If getIndiceEleve(strNomComplet, intIndiceClasse, True) <> -1 Then
            supprimerEleve intIndiceClasse, strNomComplet
            MsgBox ("Élève supprimé avec succès.")
        Else
            MsgBox ("L'élève '" & strNomComplet & "' n'a pas été trouvé dans la classe " & strNomClasse & ". Vérifiez l'orthographe.")
        End If
    Else
        MsgBox ("Opération annulée.")
    End If

End Sub

Sub supprimerEleve(intIndiceClasse As Integer, strNomComplet As String)
    Dim intIndiceEleve As Integer
    intIndiceEleve = getIndiceEleve(strNomComplet, intIndiceClasse, True)
    If intIndiceEleve = -1 Then Exit Sub
    
    Dim strNomClasse As String
    Dim strPage3 As String, strPage4 As String
    Dim intNombreDomaines As Integer, intNombreCompetences As Integer
    Dim intNombreEleves As Integer
    Dim intNombreEval As Integer
    
    ' Données initiales
    strNomClasse = getNomClasse(intIndiceClasse)
    strPage3 = "Notes (" & strNomClasse & ")"
    strPage4 = "Bilan (" & strNomClasse & ")"
    intNombreDomaines = getNombreDomaines
    intNombreCompetences = getNombreCompetences
    intNombreEleves = getNombreEleves(strNomClasse)
    
    Application.ScreenUpdating = False
    
    ' Modification Page 1
    Sheets(strPage1).Unprotect strPassword
    Sheets(strPage1).Cells(12 + intIndiceClasse, 7).Value = intNombreEleves - 1
    Sheets(strPage1).Protect strPassword
    
    ' Modification Page 2
    With Sheets(strPage2)
        .Unprotect strPassword
        If intIndiceEleve > 1 And intIndiceEleve < intNombreEleves Then
            .Cells(3 + intIndiceEleve, 2 * intIndiceClasse - 1).Delete xlShiftUp
        ElseIf intIndiceEleve = intNombreEleves Then
            .Cells(3 + intIndiceEleve, 2 * intIndiceClasse - 1).Value = .Cells(3 + intIndiceEleve - 1, 2 * intIndiceClasse - 1).Value
            .Cells(3 + intIndiceEleve - 1, 2 * intIndiceClasse - 1).Delete xlShiftUp
        Else
            .Cells(3 + intIndiceEleve, 2 * intIndiceClasse - 1).Value = .Cells(3 + intIndiceEleve + 1, 2 * intIndiceClasse - 1).Value
            .Cells(3 + intIndiceEleve + 1, 2 * intIndiceClasse - 1).Delete xlShiftUp
        End If
        .Cells.Locked = True
        .EnableSelection = xlUnlockedCells
        .Protect strPassword
    End With
    
    If Sheets.Count > 3 Then
        intNombreEval = Sheets(strPage3).Buttons.Count - 1
        ' Modification Page 3 (notes)
        With Sheets(strPage3)
            .Unprotect strPassword
            If intIndiceEleve > 1 And intIndiceEleve < intNombreEleves Then
                .Range(.Cells(5 + intIndiceEleve, 1), .Cells(5 + intIndiceEleve, 2 + (intNombreCompetences + 1) * intNombreEval)).Delete xlShiftUp
            ElseIf intIndiceEleve = intNombreEleves Then
                .Range(.Cells(5 + intIndiceEleve, 1), .Cells(5 + intIndiceEleve, 2 + (intNombreCompetences + 1) * intNombreEval)).Value = .Range(.Cells(5 + intIndiceEleve - 1, 1), .Cells(5 + intIndiceEleve - 1, 2 + (intNombreCompetences + 1) * intNombreEval)).Value
                .Range(.Cells(5 + intIndiceEleve - 1, 1), .Cells(5 + intIndiceEleve - 1, 2 + (intNombreCompetences + 1) * intNombreEval)).Delete xlShiftUp
            Else
                .Range(.Cells(5 + intIndiceEleve, 1), .Cells(5 + intIndiceEleve, 2 + (intNombreCompetences + 1) * intNombreEval)).Value = .Range(.Cells(5 + intIndiceEleve + 1, 1), .Cells(5 + intIndiceEleve + 1, 2 + (intNombreCompetences + 1) * intNombreEval)).Value
                .Range(.Cells(5 + intIndiceEleve + 1, 1), .Cells(5 + intIndiceEleve + 1, 2 + (intNombreCompetences + 1) * intNombreEval)).Delete xlShiftUp
            End If
            .EnableSelection = xlUnlockedCells
            .Protect strPassword
        End With
        
        ' Modification Page 4 (résultats)
        With Sheets(strPage4)
            .Unprotect strPassword
            If intIndiceEleve > 1 And intIndiceEleve < intNombreEleves Then
                .Range(.Cells(3 + intIndiceEleve, 1), .Cells(3 + intIndiceEleve, 1 + 4 * (intNombreCompetences + 1))).Delete xlShiftUp
            ElseIf intIndiceEleve = intNombreEleves Then
                .Range(.Cells(3 + intIndiceEleve, 1), .Cells(3 + intIndiceEleve, 1 + 4 * (intNombreCompetences + 1))).Value = .Range(.Cells(3 + intIndiceEleve - 1, 1), .Cells(3 + intIndiceEleve - 1, 1 + 4 * (intNombreCompetences + 1))).Value
                .Range(.Cells(3 + intIndiceEleve - 1, 1), .Cells(3 + intIndiceEleve - 1, 1 + 4 * (intNombreCompetences + 1))).Delete xlShiftUp
            Else
                .Range(.Cells(3 + intIndiceEleve, 1), .Cells(3 + intIndiceEleve, 1 + 4 * (intNombreCompetences + 1))).Value = .Range(.Cells(3 + intIndiceEleve + 1, 1), .Cells(3 + intIndiceEleve + 1, 1 + 4 * (intNombreCompetences + 1))).Value
                .Range(.Cells(3 + intIndiceEleve + 1, 1), .Cells(3 + intIndiceEleve + 1, 1 + 4 * (intNombreCompetences + 1))).Delete xlShiftUp
            End If
            .EnableSelection = xlUnlockedCells
            .Protect strPassword
        End With
    End If
    
    Application.ScreenUpdating = True
End Sub

