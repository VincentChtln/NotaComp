Attribute VB_Name = "Module2"
' **********************************
' Page 2 (listes élèves) - Procédure & fonctions
' **********************************

Sub creerListeEleve()

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
    For colonne = 1 To (2 * Sheets(strPage1).Cells(10, 7).Value)
        If colonne Mod 2 = 1 Then
            ' Formatage colonne paire
            nombreEleves = Sheets(strPage1).Cells(12 + (colonne + 1) / 2, 7).Value
            Columns(colonne).ColumnWidth = 40
            For ligneBtn = 1 To 2
                Set buttonCell = Cells(ligneBtn, colonne)
                Set Button = ActiveSheet.Buttons.Add(buttonCell.Left, buttonCell.Top, buttonCell.Width, buttonCell.Height)
                If ligneBtn = 1 Then
                    With Button
                        .Caption = "Ajouter élève"
                        .OnAction = "btnAjouterEleve_Click"
                    End With
                Else
                    With Button
                        .Caption = "Supprimer élève"
                        .OnAction = "btnSupprimerEleve_Click"
                    End With
                End If
            Next ligneBtn
            With Cells(3, colonne)
                .Borders.ColorIndex = 1
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlMedium
                .Interior.ColorIndex = intColorClasse
                .Value = Sheets(strPage1).Cells(12 + (colonne + 1) / 2, 6).Value
                .HorizontalAlignment = xlHAlignCenter
                .VerticalAlignment = xlVAlignCenter
                .Locked = True
            End With
            With Range(Cells(4, colonne), Cells(3 + nombreEleves, colonne))
                .Borders.ColorIndex = 1
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
                .VerticalAlignment = xlVAlignCenter
                .Locked = False
            End With
        Else
            ' Formatage colonne impaire
            Columns(colonne).ColumnWidth = 5
        End If
    Next colonne
    
    ' Creation bouton "Créer Tableaux"
    Columns(colonne).ColumnWidth = 30
    Set buttonCell = Cells(1, colonne)
    Set Button = ActiveSheet.Buttons.Add(buttonCell.Left, buttonCell.Top, buttonCell.Width, buttonCell.Height)
    With Button
        .Caption = "Créer Tableaux"
        .OnAction = "btnCreerTableaux_Click"
    End With
    With Cells(2, colonne)
        .Value = "Après avoir rempli les listes"
        .Interior.ColorIndex = 3
        .Borders.ColorIndex = xlColorIndexAutomatic
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
    End With
    
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
    
    ' Confirmation
    If MsgBox("Êtes-vous sûr(e) de valider ces listes ? Vous pourrez toujours ajouter des élèves mais il sera impossible de recréer les tableaux.", vbYesNo) = vbYes Then
        
        Application.ScreenUpdating = False
    
        ' Creation des pages 'Notes' et 'Bilan'
        Dim indexClasse As Integer
        nombreClasses = getNombreClasses
        For indexClasse = 1 To nombreClasses
            creerTableauNotes Sheets(strPage1).Cells(12 + indexClasse, 6).Value, indexClasse, Sheets(strPage1).Cells(12 + indexClasse, 7).Value
            creerTableauBilan Sheets(strPage1).Cells(12 + indexClasse, 6).Value, indexClasse, Sheets(strPage1).Cells(12 + indexClasse, 7).Value
        Next indexClasse
        
        ' Verouillage des listes
        With Sheets(strPage2)
            .Unprotect strPassword
            .Buttons(.Buttons.Count).Delete
            .Cells(2, 2 * nombreClasses + 1).Delete xlShiftUp
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

' Retourne l'index de l'élève s'il est dans la liste de la classe donnée en argument, -1 sinon
' valeurExacte = True -> on cherche la place de l'élève donné en argument (supposant qu'il fait partie de la classe)
' valeurExacte = False -> on cherche où intégrer l'élève pour respecter l'ordre alphabéthique
Function chercherIndexEleve(nomComplet As String, indexClasse As Integer, valeurExacte As Boolean) As Integer
    chercherIndexEleve = -1
    nombreEleves = getNombreEleves(Sheets(strPage1).Cells(12 + indexClasse, 6).Value)
    For indexEleve = 1 To nombreEleves
        If Not (valeurExacte) Then
            If StrComp(nomComplet, Sheets(strPage2).Cells(3 + indexEleve, indexClasse * 2 - 1).Value) = -1 Then
                chercherIndexEleve = indexEleve
                Exit For
            Else
                If indexEleve = nombreEleves Then chercherIndexEleve = nombreEleves + 1
            End If
        Else
            If StrComp(nomComplet, Sheets(strPage2).Cells(3 + indexEleve, indexClasse * 2 - 1).Value) = 0 Then chercherIndexEleve = indexEleve
        End If
    Next indexEleve
    'MsgBox ("Indice de l'élève : " & chercherIndexEleve)
End Function

' Procédure d'ajout d'un élève
Sub btnAjouterEleve_Click()
    Dim indexClasse As Integer
    Dim nomEleve As String, prenomEleve As String
    
    ' Classe
    indexClasse = WorksheetFunction.RoundUp(Val(Right(Application.Caller, 1)) / 2, 0)
    nomClasse = Sheets(strPage1).Cells(12 + indexClasse, 6).Value

    ' Eleve
    nomEleve = InputBox("Nom de l'élève à ajouter :")
    prenomEleve = InputBox("Prénom de l'élève à ajouter :")
    nomComplet = UCase(nomEleve) & " " & StrConv(prenomEleve, vbProperCase)

    'Confirmation
    If MsgBox("Voulez vous ajouter l'élève '" & nomComplet & "' à la classe '" & nomClasse & "' ?", vbYesNo) = vbYes Then
        ajouterEleve indexClasse, nomEleve, prenomEleve
        MsgBox ("Élève ajouté !")
    Else
        MsgBox ("Opération annulée.")
    End If

End Sub

Sub ajouterEleve(indexClasse As Integer, nomEleve As String, prenomEleve As String, Optional intIndiceEleveNouveau As Integer)
    Dim nomClasse As String, nomComplet As String
    
    ' Données initiales
    nomComplet = UCase(nomEleve) & " " & StrConv(prenomEleve, vbProperCase)
    nomClasse = getNomClasse(indexClasse)
    Page3 = "Notes (" & nomClasse & ")"
    Page4 = "Bilan (" & nomClasse & ")"
    nombreDomaines = getNombreDomaines
    nombreCompetences = getNombreCompetences
    nombreEleves = getNombreEleves(nomClasse)
    intIndiceEleveNouveau = chercherIndexEleve(nomComplet, indexClasse, False)
    
    If Not (intIndiceEleveNouveau > 0 And intIndiceEleveNouveau <= nombreEleves + 1) Then
        MsgBox ("L'indice de l'élève n'est pas compris dans la classe.")
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' Ajout page 1 (accueil)
    Sheets(strPage1).Unprotect strPassword
    Sheets(strPage1).Cells(12 + indexClasse, 7).Value = nombreEleves + 1
    Sheets(strPage1).Protect strPassword
    
    ' Ajout page 2 (liste)
    With Sheets(strPage2)
        .Unprotect strPassword
        If intIndiceEleveNouveau > 2 And intIndiceEleveNouveau < nombreEleves + 1 Then
            .Cells(3 + intIndiceEleveNouveau, 2 * indexClasse - 1).Insert xlShiftDown, xlFormatFromLeftOrAbove
        ElseIf intIndiceEleveNouveau = nombreEleves + 1 Then
            .Cells(3 + intIndiceEleveNouveau - 1, 2 * indexClasse - 1).Insert xlShiftDown, xlFormatFromLeftOrAbove
            .Cells(3 + intIndiceEleveNouveau - 1, 2 * indexClasse - 1).Value = .Cells(3 + intIndiceEleveNouveau, 2 * indexClasse - 1).Value
        Else
            .Cells(3 + intIndiceEleveNouveau + 1, 2 * indexClasse - 1).Insert xlShiftDown, xlFormatFromRightOrBelow
            .Cells(3 + intIndiceEleveNouveau + 1, 2 * indexClasse - 1).Value = .Cells(3 + intIndiceEleveNouveau, 2 * indexClasse - 1).Value
        End If
        .Cells(3 + intIndiceEleveNouveau, 2 * indexClasse - 1).Value = nomComplet
        .Cells.Locked = True
        .EnableSelection = xlUnlockedCells
        .Protect strPassword
    End With
    
    If Sheets.Count > 3 Then
        nombreEval = Sheets(Page3).Buttons.Count - 1
        ' Ajout page 3 (notes)
        With Sheets(Page3)
            .Unprotect strPassword
            If intIndiceEleveNouveau > 2 And intIndiceEleveNouveau < nombreEleves + 1 Then
                .Range(.Cells(5 + intIndiceEleveNouveau, 1), .Cells(5 + intIndiceEleveNouveau, 2 + (nombreCompetences + 1) * nombreEval)).Insert xlDown, xlFormatFromLeftOrAbove
            ElseIf intIndiceEleveNouveau = nombreEleves + 1 Then
                .Range(.Cells(5 + intIndiceEleveNouveau - 1, 1), .Cells(5 + intIndiceEleveNouveau - 1, 2 + (nombreCompetences + 1) * nombreEval)).Insert xlDown, xlFormatFromLeftOrAbove
                .Range(.Cells(5 + intIndiceEleveNouveau - 1, 1), .Cells(5 + intIndiceEleveNouveau - 1, 2 + (nombreCompetences + 1) * nombreEval)).Value = .Range(.Cells(5 + intIndiceEleveNouveau, 1), .Cells(5 + intIndiceEleveNouveau, 2 + (nombreCompetences + 1) * nombreEval)).Value
                .Range(.Cells(5 + intIndiceEleveNouveau, 1), .Cells(5 + intIndiceEleveNouveau, 2 + (nombreCompetences + 1) * nombreEval)).ClearContents
                .Range(.Cells(5 + intIndiceEleveNouveau - 1, 1), .Cells(5 + intIndiceEleveNouveau - 1, 2)).MergeCells = True
            Else
                .Range(.Cells(5 + intIndiceEleveNouveau + 1, 1), .Cells(5 + intIndiceEleveNouveau + 1, 2 + (nombreCompetences + 1) * nombreEval)).Insert xlDown, xlFormatFromRightOrBelow
                .Range(.Cells(5 + intIndiceEleveNouveau + 1, 1), .Cells(5 + intIndiceEleveNouveau + 1, 2 + (nombreCompetences + 1) * nombreEval)).Value = .Range(.Cells(5 + intIndiceEleveNouveau, 1), .Cells(5 + intIndiceEleveNouveau, 2 + (nombreCompetences + 1) * nombreEval)).Value
                .Range(.Cells(5 + intIndiceEleveNouveau, 1), .Cells(5 + intIndiceEleveNouveau, 2 + (nombreCompetences + 1) * nombreEval)).ClearContents
                .Range(.Cells(5 + intIndiceEleveNouveau + 1, 1), .Cells(5 + intIndiceEleveNouveau + 1, 2)).MergeCells = True
            End If
            .Range(.Cells(5 + intIndiceEleveNouveau, 1), .Cells(5 + intIndiceEleveNouveau, 2)).MergeCells = True
            .Cells(5 + intIndiceEleveNouveau, 1).Value = nomComplet
            .EnableSelection = xlUnlockedCells
            .Protect strPassword
        End With
        
        ' Ajout page 4 (résultats)
        With Sheets(Page4)
            .Unprotect strPassword
            If intIndiceEleveNouveau > 2 And intIndiceEleveNouveau < nombreEleves + 1 Then
                .Range(.Cells(3 + intIndiceEleveNouveau, 1), .Cells(3 + intIndiceEleveNouveau, 1 + 4 * (nombreCompetences + 1))).Insert xlDown, xlFormatFromLeftOrAbove
            ElseIf intIndiceEleveNouveau = nombreEleves + 1 Then
                .Range(.Cells(3 + intIndiceEleveNouveau - 1, 1), .Cells(3 + intIndiceEleveNouveau - 1, 1 + 4 * (nombreCompetences + 1))).Insert xlDown, xlFormatFromLeftOrAbove
                .Range(.Cells(3 + intIndiceEleveNouveau - 1, 1), .Cells(3 + intIndiceEleveNouveau - 1, 1 + 4 * (nombreCompetences + 1))).Value = .Range(.Cells(3 + intIndiceEleveNouveau, 1), .Cells(3 + intIndiceEleveNouveau, 1 + 4 * (nombreCompetences + 1))).Value
                .Range(.Cells(3 + intIndiceEleveNouveau, 1), .Cells(3 + intIndiceEleveNouveau, 1 + 4 * (nombreCompetences + 1))).ClearContents
            Else
                .Range(.Cells(3 + intIndiceEleveNouveau + 1, 1), .Cells(3 + intIndiceEleveNouveau + 1, 1 + 4 * (nombreCompetences + 1))).Insert xlDown, xlFormatFromRightOrBelow
                .Range(.Cells(3 + intIndiceEleveNouveau + 1, 1), .Cells(3 + intIndiceEleveNouveau + 1, 1 + 4 * (nombreCompetences + 1))).Value = .Range(.Cells(3 + intIndiceEleveNouveau, 1), .Cells(3 + intIndiceEleveNouveau, 1 + 4 * (nombreCompetences + 1))).Value
                .Range(.Cells(3 + intIndiceEleveNouveau, 1), .Cells(3 + intIndiceEleveNouveau, 1 + 4 * (nombreCompetences + 1))).ClearContents
            End If
            .Cells(3 + intIndiceEleveNouveau, 1).Value = nomComplet
            .EnableSelection = xlUnlockedCells
            .Protect strPassword
        End With
    End If
    
    Application.ScreenUpdating = True
    
End Sub

Sub btnSupprimerEleve_Click()
    Dim indexClasse As Integer
    Dim nomEleve As String, prenomEleve As String, nomComplet As String
    
    ' Classe
    indexClasse = WorksheetFunction.RoundUp(Val(Right(Application.Caller, 1)) / 2, 0)
    nomClasse = getNomClasse(indexClasse)

    ' Eleve
    nomEleve = InputBox("Nom de l'élève à supprimer (comme spécifié dans la liste, écrire en minuscule si accent) :")
    prenomEleve = InputBox("Prénom de l'élève à supprimer (comme spécifié dans la liste, écrire en minuscule si accent) :")
    nomComplet = UCase(nomEleve) & " " & StrConv(prenomEleve, vbProperCase)

    'Confirmation
    If MsgBox("Voulez vous supprimer l'élève '" & nomComplet & "' de la classe '" & nomClasse & "' ?", vbYesNo) = vbYes Then
        If chercherIndexEleve(nomComplet, indexClasse, True) <> -1 Then
            supprimerEleve indexClasse, nomComplet
            MsgBox ("Élève supprimé avec succès.")
        Else
            MsgBox ("L'élève '" & nomComplet & "' n'a pas été trouvé dans la classe " & nomClasse & ". Vérifiez l'orthographe.")
        End If
    Else
        MsgBox ("Opération annulée.")
    End If

End Sub

Sub supprimerEleve(intIndiceClasse As Integer, strNomComplet As String)
    Dim intIndiceEleve As Integer
    Dim strClasse As String
    
    ' Données initiales
    strClasse = getNomClasse(intIndiceClasse)
    strPage3 = "Notes (" & strClasse & ")"
    strPage4 = "Bilan (" & strClasse & ")"
    intNombreDomaines = getNombreDomaines
    intNombreCompetences = getNombreCompetences
    intNombreEleves = getNombreEleves(strClasse)
    intIndiceEleve = chercherIndexEleve(strNomComplet, intIndiceClasse, True)
    
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




