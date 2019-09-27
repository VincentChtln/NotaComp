Attribute VB_Name = "Module2"
' **********************************
' Page 2 (listes �l�ves) - Proc�dure & fonctions
' **********************************

Sub creerListeEleve()

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
                        .Caption = "Ajouter �l�ve"
                        .OnAction = "btnAjouterEleve_Click"
                    End With
                Else
                    With Button
                        .Caption = "Supprimer �l�ve"
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
    
    ' Creation bouton "Cr�er Tableaux"
    Columns(colonne).ColumnWidth = 30
    Set buttonCell = Cells(1, colonne)
    Set Button = ActiveSheet.Buttons.Add(buttonCell.Left, buttonCell.Top, buttonCell.Width, buttonCell.Height)
    With Button
        .Caption = "Cr�er Tableaux"
        .OnAction = "btnCreerTableaux_Click"
    End With
    With Cells(2, colonne)
        .Value = "Apr�s avoir rempli les listes"
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
End Sub

' *** Origine: bouton "Cr�er Tableaux"
' *** Action: cr�e la feuille de listes de classes et tous les tableaux 'Classes' et 'Eval'
Sub btnCreerTableaux_Click()
    
    ' Confirmation
    If MsgBox("�tes-vous s�r(e) de valider ces listes ? Vous pourrez toujours ajouter des �l�ves mais il sera impossible de recr�er les tableaux.", vbYesNo) = vbYes Then
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
            '.Buttons(.Buttons.Count).Delete
            .Cells(2, 2 * nombreClasses + 1).Delete xlShiftUp
            .Cells.Locked = True
            .EnableSelection = xlUnlockedCells
            .Protect strPassword
        End With
        
        MsgBox ("Tableaux de notes et de bilan cr��s avec succ�s !")
    Else
        MsgBox ("Op�ration annul�e.")
    End If
    
End Sub

' Retourne l'index de l'�l�ve s'il est dans la liste de la classe donn�e en argument, -1 sinon
' valeurExacte = True -> on cherche la place de l'�l�ve donn� en argument (supposant qu'il fait partie de la classe)
' valeurExacte = False -> on cherche o� int�grer l'�l�ve pour respecter l'ordre alphab�thique
Function chercherIndexEleve(nomComplet As String, indexClasse As Integer, valeurExacte As Boolean) As Integer
    Dim resultatPrec As Integer
    chercherIndexEleve = -1
    resultatPrec = -1
    nombreEleve = getNombreEleves(Sheets(strPage1).Cells(12 + indexClasse, 6).Value)
    For indexEleve = 1 To nombreEleve
        If Not (valeurExacte) Then
            If StrComp(nomComplet, Sheets(strPage2).Cells(3 + indexEleve, indexClasse * 2 - 1).Value) <> resultatPrec Then chercherIndexEleve = indexEleve - 1
        Else
            If StrComp(nomComplet, Sheets(strPage2).Cells(3 + indexEleve, indexClasse * 2 - 1).Value) = 0 Then chercherIndexEleve = indexEleve
        End If
    Next indexEleve
    MsgBox ("index de l'�l�ve : " & chercherIndexEleve)
End Function

' D�clenche la proc�dure d'ajout d'un �l�ve
Sub btnAjouterEleve_Click()
    Dim indexClasse As Integer
    Dim nomEleve As String, prenomEleve As String
    
    ' Classe
    indexClasse = WorksheetFunction.RoundUp(Val(Right(Application.Caller, 1)) / 2, 0)
    nomClasse = Sheets(strPage1).Cells(12 + indexClasse, 6).Value

    ' Eleve
    nomEleve = InputBox("Nom de l'�l�ve � ajouter :")
    prenomEleve = InputBox("Pr�nom de l'�l�ve � ajouter :")
    nomComplet = UCase(nomEleve) & " " & StrConv(prenomEleve, vbProperCase)

    'Confirmation
    If MsgBox("Voulez vous ajouter l'�l�ve '" & nomComplet & "' � la classe '" & nomClasse & "' ?", vbYesNo) = vbYes Then
        ajouterEleve indexClasse, nomEleve, prenomEleve
        MsgBox ("�l�ve ajout� !")
    Else
        MsgBox ("Op�ration annul�e.")
    End If

End Sub

Sub ajouterEleve(indexClasse As Integer, nomEleve As String, prenomEleve As String)
    nomComplet = UCase(nomEleve) & " " & StrConv(prenomEleve, vbProperCase)
    nomClasse = Sheets(strPage1).Cells(12 + indexClasse, 6).Value
    Page3 = "Notes (" & nomClasse & ")"
    Page4 = "Bilan (" & nomClasse & ")"
    nombreDomaines = getNombreDomaines
    nombreCompetences = getNombreCompetences
    nombreEleves = Sheets(strPage1).Cells(12 + indexClasse, 7).Value
    
    ' Ajout dans page 1 (accueil)
    Sheets(strPage1).Unprotect strPassword
    Sheets(strPage1).Cells(12 + indexClasse, 7).Value = nombreEleves + 1
    Sheets(strPage1).Protect strPassword
    
    ' Ajout dans page 2 (liste)
    With Sheets(strPage2)
        .Unprotect strPassword
        .Cells(3 + nombreEleves, 2 * indexClasse - 1).Insert xlShiftDown, xlFormatFromLeftOrAbove
        .Cells(3 + nombreEleves, 2 * indexClasse - 1).Value = .Cells(3 + nombreEleves + 1, 2 * indexClasse - 1).Value
        .Cells(3 + nombreEleves + 1, 2 * indexClasse - 1).Value = nomComplet
        .Cells.Locked = True
        .EnableSelection = xlUnlockedCells
        .Protect strPassword
    End With
    
    If Sheets.Count > 3 Then
        nombreEval = Sheets(Page3).Buttons.Count - 1
        ' Ajout dans page 3 (notes)
        With Sheets(Page3)
            .Unprotect strPassword
            .Range(.Cells(5 + nombreEleves, 1), .Cells(5 + nombreEleves, 2 + (nombreCompetences + 1) * nombreEval)).Insert xlDown, xlFormatFromLeftOrAbove
            .Range(.Cells(5 + nombreEleves, 1), .Cells(5 + nombreEleves, 2 + (nombreCompetences + 1) * nombreEval)).Value = .Range(.Cells(5 + nombreEleves + 1, 1), .Cells(5 + nombreEleves + 1, 2 + (nombreCompetences + 1) * nombreEval)).Value
            .Range(.Cells(5 + nombreEleves, 1), .Cells(5 + nombreEleves, 2)).MergeCells = True
            .Range(.Cells(5 + nombreEleves + 1, 1), .Cells(5 + nombreEleves + 1, 2 + (nombreCompetences + 1) * nombreEval)).ClearContents
            .Cells(5 + nombreEleves + 1, 1).Value = nomComplet
            .EnableSelection = xlUnlockedCells
            .Protect strPassword
        End With
        
        ' Ajout dans page 4 (r�sultats)
        With Sheets(Page4)
            .Unprotect strPassword
            .Range(.Cells(3 + nombreEleves, 1), .Cells(3 + nombreEleves, 1 + 4 * (nombreCompetences + 1))).Insert xlDown, xlFormatFromLeftOrAbove
            .Range(.Cells(3 + nombreEleves, 1), .Cells(3 + nombreEleves, 1 + 4 * (nombreCompetences + 1))).Value = .Range(.Cells(3 + nombreEleves + 1, 1), .Cells(3 + nombreEleves + 1, 1 + 4 * (nombreCompetences + 1))).Value
            .Range(.Cells(3 + nombreEleves + 1, 1), .Cells(3 + nombreEleves + 1, 1 + 4 * (nombreCompetences + 1))).ClearContents
            .Cells(3 + nombreEleves + 1, 1).Value = nomComplet
            .EnableSelection = xlUnlockedCells
            .Protect strPassword
        End With
    End If
    
End Sub

Sub btnSupprimerEleve_Click()
    Dim indexClasse As Integer
    Dim nomEleve As String, prenomEleve As String, nomComplet As String
    
    ' Classe
    indexClasse = WorksheetFunction.RoundUp(Val(Right(Application.Caller, 1)) / 2, 0)
    nomClasse = Sheets(strPage1).Cells(12 + indexClasse, 6).Value

    ' Eleve
    nomEleve = InputBox("Nom de l'�l�ve � supprimer (comme �crit dans la liste) :")
    prenomEleve = InputBox("Pr�nom de l'�l�ve � supprimer (comme �crit dans la liste) :")
    nomComplet = UCase(nomEleve) & " " & StrConv(prenomEleve, vbProperCase)

    'Confirmation
    If MsgBox("Voulez vous supprimer l'�l�ve '" & nomComplet & "' � la classe '" & nomClasse & "' ?", vbYesNo) = vbYes Then
        If chercherIndexEleve(nomComplet, indexClasse, True) <> -1 Then
            supprimerEleve indexClasse, nomEleve, prenomEleve
            MsgBox ("�l�ve supprim� !")
        Else
            MsgBox ("L'�l�ve '" & nomComplet & "' n'a pas �t� trouv� dans la classe " & nomClasse & ". Veuillez v�rifier l'orthographe.")
        End If
    Else
        MsgBox ("Op�ration annul�e.")
    End If

End Sub

Sub supprimerEleve(indexClasse, nomEleve, prenomEleve)

End Sub



