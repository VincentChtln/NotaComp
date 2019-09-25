Attribute VB_Name = "Module1"
' Nom des pages & protection
Global Const Page0 As String = "Caché"
Global Const Page1 As String = "Page d'accueil"
Global Const Page2 As String = "Liste de classe"
Global Const Password As String = "Saint-Martin"
' Couleurs
Global Const colorindexDomaine As Integer = 10
Global Const colorindexDomaine2 As Integer = 35
Global Const colorindexClasse As Integer = 44
Global Const colorindexEval As Integer = colorindexClasse
Global Const colorindexNote As Integer = 33
Global Const colorindexNote2 As Integer = 34
Global Const colorindexBilan As Integer = colorindexClasse
' Valeurs min & max
Global Const nombreMinCompetences As Integer = 1
Global Const nombreMaxCompetences As Integer = 8
Global Const nombreMinEleves As Integer = 1
Global Const nombreMaxEleves As Integer = 40

' **********************************
' Page 1 (accueil) - Procédure & fonctions
' **********************************

Sub btnCreerDomaines_Click()
    nombreDomaines = Sheets("Page d'accueil").Cells(10, 3).Value
    If IsNumeric(nombreDomaines) And nombreDomaines > 2 And nombreDomaines < 11 Then
        creerDomaines (nombreDomaines)
        If Sheets("Page d'accueil").Cells(10, 7).Value > 0 Then
            creerBtnListeEleve
        End If
    Else
        MsgBox ("Veuillez entrer un nombre compris entre 3 et 10")
    End If
End Sub

Sub creerDomaines(nombreDomaines)
    Dim ligneDomaine As Integer, colonneDomaine As Integer
    ligneDomaine = 10
    colonneDomaine = 2
    
    ' Retrait protection
    Sheets(Page1).Unprotect Password
    
    ' Ajout des cellules & formatage
    Range(Cells(ligneDomaine + 1, colonneDomaine), Cells(ligneDomaine + 20, colonneDomaine + 1)).Delete Shift:=xlUp
    Cells(ligneDomaine + 2, colonneDomaine).Value = "Domaines"
    Cells(ligneDomaine + 2, colonneDomaine + 1).Value = "Nombre compétences"
    With Range(Cells(ligneDomaine + 2, colonneDomaine), Cells(ligneDomaine + 2, colonneDomaine + 1))
        .Interior.ColorIndex = colorindexDomaine
        .HorizontalAlignment = xlHAlignCenter
    End With
    For Index = 1 To nombreDomaines
        Cells(ligneDomaine + 2 + Index, colonneDomaine).Value = "Domaine " & Index
    Next Index
    With Range(Cells(ligneDomaine + 2, colonneDomaine), Cells(ligneDomaine + 2 + nombreDomaines, colonneDomaine + 1))
        .Borders.ColorIndex = xlColorIndexAutomatic
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
    End With
    Range(Cells(ligneDomaine + 3, colonneDomaine + 1), Cells(ligneDomaine + 2 + nombreDomaines, colonneDomaine + 1)).Cells.Locked = False
    
    ' Protection
    Sheets(Page1).EnableSelection = xlUnlockedCells
    Sheets(Page1).Protect Password
    
End Sub

Function getNombreDomaines()
    getNombreDomaines = Sheets(Page1).Cells(10, 3).Value
End Function

Function getNombreCompetences(Optional indexDomaine As Integer)

    nombreDomaines = getNombreDomaines
    If IsMissing(indexDomaine) Or indexDomaine = 0 Then
        result = Application.Sum(Sheets(Page1).Range(Sheets(Page1).Cells(13, 3), Sheets(Page1).Cells(12 + nombreDomaines, 3)))
    Else
        result = Application.VLookup("Domaine " & indexDomaine, Sheets(Page1).Range(Sheets(Page1).Cells(13, 2), Sheets(Page1).Cells(12 + nombreDomaines, 3)), 2, False)
    End If
    getNombreCompetences = result
    
End Function

Sub btnCreerClasses_Click()
    nombreClasses = Sheets(Page1).Cells(10, 7).Value
    If IsNumeric(nombreClasses) And nombreClasses > 0 And nombreClasses < 21 Then
        creerClasses (nombreClasses)
        If Sheets("Page d'accueil").Cells(10, 3).Value > 0 Then
            creerBtnListeEleve
        End If
    Else
        MsgBox ("Veuillez entrer un nombre compris entre 1 et 20")
    End If

End Sub

Sub creerClasses(nombreClasses)
    Dim ligneClasse As Integer, colonneClasse As Integer
    ligneClasse = 10
    colonneClasse = 6
    
    ' Retrait protection
    Sheets(Page1).Unprotect Password
    
    ' Ajout des cellules & formatage
    Range(Cells(ligneClasse + 2, colonneClasse), Cells(ligneClasse + 20, colonneClasse + 1)).Delete Shift:=xlUp
    Cells(ligneClasse + 2, colonneClasse).Value = "Nom de la classe"
    Cells(ligneClasse + 2, colonneClasse + 1).Value = "Nombre d'élèves"
    With Range(Cells(ligneClasse + 2, colonneClasse), Cells(ligneClasse + 2, colonneClasse + 1)):
        .Interior.ColorIndex = colorindexClasse
    End With
    With Range(Cells(ligneClasse + 2, colonneClasse), Cells(ligneClasse + 2 + nombreClasses, colonneClasse + 1))
        .Borders.ColorIndex = xlColorIndexAutomatic
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
    End With
    With Range(Cells(ligneClasse + 3, colonneClasse), Cells(ligneClasse + 2 + nombreClasses, colonneClasse + 1))
        .Cells.Locked = False
    End With

    ' Protection
    Sheets(Page1).EnableSelection = xlUnlockedCells
    Sheets(Page1).Protect Password
    
End Sub

Function getNombreClasses()
    getNombreClasses = Sheets(Page1).Cells(10, 7).Value
End Function

Function getNombreEleves(Optional nomClasse As String)
    nombreClasses = getNombreClasses
    If IsMissing(nomClasse) Then
        result = Application.Sum(Range(Sheets(Page1).Cells(13, 7), Cells(12 + nombreClasses, 7)))
    Else
        result = Application.VLookup(nomClasse, Sheets(Page1).Range(Sheets(Page1).Cells(13, 6), Sheets(Page1).Cells(12 + nombreClasses, 7)), 2, False)
    End If
    getNombreEleves = result
    
End Function

Sub creerBtnListeEleve()
    
    ' Retrait protection
    Sheets(Page1).Unprotect Password
    
    ' Ajout bouton
    Set buttonCell = Range("J10:J11")
    If Sheets(Page1).Buttons.Count > 2 Then
        For indexButton = 3 To Sheets(Page1).Buttons.Count
            Sheets(Page1).Buttons(indexButton).Delete
        Next indexButton
    End If
    Set Button = Sheets(Page1).Buttons.Add(buttonCell.Left, buttonCell.Top, buttonCell.Width, buttonCell.Height)
    With Button
        .Caption = "Valider les données & créer les listes"
        .OnAction = "btnCreerListeEleve_Click"
    End With

    ' Protection
    Sheets(Page1).EnableSelection = xlUnlockedCells
    Sheets(Page1).Protect Password
    
End Sub

Sub btnCreerListeEleve_Click()

    ' Verification données
    nombreDomaines = getNombreDomaines
    nombreClasses = getNombreClasses
    cmp = 0
    ' Test des domaines/compétences
    For indexDomaine = 1 To nombreDomaines
        If Not (IsEmpty(Sheets(Page1).Cells(12 + indexDomaine, 3).Value)) And IsNumeric(Sheets(Page1).Cells(12 + indexDomaine, 3).Value) Then
            cmp = cmp + 1
            If Sheets(Page1).Cells(12 + indexDomaine, 3).Value > nombreMaxCompetences Then Sheets(Page1).Cells(12 + indexDomaine, 3).Value = nombreMaxCompetences
        End If
    Next indexDomaine
    If cmp = nombreDomaines Then
        cmp = 0
    Else
        MsgBox ("Veuillez entrer des nombres valides pour les compétences")
        Exit Sub
    End If
    ' Test des classes/élèves
    For indexDomaine = 1 To nombreDomaines
        If Not (IsEmpty(Sheets(Page1).Cells(12 + indexDomaine, 6).Value)) And IsNumeric(Sheets(Page1).Cells(12 + indexDomaine, 7).Value) Then
            cmp = cmp + 1
            If Sheets(Page1).Cells(12 + indexDomaine, 7).Value > nombreMaxEleves Then Sheets(Page1).Cells(12 + indexDomaine, 7).Value = nombreMaxEleves
        End If
    Next indexDomaine
    If cmp <> nombreClasses Then
        MsgBox ("Veuillez entrer des valeurs pour le nombre d'élèves de chaque classe")
        Exit Sub
    End If
    
    ' Confirmation
    If MsgBox("Êtes-vous sûr(e) de valider ces données ? Il ne sera pas possible de les modifier par la suite.", vbYesNo) = vbYes Then
        ' Verrouillage toutes cellules + protection feuille
        Sheets(Page1).Unprotect Password
        Sheets(Page1).EnableSelection = xlUnlockedCells
        Sheets(Page1).Cells.Locked = True
        'Sheets(Page1).Buttons.Delete
        For ligne = 3 To 30
            Sheets(Page1).Rows(ligne).RowHeight = 20
        Next ligne
        Sheets(Page1).Protect Password
        
        ' Creation liste elèves
        creerListeEleve
        
        MsgBox ("Listes de classe créées avec succès !")
    Else
        MsgBox ("Opération annulée")
    End If
    
End Sub
