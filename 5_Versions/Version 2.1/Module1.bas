Attribute VB_Name = "Module1"
Global Const strVersion As String = "v 2.1"
' Nom des pages & protection
Global Const strPage1 As String = "Page d'accueil"
Global Const strPage2 As String = "Liste de classe"
Global Const strPassword As String = "Saint-Martin"
' Couleurs
Global Const intColorDomaine As Integer = 10
Global Const intColorDomaine2 As Integer = 35
Global Const intColorClasse As Integer = 44
Global Const intColorEval As Integer = intColorClasse
Global Const intColorNote As Integer = 33
Global Const intColorNote2 As Integer = 34
Global Const intColorBilan As Integer = intColorClasse
' Valeurs min & max
Global Const intNombreMinCompetences As Integer = 1
Global Const intNombreMaxCompetences As Integer = 8
Global Const intNombreMinEleves As Integer = 1
Global Const intNombreMaxEleves As Integer = 40

' **********************************
' Page 1 (accueil) - Procédure & fonctions
' **********************************

Sub btnCreerDomaines_Click()
    nombreDomaines = getNombreDomaines
    If IsNumeric(nombreDomaines) And nombreDomaines > 2 And nombreDomaines < 11 Then
        creerDomaines (nombreDomaines)
        If Sheets(strPage1).Cells(10, 7).Value > 0 Then
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
    Sheets(strPage1).Unprotect strPassword
    
    ' Ajout des cellules & formatage
    Range(Cells(ligneDomaine + 1, colonneDomaine), Cells(ligneDomaine + 20, colonneDomaine + 1)).Delete Shift:=xlUp
    Cells(ligneDomaine + 2, colonneDomaine).Value = "Domaines"
    Cells(ligneDomaine + 2, colonneDomaine + 1).Value = "Nombre compétences"
    With Range(Cells(ligneDomaine + 2, colonneDomaine), Cells(ligneDomaine + 2, colonneDomaine + 1))
        .Interior.ColorIndex = intColorDomaine
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
    Range(Cells(ligneDomaine + 3, colonneDomaine), Cells(ligneDomaine + 2 + nombreDomaines, colonneDomaine)).Cells.Locked = False
    
    ' Protection
    Sheets(strPage1).EnableSelection = xlUnlockedCells
    Sheets(strPage1).Protect strPassword
    
End Sub

Function getNombreDomaines()
    getNombreDomaines = Sheets(strPage1).Cells(10, 3).Value
End Function

Function getNombreCompetences(Optional indexDomaine As Integer)
    nombreDomaines = getNombreDomaines
    If IsMissing(indexDomaine) Or indexDomaine = 0 Then
        result = Application.Sum(Sheets(strPage1).Range(Sheets(strPage1).Cells(13, 3), Sheets(strPage1).Cells(12 + nombreDomaines, 3)))
    Else
        result = Application.VLookup("Domaine " & indexDomaine, Sheets(strPage1).Range(Sheets(strPage1).Cells(13, 2), Sheets(strPage1).Cells(12 + nombreDomaines, 3)), 2, False)
    End If
    getNombreCompetences = result
    
End Function

Sub btnCreerClasses_Click()
    nombreClasses = Sheets(strPage1).Cells(10, 7).Value
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
    
    Application.ScreenUpdating = False
    
    ' Retrait protection
    Sheets(strPage1).Unprotect strPassword
    
    ' Ajout des cellules & formatage
    Range(Cells(ligneClasse + 2, colonneClasse), Cells(ligneClasse + 20, colonneClasse + 1)).Delete Shift:=xlUp
    Cells(ligneClasse + 2, colonneClasse).Value = "Nom de la classe"
    Cells(ligneClasse + 2, colonneClasse + 1).Value = "Nombre d'élèves"
    With Range(Cells(ligneClasse + 2, colonneClasse), Cells(ligneClasse + 2, colonneClasse + 1)):
        .Interior.ColorIndex = intColorClasse
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
    Sheets(strPage1).EnableSelection = xlUnlockedCells
    Sheets(strPage1).Protect strPassword
    
    Application.ScreenUpdating = True
    
End Sub

Function getNombreClasses()
    getNombreClasses = Sheets(strPage1).Cells(10, 7).Value
End Function

Function getNomClasse(intIndiceClasse As Integer) As String
    getNomClasse = Sheets(strPage1).Cells(12 + intIndiceClasse, 6).Value
End Function

Function getIndiceClasse(strNomClasse As String) As Integer
    Dim intIndice As Integer
    getIndiceClasse = 0
    For intIndice = 1 To getNombreClasse
        If strNomClasse = getNomClasse(intIndice) Then getIndiceClasse = intIndice
    Next intIndice
End Function

Function getNombreEleves(Optional nomClasse As String)
    nombreClasses = getNombreClasses
    If IsMissing(nomClasse) Then
        result = Application.Sum(Range(Sheets(strPage1).Cells(13, 7), Cells(12 + nombreClasses, 7)))
    Else
        result = Application.VLookup(nomClasse, Sheets(strPage1).Range(Sheets(strPage1).Cells(13, 6), Sheets(strPage1).Cells(12 + nombreClasses, 7)), 2, False)
    End If
    getNombreEleves = result
    
End Function

Sub creerBtnListeEleve()
    
    ' Retrait protection
    Sheets(strPage1).Unprotect strPassword
    
    ' Ajout bouton
    Set buttonCell = Range("J10:J11")
    If Sheets(strPage1).Buttons.Count > 2 Then
        For indexButton = 3 To Sheets(strPage1).Buttons.Count
            Sheets(strPage1).Buttons(indexButton).Delete
        Next indexButton
    End If
    Set Button = Sheets(strPage1).Buttons.Add(buttonCell.Left, buttonCell.Top, buttonCell.Width, buttonCell.Height)
    With Button
        .Caption = "Valider les données & créer les listes"
        .OnAction = "btnCreerListeEleve_Click"
    End With

    ' Protection
    Sheets(strPage1).EnableSelection = xlUnlockedCells
    Sheets(strPage1).Protect strPassword
    
End Sub

Sub btnCreerListeEleve_Click()

    ' Verification données
    nombreDomaines = getNombreDomaines
    nombreClasses = getNombreClasses
    cmp = 0
    ' Test des domaines/compétences
    For indexDomaine = 1 To nombreDomaines
        If Not (IsEmpty(Sheets(strPage1).Cells(12 + indexDomaine, 3).Value)) And IsNumeric(Sheets(strPage1).Cells(12 + indexDomaine, 3).Value) Then
            cmp = cmp + 1
            If Sheets(strPage1).Cells(12 + indexDomaine, 3).Value > intNombreMaxCompetences Then Sheets(strPage1).Cells(12 + indexDomaine, 3).Value = intNombreMaxCompetences
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
        If Not (IsEmpty(Sheets(strPage1).Cells(12 + indexDomaine, 6).Value)) And IsNumeric(Sheets(strPage1).Cells(12 + indexDomaine, 7).Value) Then
            cmp = cmp + 1
            If Sheets(strPage1).Cells(12 + indexDomaine, 7).Value > intNombreMaxEleves Then Sheets(strPage1).Cells(12 + indexDomaine, 7).Value = intNombreMaxEleves
        End If
    Next indexDomaine
    If cmp <> nombreClasses Then
        MsgBox ("Veuillez entrer des valeurs pour le nombre d'élèves de chaque classe")
        Exit Sub
    End If
    
    ' Confirmation
    If MsgBox("Êtes-vous sûr(e) de valider ces données ? Il ne sera pas possible de les modifier par la suite.", vbYesNo) = vbYes Then
        ' Verrouillage toutes cellules + protection feuille
        Sheets(strPage1).Unprotect strPassword
        Sheets(strPage1).EnableSelection = xlUnlockedCells
        Sheets(strPage1).Cells.Locked = True
        'Sheets(strPage1).Buttons.Delete
        For ligne = 3 To 30
            Sheets(strPage1).Rows(ligne).RowHeight = 20
        Next ligne
        Sheets(strPage1).Protect strPassword
        
        ' Creation liste elèves
        creerListeEleve
        
        MsgBox ("Listes de classe créées avec succès !")
    Else
        MsgBox ("Opération annulée")
    End If
    
End Sub

