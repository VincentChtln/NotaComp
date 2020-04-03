Attribute VB_Name = "Module1"
' ##################################
' PAGE 1 (accueil)
' ##################################

Option Explicit

' **********************************
' CONSTANTES
' **********************************
' Version de l'outil
Global Const strVersion As String = "v 2.3"
' Nom des pages & protection
Global Const strPage1 As String = "Page d'accueil"
Global Const strPage2 As String = "Liste de classe"
Global Const strPassword As String = "Saint-Martin"
' Couleurs
Global Const intColorDomaine As Integer = 50
Global Const intColorDomaine2 As Integer = 35
Global Const intColorClasse As Integer = 44
Global Const intColorEval As Integer = intColorClasse
Global Const intColorNote As Integer = 33
Global Const intColorNote2 As Integer = 34
Global Const intColorBilan As Integer = intColorClasse
' Valeurs min & max
Global Const intNombreMinDomaines As Integer = 1
Global Const intNombreMaxDomaines As Integer = 7
Global Const intNombreMinCompetences As Integer = 1
Global Const intNombreMaxCompetences As Integer = 8
Global Const intNombreMinClasses As Integer = 1
Global Const intNombreMaxClasses As Integer = 20
Global Const intNombreMinEleves As Integer = 5
Global Const intNombreMaxEleves As Integer = 40
' Lignes & colonnes de référence
Const intLigDomaine As Integer = 10
Const intColDomaine As Integer = 2
Const intLigClasse As Integer = 10
Const intColClasse As Integer = 6


' **********************************
' PROCEDURES GENERALES
' freezePanes (wdw As Application, intIndiceColonne As Integer, intIndiceLigne As Integer)
' unloadAllUserForms ()
' **********************************

Sub freezePanes(wdw As Application, intIndiceColonne As Integer, intIndiceLigne As Integer)
    With wdw
        .SplitColumn = intIndiceColonne
        .SplitRow = intIndiceLigne
        .freezePanes = True
    End With
End Sub

Sub unloadAllUserForms()
    For Each frm In VBA.UserForms
        If TypeOf frm Is UserForm Then Unload frm
    Next frm
End Sub


' **********************************
' FONCTIONS
' **********************************
' getNombreDomaines() As Integer
' getNombreCompetences(Optional intIndiceDomaine As Integer) As Integer
' getNombreClasses() As Integer
' getNomClasse(intIndiceClasse As Integer) As String
' getIndiceClasse(strNomClasse As String) As Integer
' getNombreEleves(Optional strNomClasse As String, Optional intIndiceClasse As Integer) As Integer
' **********************************

Function getNombreDomaines() As Integer
    getNombreDomaines = Sheets(strPage1).Cells(10, 3).Value
End Function

Function getNombreCompetences(Optional intIndiceDomaine As Integer) As Integer
    Dim intNbDomaines As Integer
    intNbDomaines = getNombreDomaines
    getNombreCompetences = -1
    If intIndiceDomaine = 0 Then
        getNombreCompetences = Application.Sum(Sheets(strPage1).Range(Sheets(strPage1).Cells(13, 3), Sheets(strPage1).Cells(12 + intNbDomaines, 3)))
    ElseIf intIndiceDomaine >= intNombreMinDomaines And intIndiceDomaine <= intNombreMaxDomaines Then
        getNombreCompetences = Application.VLookup("Domaine " & intIndiceDomaine, Sheets(strPage1).Range(Sheets(strPage1).Cells(13, 2), Sheets(strPage1).Cells(12 + intNbDomaines, 3)), 2, False)
    Else
        MsgBox ("Fonction getNombreCompetences - Indice hors plage")
    End If
    
End Function

Function getNombreClasses() As Integer
    getNombreClasses = Sheets(strPage1).Cells(intLigClasse, intColClasse + 1).Value
End Function

Function getNomClasse(intIndiceClasse As Integer) As String
    getNomClasse = Sheets(strPage1).Cells(12 + intIndiceClasse, intColClasse).Value
End Function

Function getIndiceClasse(strNomClasse As String) As Integer
    Dim intIndice As Integer, intNombreClasses As Integer
    getIndiceClasse = 0
    intNombreClasses = getNombreClasses
    For intIndice = 1 To intNombreClasses
        If strNomClasse = getNomClasse(intIndice) Then getIndiceClasse = intIndice
    Next intIndice
End Function

Function getNombreEleves(Optional varClasse As Variant) As Integer
    Dim intNombreClasses As Integer
    intNombreClasses = getNombreClasses
    If IsMissing(varClasse) Then
        getNombreEleves = Application.Sum(Range(Sheets(strPage1).Cells(13, intColClasse + 1), Cells(12 + intNombreClasses, intColClasse + 1)))
    Else
        Select Case VarType(varClasse)
        Case vbInteger              ' Parameter is an Interger -> Class index
            getNombreEleves = Sheets(strPage1).Cells(12 + varClasse, intColClasse + 1).Value
        Case vbDouble               ' Parameter is a Double -> Class index
            getNombreEleves = Sheets(strPage1).Cells(12 + varClasse, intColClasse + 1).Value
        Case vbString               ' Parameter is a String -> Class name
            getNombreEleves = Application.VLookup(varClasse, Sheets(strPage1).Range(Sheets(strPage1).Cells(13, 6), Sheets(strPage1).Cells(12 + intNombreClasses, 7)), 2, False)
        Case Else                   ' Parameter is another type
            getNombreEleves = -1
        End Select
    End If
End Function

' **********************************
' PROCÉDURES
' **********************************

Sub btnCreerDomaines_Click()
    Dim intNombreDomaines As Integer, intNombreClasses As Integer
    
    ' Retrait protection
    Application.ScreenUpdating = False
    Sheets(strPage1).Unprotect strPassword
    ' Conversion si nécessaire
    Sheets(strPage1).Cells(intLigDomaine, intColDomaine + 1).Value = Int(getNombreDomaines)
    ' Protection
    Sheets(strPage1).EnableSelection = xlUnlockedCells
    Sheets(strPage1).Protect strPassword
    Application.ScreenUpdating = True
    
    intNombreDomaines = getNombreDomaines
    If IsNumeric(intNombreDomaines) And intNombreDomaines >= intNombreMinDomaines And intNombreDomaines <= intNombreMaxDomaines Then
        Call creerDomaines(intNombreDomaines)
        intNombreClasses = getNombreClasses
        If intNombreClasses >= intNombreMinClasses And intNombreClasses <= intNombreMaxClasses Then
            Call creerBtnListeEleve
        End If
    Else
        MsgBox ("Veuillez entrer un nombre compris entre " & intNombreMinDomaines & " et " & intNombreMaxDomaines)
    End If
End Sub

Sub creerDomaines(intNombreDomaines As Integer)
    Dim intIndiceDomaine As Integer
    
    ' Retrait protection
    Application.ScreenUpdating = False
    Sheets(strPage1).Unprotect strPassword
    
    ' Ajout des cellules & formatage
    Range(Cells(intLigDomaine + 1, intColDomaine), Cells(intLigDomaine + intNombreMaxDomaines + 1, intColDomaine + 1)).Delete Shift:=xlUp
    Cells(intLigDomaine + 2, intColDomaine).Value = "Domaines"
    Cells(intLigDomaine + 2, intColDomaine + 1).Value = "Nombre compétences"
    With Range(Cells(intLigDomaine + 2, intColDomaine), Cells(intLigDomaine + 2, intColDomaine + 1))
        .Interior.ColorIndex = intColorDomaine
        .HorizontalAlignment = xlHAlignCenter
    End With
    For intIndiceDomaine = 1 To intNombreDomaines
        Cells(intLigDomaine + 2 + intIndiceDomaine, intColDomaine).Value = "Domaine " & intIndiceDomaine
    Next intIndiceDomaine
    With Range(Cells(intLigDomaine + 2, intColDomaine), Cells(intLigDomaine + 2 + intNombreDomaines, intColDomaine + 1))
        .Borders.ColorIndex = xlColorIndexAutomatic
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
    End With
    Range(Cells(intLigDomaine + 3, intColDomaine), Cells(intLigDomaine + 2 + intNombreDomaines, intColDomaine + 1)).Cells.Locked = False
    
    ' Protection
    Sheets(strPage1).EnableSelection = xlUnlockedCells
    Sheets(strPage1).Protect strPassword
    Application.ScreenUpdating = True
    
End Sub

Sub btnCreerClasses_Click()
    Dim intNombreClasses As Integer, intNombreDomaines As Integer
    intNombreClasses = getNombreClasses
            
    ' Retrait protection
    Application.ScreenUpdating = False
    Sheets(strPage1).Unprotect strPassword
    ' Conversion si nécessaire
    Sheets(strPage1).Cells(intLigClasse, intColClasse + 1).Value = Int(getNombreClasses)
    ' Protection
    Sheets(strPage1).EnableSelection = xlUnlockedCells
    Sheets(strPage1).Protect strPassword
    Application.ScreenUpdating = True
        
    If IsNumeric(intNombreClasses) And intNombreClasses >= intNombreMinClasses And intNombreClasses <= intNombreMaxClasses Then
        Call creerClasses(intNombreClasses)
        intNombreDomaines = getNombreDomaines
        If intNombreDomaines >= intNombreMinDomaines And intNombreDomaines <= intNombreMaxDomaines Then
            Call creerBtnListeEleve
        End If
    Else
        MsgBox ("Veuillez entrer un nombre compris entre " & intNombreMinClasses & " et " & intNombreMaxClasses)
    End If

End Sub

Sub creerClasses(intNombreClasses As Integer)
    
    ' Retrait protection
    Application.ScreenUpdating = False
    Sheets(strPage1).Unprotect strPassword
    
    ' Ajout des cellules & formatage
    Range(Cells(intLigClasse + 2, intColClasse), Cells(intLigClasse + 20, intColClasse + 1)).Delete Shift:=xlUp
    Cells(intLigClasse + 2, intColClasse).Value = "Nom de la classe"
    Cells(intLigClasse + 2, intColClasse + 1).Value = "Nombre d'élèves"
    With Range(Cells(intLigClasse + 2, intColClasse), Cells(intLigClasse + 2, intColClasse + 1)):
        .Interior.ColorIndex = intColorClasse
    End With
    With Range(Cells(intLigClasse + 2, intColClasse), Cells(intLigClasse + 2 + intNombreClasses, intColClasse + 1))
        .Borders.ColorIndex = xlColorIndexAutomatic
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
    End With
    With Range(Cells(intLigClasse + 3, intColClasse), Cells(intLigClasse + 2 + intNombreClasses, intColClasse + 1))
        .Cells.Locked = False
    End With

    ' Protection
    Sheets(strPage1).EnableSelection = xlUnlockedCells
    Sheets(strPage1).Protect strPassword
    Application.ScreenUpdating = True
    
End Sub

Sub creerBtnListeEleve()
    Dim rngBouton As Range
    Dim intIndiceBouton As Integer
    Dim btnCreerListeEleve As Variant
    
    ' Retrait protection
    Application.ScreenUpdating = False
    Sheets(strPage1).Unprotect strPassword
    
    ' Ajout bouton
    Set rngBouton = Range("J10:J11")
    If Sheets(strPage1).Buttons.Count > 2 Then
        For intIndiceBouton = 3 To Sheets(strPage1).Buttons.Count
            Sheets(strPage1).Buttons(intIndiceBouton).Delete
        Next intIndiceBouton
    End If
    Set btnCreerListeEleve = Sheets(strPage1).Buttons.Add(rngBouton.Left, rngBouton.Top, rngBouton.Width, rngBouton.Height)
    With btnCreerListeEleve
        .Caption = "Valider les données & créer les listes"
        .OnAction = "btnCreerListeEleve_Click"
    End With
    Set rngBouton = Nothing
    Set btnCreerListeEleve = Nothing

    ' Protection
    Sheets(strPage1).EnableSelection = xlUnlockedCells
    Sheets(strPage1).Protect strPassword
    Application.ScreenUpdating = True
    
End Sub

Sub btnCreerListeEleve_Click()
    Dim intNombreDomaines As Integer, intNombreCompetences As Integer, intIndiceDomaine As Integer
    Dim intNombreClasses As Integer, intNombreEleves As Integer, intIndiceClasse As Integer
    Dim intCompteur As Integer

    ' Données
    intNombreDomaines = getNombreDomaines
    intNombreClasses = getNombreClasses
    intCompteur = 0
    
    ' Test des domaines/compétences
    For intIndiceDomaine = 1 To intNombreDomaines
        intNombreCompetences = getNombreCompetences(intIndiceDomaine)
        If Not IsEmpty(intNombreCompetences) And IsNumeric(intNombreCompetences) Then
            If intNombreCompetences > intNombreMaxCompetences Then
                Sheets(strPage1).Cells(12 + intIndiceDomaine, 3).Value = intNombreMaxCompetences
            ElseIf intNombreCompetences < intNombreMinCompetences Then
                Sheets(strPage1).Cells(12 + intIndiceDomaine, 3).Value = intNombreMinCompetences
            End If
            intCompteur = intCompteur + 1
        End If
    Next intIndiceDomaine
    If intCompteur = intNombreDomaines Then
        intCompteur = 0
    Else
        MsgBox ("Veuillez entrer des valeurs correctes pour les compétences de chaque domaine.")
        Exit Sub
    End If
    
    ' Test des classes/élèves
    For intIndiceClasse = 1 To intNombreClasses
        intNombreEleves = getNombreEleves(intIndiceClasse)
        If Not IsEmpty(intNombreEleves) And IsNumeric(intNombreEleves) Then
            If intNombreEleves > intNombreMaxEleves Then
                Sheets(strPage1).Cells(12 + intIndiceClasse, 7).Value = intNombreMaxEleves
            ElseIf intNombreEleves < intNombreMinEleves Then
                Sheets(strPage1).Cells(12 + intIndiceClasse, 7).Value = intNombreMinEleves
            End If
            intCompteur = intCompteur + 1
        End If
    Next intIndiceClasse
    If intCompteur <> intNombreClasses Then
        MsgBox ("Veuillez entrer des valeurs correctes pour le nombre d'élèves de chaque classe.")
        Exit Sub
    End If
    
    ' Confirmation
    If MsgBox("Êtes-vous sûr(e) de valider ces données ? Il ne sera pas possible de les modifier par la suite.", vbYesNo) = vbYes Then
        
        ' Verrouillage toutes cellules + protection feuille
        Application.ScreenUpdating = False
        Sheets(strPage1).Unprotect strPassword
        Sheets(strPage1).EnableSelection = xlUnlockedCells
        Sheets(strPage1).Cells.Locked = True
        Sheets(strPage1).Buttons.Delete
        Sheets(strPage1).Range("A3:A40").Rows.RowHeight = 20
        Sheets(strPage1).Protect strPassword
        Application.ScreenUpdating = True
        
        ' Creation liste elèves
        creerListeEleve
        
        MsgBox ("Listes de classe créées avec succès !")
    Else
        MsgBox ("Opération annulée")
    End If
End Sub

