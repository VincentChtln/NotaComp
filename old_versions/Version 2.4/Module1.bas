Attribute VB_Name = "Module1"
' *******************************************************************************
'                               NotaComp
'
'   Outil Excel dédié à la notation par compétence en milieu scolaire
'
'   Classeur vierge, documentation et fichiers source disponibles
'   sur le site <https://github.com/VincentChtln/NotaComp>
'
'   V1      Version initiale
'   V1.1    Amélioration de plusieurs fonctionnalités
'   V2      Refonte graphique et fonctionnelle
'   V2.4    Ajout de UserForm pour la modification des listes de classes
'   V2.5    Ajout de UserForm pour la modification des évaluations, modifications graphiques
'
' *******************************************************************************
'                       GNU General Public License V3
'   Copyright (C)
'   Date: 2021
'   Auteur: Vincent Chatelain
'
'   This file is part of NotaComp.
'
'   NotaComp is free software: you can redistribute it and/or modify
'   it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or
'   (at your option) any later version.
'
'   NotaComp is distributed in the hope that it will be useful,
'   but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
'   GNU General Public License for more details.
'
'   You should have received a copy of the GNU General Public License
'   along with NotaComp. If not, see <https://www.gnu.org/licenses/>.
'
'
'               GNU General Public License V3 - Traduction française
'
'   Ce fichier fait partie de NotaComp.
'
'   NotaComp est un logiciel libre ; vous pouvez le redistribuer ou le modifier
'   suivant les termes de la GNU General Public License telle que publiée par la
'   Free Software Foundation, soit la version 3 de la Licence, soit (à votre gré)
'   toute version ultérieure.
'
'   NotaComp est distribué dans l’espoir qu’il sera utile, mais SANS AUCUNE
'   GARANTIE : sans même la garantie implicite de COMMERCIALISABILITÉ
'   ni d’ADÉQUATION À UN OBJECTIF PARTICULIER. Consultez la GNU
'   General Public License pour plus de détails.
'
'   Vous devriez avoir reçu une copie de la GNU General Public License avec NotaComp;
'   si ce n’est pas le cas, consultez : <http://www.gnu.org/licenses/>.
'
' *******************************************************************************


' *******************************************************************************
'
'                               Page 1 - Accueil
'
'   Fonctions globales
'       getNomPage3(intIndiceClasse as Integer) as String
'       getNomPage4(intIndiceClasse as Integer) as String
'
'   Procédures globales
'       protectWorksheet(Optional intIndiceClasse As Integer = 0)
'       unprotectWorksheet(Optional intIndiceClasse As Integer = 0)
'       protectWorkbook()
'       unprotectWorkbook()
'       freezePanes (wdw As Application, intIndiceLigne As Integer, intIndiceColonne As Integer)
'       limitScrollArea(ws As Worksheet)
'       removeScrollArea(Optional ws As Worksheet)
'       unloadAllUserForms ()
'
'   Fonctions locales
'       getNombreDomaines() As Integer
'       getNombreCompetences(Optional intIndiceDomaine As Integer) As Integer
'       getNombreClasses() As Integer
'       getNomClasse(intIndiceClasse As Integer) As String
'       getIndiceClasse(strNomWorksheet As String) As Integer
'       getNombreEleves(Optional strNomClasse As String, Optional intIndiceClasse As Integer) As Integer
'
'   Procédures locales
'       preInitNotaComp()
'       btnDemarrerNotacomp_Click()
'       demarrerNotaComp()
'       setNombreEleves(intIndiceClasse As Integer, intNouveauNombreEleves As Integer)
'       btnCreerDomaines_Click()
'       creerDomaines(intNbDomaines As Integer)
'       btnCreerClasses_Click()
'       creerClasses(intNbClasses As Integer)
'
' Lien pour surlignage ligne: https://answers.microsoft.com/en-us/msoffice/forum/msoffice_excel-mso_other-mso_2016/vba-excel-highlighting-active-row-in-protected/c292192d-c624-4a20-9acd-ddc19c7d2a32
' *******************************************************************************

Option Explicit

' *******************************************************************************
'                                   Constantes
' *******************************************************************************

' *** VERSION ***
Global Const strVersion As String = "v 2.4"

' *** DONNEES PROTECTION ***
Global Const strPage1 As String = "Accueil"
Global Const strPage2 As String = "Liste"
Global Const strPassword As String = ""

' *** COULEURS ***
Const intColorInfos As Integer = 48
Const intColorVersion As Integer = 37
Global Const intColorDomaine As Integer = 50
Global Const intColorDomaine2 As Integer = 35
Global Const intColorClasse As Integer = 44
Global Const intColorEval As Integer = 45
Global Const intColorEval2 As Integer = 40
Global Const intColorNote As Integer = 33
Global Const intColorNote2 As Integer = 34
Global Const intColorBilan As Integer = 44

' *** VALEURS MIN ET MAX*
Global Const intNbMinDomaines As Integer = 1
Global Const intNbMaxDomaines As Integer = 7
Global Const intNbMinCompetences As Integer = 1
Global Const intNbMaxCompetences As Integer = 8
Global Const intNbMinClasses As Integer = 1
Global Const intNbMaxClasses As Integer = 20
Global Const intNbMinEleves As Integer = 5
Global Const intNbMaxEleves As Integer = 40

' *** LIGNES ET COLONNES REFERENCE ***
Const intLigInfos As Integer = 4
Const intColInfos As Integer = 2
Const intLigVersion As Integer = 4
Const intColVersion As Integer = 6
Const intLigDomaine As Integer = 10
Const intColDomaine As Integer = 2
Const intLigClasse As Integer = 10
Const intColClasse As Integer = 6
Global Const intLigListePage2 As Integer = 1
Global Const intLigListePage3 As Integer = 5
Global Const intLigListePage4 As Integer = 3

' *******************************************************************************
'                               Fonctions générales
' *******************************************************************************

Function getNomPage3(intIndiceClasse As Integer) As String
    getNomPage3 = "Notes (" & getNomClasse(intIndiceClasse) & ")"
End Function

Function getNomPage4(intIndiceClasse As Integer) As String
    getNomPage4 = "Bilan (" & getNomClasse(intIndiceClasse) & ")"
End Function

' *******************************************************************************
'                               Procédures générales
' *******************************************************************************

Sub protectWorksheet(Optional intIndiceClasse As Integer = 0)
    ' *** DECLARATION VARIABLES ***
    Dim ws As Worksheet
    Dim strNomClasse As String
    
    ' *** AFFECTATION VARIABLES ***
    If intIndiceClasse = 0 Then
        strNomClasse = vbNullString
    Else
        strNomClasse = getNomClasse(intIndiceClasse)
    End If
    
    ' *** PROTECTION ON ***
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = strPage1 Or ws.Name = strPage2 Or InStr(ws.Name, strNomClasse) <> 0 Then
            ws.EnableSelection = xlUnlockedCells
            ws.Protect Password:=strPassword, UserInterFaceOnly:=True
        End If
    Next ws
End Sub

Sub unprotectWorksheet(Optional intIndiceClasse As Integer = 0)
    ' *** DECLARATION VARIABLES ***
    Dim ws As Worksheet
    Dim strNomClasse As String

    ' *** AFFECTATION VARIABLES ***
    If intIndiceClasse = 0 Then
        strNomClasse = vbNullString
    Else
        strNomClasse = getNomClasse(intIndiceClasse)
    End If

    ' *** PROTECTION OFF ***
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = strPage1 Or ws.Name = strPage2 Or InStr(ws.Name, strNomClasse) <> 0 Then
            ws.Unprotect strPassword
        End If
    Next ws
End Sub

Sub protectWorkbook()
    ThisWorkbook.Protect strPassword, True, True
End Sub

Sub unprotectWorkbook()
    ThisWorkbook.Unprotect strPassword
End Sub

Sub freezePanes(wdw As Window, intIndiceLigne As Integer, intIndiceColonne As Integer)
    ' *** BLOQUAGE VOLETS ***
    With wdw
        .SplitRow = intIndiceLigne
        .SplitColumn = intIndiceColonne
        .freezePanes = True
    End With
End Sub

Sub removeScrollArea(Optional wsSource As Worksheet)
    If Not IsMissing(wsSource) Then
        wsSource.ScrollArea = vbNullString
    Else
        Dim ws As Worksheet
        For Each ws In Worksheets
            ws.ScrollArea = vbNullString
        Next ws
    End If
End Sub

Sub limitScrollArea(ws As Worksheet)
    Dim intLastRow As Integer
    Dim intLastCol As Integer
    Dim strLastCol As String
    With ws
        intLastRow = 10 + .Cells(.Rows.Count, "A").End(xlUp).Row
        intLastCol = 10 + .Cells(3, .Columns.Count).End(xlToLeft).Column
        strLastCol = Split(.Cells(1, intLastCol).Address, "$")(1)
        .ScrollArea = "A1:" & strLastCol & CStr(intLastRow)
    End With
End Sub

Sub unloadAllUserForms()
    ' *** DECLARATION VARIABLES ***
    Dim frm As UserForm
    
    ' *** UNLOAD USERFORMS ***
    For Each frm In VBA.UserForms
        If TypeOf frm Is UserForm Then Unload frm
    Next frm
End Sub

' *******************************************************************************
'                               Fonctions locales
' *******************************************************************************

Function isWorksheetProtected(ws As Worksheet) As Boolean
    With ws
        isWorksheetProtected = .ProtectContents Or .ProtectDrawingObjects Or .ProtectScenarios
    End With
End Function

Function isWorkbookProtected(wb As Workbook) As Boolean
    With wb
        isWorkbookProtected = .ProtectWindows Or .ProtectStructure
    End With
End Function

Function getNombreDomaines() As Integer
    getNombreDomaines = Worksheets(strPage1).Cells(intLigDomaine, intColDomaine + 1).Value
End Function

Function getNombreCompetences(Optional intIndiceDomaine As Integer) As Integer
    ' *** DECLARATION VARIABLES ***
    Dim intNbDomaines As Integer
    
    ' *** AFFECTATION VARIABLES ***
    intNbDomaines = getNombreDomaines
    getNombreCompetences = -1
    
    ' *** CALCUL ***
    If intIndiceDomaine = 0 Then
        getNombreCompetences = Application.Sum(Worksheets(strPage1).Range(Worksheets(strPage1).Cells(intLigDomaine + 3, intColDomaine + 1), Worksheets(strPage1).Cells(intLigDomaine + intNbDomaines + 1, intColDomaine + 1)))
    ElseIf intIndiceDomaine >= intNbMinDomaines And intIndiceDomaine <= intNbMaxDomaines Then
        getNombreCompetences = Application.VLookup("Domaine " & intIndiceDomaine, Worksheets(strPage1).Range(Worksheets(strPage1).Cells(intLigDomaine + 3, intColDomaine + 1), Worksheets(strPage1).Cells(intLigDomaine + intNbDomaines + 1, intColDomaine + 1)), 2, False)
    Else
        MsgBox ("Fonction getNombreCompetences - Indice hors plage")
    End If
End Function

Function getNombreClasses() As Integer
    getNombreClasses = Worksheets(strPage1).Cells(intLigClasse, intColClasse + 1).Value
End Function

Function getNomClasse(intIndiceClasse As Integer) As String
    getNomClasse = Worksheets(strPage1).Cells(12 + intIndiceClasse, intColClasse).Value
End Function

Function getIndiceClasse(strNomWorksheet As String) As Integer
    ' *** DECLARATION VARIABLES ***
    Dim intIndiceClasse As Integer
    Dim intNbClasses As Integer
    
    ' *** AFFECTATION VARIABLES ***
    getIndiceClasse = 0
    intNbClasses = getNombreClasses
    
    ' *** CALCUL ***
    For intIndiceClasse = 1 To intNbClasses
        If InStr(strNomWorksheet, getNomClasse(intIndiceClasse)) <> 0 Then getIndiceClasse = intIndiceClasse
    Next intIndiceClasse
End Function

Function getNombreEleves(Optional intIndiceClasse As Integer) As Integer
    ' *** DECLARATION VARIABLES ***
    Dim intNbClasses As Integer
    
    ' *** AFFECTATION VARIABLES ***
    intNbClasses = getNombreClasses
    
    ' *** CALCUL ***
    With Worksheets(strPage1)
        If IsMissing(intIndiceClasse) Then
            getNombreEleves = Application.Sum(.Range(.Cells(13, intColClasse + 1), .Cells(12 + intNbClasses, intColClasse + 1)))
        Else
            getNombreEleves = .Cells(12 + intIndiceClasse, intColClasse + 1).Value
        End If
    End With
End Function

Function isInteger(ByVal Value As Variant) As Boolean
    isInteger = False
    On Error Resume Next
    isInteger = (Int(Value) = Value)
    On Error GoTo 0
End Function

' *******************************************************************************
'                               Procédures locales
' *******************************************************************************

Sub preInitNotaComp()
    ' *** DECLARATION VARIABLES ***
    Dim rngBtnDemarrerNotacomp As Range
    Dim btnDemarrerNotacomp As Variant
    Dim wb As Workbook
    
    ' *** REFRESH ECRAN OFF ***
    Application.ScreenUpdating = False

    ' *** VERIFICATION PROTECTION WORKBOOK ***
    Set wb = ThisWorkbook
    If isWorkbookProtected(wb) Then
        MsgBox "Ce classeur est protégé, merci de désactiver la protection avant de continuer."
        GoTo EOP
    End If
        
    With ThisWorkbook
        ' *** VERIFICATION WORKSHEET UNIQUE ***
        If .Worksheets.Count <> 1 Then
            MsgBox "Ce classeur contient plusieurs feuilles, merci de les supprimer avant de continuer."
            GoTo EOP
        End If
        
        ' *** VERIFICATION PROTECTION FEUILLE ***
        If isWorksheetProtected(.Worksheets(1)) Then
            MsgBox "Cette feuille est protégée, merci de désactiver la protection avant de continuer."
            GoTo EOP
        End If
            
        With .Worksheets(1)
            ' *** VERIFICATION FEUILLE VIDE ***
            If Not (.UsedRange.Address = "$A$1" And .Range("A1") = "") Then
                Select Case MsgBox("La feuille n'est pas vide, cliquer sur 'OK' pour supprimer son contenu et continuer, ou sur 'Annuler' pour revenir en arrière", vbOKCancel)
                Case vbOK
                    .Cells.Clear
                Case vbCancel
                    GoTo EOP
                End Select
            End If
            
            ' *** MISE EN PAGE ***
            .Rows().RowHeight = 20
            .Rows(2).RowHeight = 50
            .Columns().ColumnWidth = 25
            .Range(.Columns(1), .Columns(5), .Columns(9)).ColumnWidth = 10
            With .Range("B2")
                .Font.Size = 48
                .Value = "NotaComp"
            End With
            .Name = strPage1
            
            ' *** AJOUT BOUTON 'DEMARRER NOTACOMP' ***
            Set rngBtnDemarrerNotacomp = .Range("D2")
            Set btnDemarrerNotacomp = .Buttons.Add(rngBtnDemarrerNotacomp.Left, rngBtnDemarrerNotacomp.Top, rngBtnDemarrerNotacomp.Width, rngBtnDemarrerNotacomp.Height)
            With btnDemarrerNotacomp
                .Caption = "Démarrer NotaComp"
                .OnAction = "btnDemarrerNotacomp_Click"
            End With
        End With
    End With
    
    ' *** PROTECTION ON ***
    protectWorksheet
    
EOP:
    ' *** REFRESH ECRAN ON ***
    Application.ScreenUpdating = True
     
End Sub

Sub btnDemarrerNotacomp_Click()
    ' *** PROTECTION + REFRESH ECRAN OFF ***
    Application.ScreenUpdating = False
    unprotectWorksheet
    
    ' *** APPEL PROCEDURE ***
    demarrerNotacomp

    ' *** PROTECTION + REFRESH ECRAN ON ***
    protectWorksheet
    Application.ScreenUpdating = True
End Sub

Sub demarrerNotacomp()
' Création tableau infos
' Création tableau version
' Utiliser .MergeCells = True
    MsgBox Application.UserName
End Sub

Sub setNombreEleves(intIndiceClasse As Integer, intNouveauNombreEleves As Integer)
    Worksheets(strPage1).Cells(12 + intIndiceClasse, intColClasse + 1).Value = intNouveauNombreEleves
End Sub

Sub btnCreerDomaines_Click()
    ' *** DECLARATION VARIABLES ***
    Dim intNbDomaines As Integer
    Dim intNbClasses As Integer
    
    ' *** PROTECTION + REFRESH ECRAN OFF ***
    Application.ScreenUpdating = False
    unprotectWorksheet
    
    ' *** AFFECTATION VARIABLES ***
    intNbDomaines = getNombreDomaines
    
    ' *** VERIFICATION TYPAGE ***
    If Not IsNumeric(intNbDomaines) Then
        MsgBox "Veuillez entrer un nombre valide"
        GoTo EOP
    End If
    
    ' *** FORMATAGE VALEUR ***
    intNbDomaines = CInt(intNbDomaines)
    Worksheets(strPage1).Cells(intLigDomaine, intColDomaine + 1).Value = intNbDomaines
    
    ' *** VERIFICATION VALEUR ***
    If intNbDomaines < intNbMinDomaines Or intNbDomaines > intNbMaxDomaines Then
        MsgBox ("Veuillez entrer un nombre compris entre " & intNbMinDomaines & " et " & intNbMaxDomaines)
        GoTo EOP
    End If
    
    ' *** CREATION DOMAINES ***
    creerDomaines intNbDomaines
    
    ' *** AJOUT BOUTON 'CREER LISTE ELEVE' ***
    intNbClasses = getNombreClasses
    If intNbClasses >= intNbMinClasses And intNbClasses <= intNbMaxClasses Then
        Call creerBtnListeEleve
    End If
    
    ' *** PROTECTION + REFRESH ECRAN ON ***
EOP:
    protectWorksheet
    Application.ScreenUpdating = True
End Sub

Sub creerDomaines(intNbDomaines As Integer)
    ' *** DECLARATION VARIABLES ***
    Dim intIndiceDomaine As Integer
    
    ' *** AJOUT CELLULES ***
    With Worksheets(strPage1)
        .Range(.Cells(intLigDomaine + 1, intColDomaine), .Cells(intLigDomaine + intNbMaxDomaines + 1, intColDomaine + 1)).Delete Shift:=xlUp
        .Cells(intLigDomaine + 2, intColDomaine).Value = "Domaines"
        .Cells(intLigDomaine + 2, intColDomaine + 1).Value = "Nombre compétences"
        With .Range(.Cells(intLigDomaine + 2, intColDomaine), .Cells(intLigDomaine + 2, intColDomaine + 1))
            .Interior.ColorIndex = intColorDomaine
            .HorizontalAlignment = xlHAlignCenter
        End With
        For intIndiceDomaine = 1 To intNbDomaines
            .Cells(intLigDomaine + 2 + intIndiceDomaine, intColDomaine).Value = "Domaine " & intIndiceDomaine
        Next intIndiceDomaine
        With .Range(.Cells(intLigDomaine + 2, intColDomaine), .Cells(intLigDomaine + 2 + intNbDomaines, intColDomaine + 1))
            .Borders.ColorIndex = xlColorIndexAutomatic
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
        End With
        .Range(.Cells(intLigDomaine + 3, intColDomaine), .Cells(intLigDomaine + 2 + intNbDomaines, intColDomaine + 1)).Cells.Locked = False
        .Range("A3:A40").Rows.RowHeight = 20
    End With
End Sub

Sub btnCreerClasses_Click()
    ' *** DECLARATION VARIABLES ***
    Dim intNbClasses As Integer
    Dim intNbDomaines As Integer
    
    ' *** PROTECTION + REFRESH ECRAN OFF
    Application.ScreenUpdating = False
    unprotectWorksheet
    
    ' *** AFFECTATION VARIABLES ***
    intNbClasses = getNombreClasses
    
    ' *** VERIFICATION TYPAGE ***
    If Not IsNumeric(intNbClasses) Then
        MsgBox "Veuillez entrer un nombre valide"
        GoTo EOP
    End If
    
    ' *** FORMATAGE VALEUR ***
    intNbClasses = CInt(intNbClasses)
    Worksheets(strPage1).Cells(intLigClasse, intColClasse + 1).Value = intNbClasses
    
    ' *** VERIFICATION VALEUR ***
    If intNbClasses < intNbMinClasses Or intNbClasses > intNbMaxClasses Then
        MsgBox ("Veuillez entrer un nombre compris entre " & intNbMinClasses & " et " & intNbMaxClasses)
        GoTo EOP
    End If
    
    ' *** CREATION CLASSES ***
    creerClasses intNbClasses
    
    ' *** AJOUT BOUTON 'CREER LISTE ELEVES' ***
    intNbDomaines = getNombreDomaines
    If intNbDomaines >= intNbMinDomaines And intNbDomaines <= intNbMaxDomaines Then
        creerBtnListeEleve
    End If
    
    ' *** PROTECTION + REFRESH ECRAN ON ***
EOP:
    protectWorksheet
    Application.ScreenUpdating = True
End Sub

Sub creerClasses(intNbClasses As Integer)
    
    ' *** AJOUT CELLULES ET MISE EN FORME ***
    With Worksheets(strPage1)
        .Range(.Cells(intLigClasse + 2, intColClasse), .Cells(intLigClasse + 20, intColClasse + 1)).Delete Shift:=xlUp
        .Cells(intLigClasse + 2, intColClasse).Value = "Nom de la classe"
        .Cells(intLigClasse + 2, intColClasse + 1).Value = "Nombre d'élèves"
        With .Range(.Cells(intLigClasse + 2, intColClasse), .Cells(intLigClasse + 2, intColClasse + 1)):
            .Interior.ColorIndex = intColorClasse
        End With
        With .Range(.Cells(intLigClasse + 2, intColClasse), .Cells(intLigClasse + 2 + intNbClasses, intColClasse + 1))
            .Borders.ColorIndex = xlColorIndexAutomatic
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
        End With
        .Range(.Cells(intLigClasse + 3, intColClasse), .Cells(intLigClasse + 2 + intNbClasses, intColClasse + 1)).Cells.Locked = False
        .Range("A3:A40").Rows.RowHeight = 20
    End With
End Sub

Sub creerBtnListeEleve()
    ' *** DECLARATION VARIABLES ***
    Dim intIndiceBouton As Integer
    Dim rngBtnCreerListeEleve As Range
    Dim btnCreerListeEleve As Variant
    
    ' *** SUPPRESSION BOUTONS EN TROP ***
    If Worksheets(strPage1).Buttons.Count > 2 Then
        For intIndiceBouton = 3 To Worksheets(strPage1).Buttons.Count
            Worksheets(strPage1).Buttons(intIndiceBouton).Delete
        Next intIndiceBouton
    End If
    
    ' *** AJOUT BOUTON ***
    Set rngBtnCreerListeEleve = Worksheets(strPage1).Range("J10:J11")
    Set btnCreerListeEleve = Worksheets(strPage1).Buttons.Add(rngBtnCreerListeEleve.Left, rngBtnCreerListeEleve.Top, rngBtnCreerListeEleve.Width, rngBtnCreerListeEleve.Height)
    With btnCreerListeEleve
        .Caption = "Valider les données & créer les listes"
        .OnAction = "btnCreerListeEleve_Click"
    End With
End Sub

Sub btnCreerListeEleve_Click()
    ' *** DECLARATION VARIABLES ***
    Dim intNbDomaines As Integer
    Dim intNbCompetences As Integer
    Dim intIndiceDomaine As Integer
    Dim intNbClasses As Integer
    Dim intNbEleves As Integer
    Dim intIndiceClasse As Integer
    Dim intCompteur As Integer

    ' *** AFFECTATION VARIABLES ***
    intNbDomaines = getNombreDomaines
    intNbClasses = getNombreClasses
    intCompteur = 0
    
    ' *** PROTECTION + REFRESH ECRAN OFF ***
    Application.ScreenUpdating = False
    unprotectWorksheet
    
    ' *** VERIFICATION DONNEES COMPETENCES ***
    For intIndiceDomaine = 1 To intNbDomaines
        intNbCompetences = getNombreCompetences(intIndiceDomaine)
        If Not IsEmpty(intNbCompetences) And IsNumeric(intNbCompetences) Then
            If intNbCompetences > intNbMaxCompetences Then
                Worksheets(strPage1).Cells(12 + intIndiceDomaine, 3).Value = intNbMaxCompetences
            ElseIf intNbCompetences < intNbMinCompetences Then
                Worksheets(strPage1).Cells(12 + intIndiceDomaine, 3).Value = intNbMinCompetences
            End If
            intCompteur = intCompteur + 1
        End If
    Next intIndiceDomaine
    If intCompteur = intNbDomaines Then
        intCompteur = 0
    Else
        MsgBox ("Veuillez entrer des valeurs correctes pour les compétences de chaque domaine.")
        Exit Sub
    End If
    
    ' *** VERIFICATION DONNEES CLASSES ***
    For intIndiceClasse = 1 To intNbClasses
        intNbEleves = getNombreEleves(intIndiceClasse)
        If Not IsEmpty(intNbEleves) And IsNumeric(intNbEleves) Then
            If intNbEleves > intNbMaxEleves Then
                Worksheets(strPage1).Cells(12 + intIndiceClasse, 7).Value = intNbMaxEleves
            ElseIf intNbEleves < intNbMinEleves Then
                Worksheets(strPage1).Cells(12 + intIndiceClasse, 7).Value = intNbMinEleves
            End If
            intCompteur = intCompteur + 1
        End If
    Next intIndiceClasse
    If intCompteur <> intNbClasses Then
        MsgBox ("Veuillez entrer des valeurs correctes pour le nombre d'élèves de chaque classe.")
        Exit Sub
    End If
    
    ' *** DEMANDE CONFIRMATION ***
    If MsgBox("Êtes-vous sûr(e) de valider ces données ? Il ne sera pas possible de les modifier par la suite.", vbYesNo) = vbNo Then
        MsgBox ("Opération annulée")
        GoTo EOP
    End If
        
    ' *** VERROUILLAGE CELLULES + SUPPRESSION BOUTONS ***
    With Worksheets(strPage1)
        .Cells.Locked = True
        .Buttons.Delete
    End With
    
    ' *** CREATION LISTE ELEVE ***
    creerListeEleve
    
    ' *** MESSAGE INFORMATION ***
    MsgBox ("Listes de classe créées avec succès !")
    
    ' *** PROTECTION + REFRESH ECRAN ON ***
EOP:
    protectWorksheet
    Application.ScreenUpdating = True
End Sub


