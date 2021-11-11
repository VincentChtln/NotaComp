Attribute VB_Name = "Module1"

' *******************************************************************************
'   Copyright (C)
'   Date: 2021
'   Auteur: Vincent Chatelain
' *******************************************************************************
'                               NotaComp
'
'   Outil Excel dédié a la notation par compétences en milieu scolaire
'
'   Classeur vierge, documentation et fichiers source disponibles
'   sur le site <https://github.com/VincentChtln/NotaComp>
'
'   V1      Version initiale
'   V1.1    Amelioration de plusieurs fonctionnalités
'   V2      Refonte graphique et fonctionnelle
'   V2.4    Ajout de UserForm pour la modification des listes de classes
'   V2.5    Amélioration fonctionnelle du code, modifications graphiques
'   V2.6    Ajout de UserForm pour la modification des évaluations
'
' *******************************************************************************

'                       GNU General Public License V3
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
'                               Module 1 - Accueil
' *******************************************************************************
'
'   Fonctions publiques
'       IsWorkbookProtected() As Boolean
'       IsWorksheetProtected(ByVal ws As Worksheet) As Boolean
'       GetNomPage3(ByVal byClasse As Byte) As String
'       GetNomPage4(ByVal byClasse As Byte) As String
'       GetNomClasse(ByVal byClasse As Byte) As String
'       GetIndiceClasse(strNomWs As String) As Byte
'       GetNombreClasses() As Byte
'       GetNombreEleves(ByVal byClasse As Byte) As Byte
'       GetIndexInArray(ByRef arrSearch As Variant, ByVal varValue As Variant) As Long
'       GetSizeOfArray(ByRef arr As Variant) As Byte
'       GetSizeOfJaggedArray(ByRef arr As Variant) As Byte
'       GetArrayDomaines() As String()
'       GetArrayChoixDomaines() As Variant()
'       GetArrayChoixCompetences() As Variant()
'
'   Procédures publiques
'       ProtectWorkbook()
'       UnprotectWorkbook()
'       ProtectWorksheet(ByRef ws As Worksheet)
'       ProtectAllWorksheets()
'       UnprotectWorksheet(ByRef ws As Worksheet)
'       EnableUpdates()
'       DisableUpdates()
'       DisplayTemporaryMessage(ByVal strMessage As String, ByVal byDuration As Byte)
'       ClearStatusBar()
'       FreezePanes(ByRef wdw As Window, ByVal byLig As Byte, ByVal byCol As Byte)
'       UnloadAllUserForms()
'       DeleteAllButtons(ByRef ws As Worksheet)
'       AddWorksheet(ByVal sNom As String)
'       SetNombreEleves(ByVal byClasse As Byte, ByVal byNombre As Byte)
'       CreerTableau(ByVal strNomWs As String, ByVal rngCelOrigine As Range, _
'                    ByVal iHaut As Byte, ByVal iLarg As Byte, ByVal iOrientation As Byte, _
'                    ByRef arrAttribut() As String, ByVal byCouleur As Byte, Optional ByVal bLocked As Boolean = True)
'
'   Fonctions privées
'       IsInfosOK() As Boolean
'       IsNbClassesOK() As Boolean
'       IsNbEleveOK(ByVal varNbEleve As Variant) As Boolean
'       IsDonneesClassesOK() As Boolean
'       GetArrayCompetences(ByVal iCycle As Byte) As Variant()
'       GetNombreCompetences(ByVal iCycle As Byte) As Byte
'       IsCompetOK() As Boolean
'
'   Procédures privées
'       InitNotaComp()
'       BtnDemarrerConfiguration_Click()
'       BtnValiderNombreClasses_Click()
'       BtnModifierNombreClasses_Click()
'       BtnValiderClasses_Click()
'       DrpChoixCycle_Change()
'       BtnValiderCompetences_Click()
'       CreerTableauLogiciel()
'       CreerTableauInformations()
'       CreerTableauNombreClasses()
'       CreerTableauClasses(ByVal byNbClasses As Byte)
'       CreerDropdownCycle()
'       CreerTableauChoixCompetences(ByVal iDrpValue As Byte)
'
' *******************************************************************************

Option Explicit

' *******************************************************************************
'                                   Constantes
' *******************************************************************************

' *** LOGICIEL ***
Public Const strVersion         As String = "v2.5 - Révision 2021"
       Const strLienGithub      As String = "https://github.com/VincentChtln/NotaComp"
       Const strLienSocleCommun As String = "https://www.education.gouv.fr/bo/15/Hebdo17/MENE1506516D.htm?cid_bo=87834"

' *** DONNEES PROTECTION ***
Public Const strPage1           As String = "Accueil"
Public Const strPage2           As String = "Listes"
Public Const strPassword        As String = vbNullString

' *** COULEURS ***
       Const byCouleurLogiciel  As Byte = 22
       Const byCouleurInfos     As Byte = 15
Public Const byCouleurCompet_1  As Byte = 50
Public Const byCouleurCompet_2  As Byte = 35
Public Const byCouleurClasse    As Byte = 44
Public Const byCouleurEval_1    As Byte = 45
Public Const byCouleurEval_2    As Byte = 40
Public Const byCouleurNote_1    As Byte = 42
Public Const byCouleurNote_2    As Byte = 34
Public Const byCouleurBilan     As Byte = 45

' *** LIMITES MIN ET MAX ***
       Const byNbClasses_Min    As Byte = 1
       Const byNbClasses_Max    As Byte = 20
       Const byNbEleves_Min     As Byte = 5
       Const byNbEleves_Max     As Byte = 50

' *** LIGNES ET COLONNES REFERENCE ***
       Const byLigTabLogiciel   As Byte = 5
       Const byColTabLogiciel   As Byte = 2
       Const byLigTabInfos      As Byte = byLigTabLogiciel + 5
       Const byColTabInfos      As Byte = 2
Public Const byLigTabClasses    As Byte = byLigTabInfos + 6
Public Const byColTabClasses    As Byte = 2
Public Const byColTabCompet     As Byte = 2
Public Const byLigListePage2    As Byte = 1
Public Const byLigListePage3    As Byte = 6
Public Const byLigListePage4    As Byte = 3

' *** NOTES REFERENCE ***
Public Const dblNoteA_Min       As Double = 3.4
Public Const dblNoteB_Min       As Double = 2.4
Public Const dblNoteC_Min       As Double = 1.4
Public Const dblNoteD_Min       As Double = 0

' *******************************************************************************
'                               Fonctions publiques
' *******************************************************************************

Public Function IsWorkbookProtected() As Boolean
    IsWorkbookProtected = ThisWorkbook.ProtectWindows Or ThisWorkbook.ProtectStructure
End Function

Public Function IsWorksheetProtected(ByVal ws As Worksheet) As Boolean
    IsWorksheetProtected = ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios
End Function

Public Function GetNomPage3(ByVal byClasse As Byte) As String
    GetNomPage3 = "Notes (" & GetNomClasse(byClasse) & ")"
End Function

Public Function GetNomPage4(ByVal byClasse As Byte) As String
    GetNomPage4 = "Bilan (" & GetNomClasse(byClasse) & ")"
End Function

Public Function GetNomClasse(ByVal byClasse As Byte) As String
    GetNomClasse = ThisWorkbook.Worksheets(strPage1).Cells(byLigTabClasses + byClasse + 2, byColTabClasses).Value
End Function

Public Function GetIndiceClasse(strNomWs As String) As Byte
    Dim byClasse As Byte
    
    GetIndiceClasse = 0
    
    For byClasse = 1 To GetNombreClasses
        If InStr(strNomWs, GetNomClasse(byClasse)) <> 0 Then GetIndiceClasse = byClasse
    Next byClasse
End Function

Public Function GetNombreClasses() As Byte
    Dim byNbClasses As Byte
    byNbClasses = ThisWorkbook.Worksheets(strPage1).Cells(byLigTabClasses, byColTabClasses + 1).Value
    
    GetNombreClasses = 0
    If IsNumeric(byNbClasses) Then
        If (byNbClasses >= byNbClasses_Min) And (byNbClasses <= byNbClasses_Max) Then
            GetNombreClasses = CByte(byNbClasses)
        End If
    End If
End Function

Public Function GetNombreEleves(ByVal byClasse As Byte) As Byte
    GetNombreEleves = CByte(ThisWorkbook.Worksheets(strPage1).Cells(byLigTabClasses + byClasse + 2, byColTabClasses + 1).Value)
End Function

Public Function GetSizeOfArray(ByRef arr As Variant) As Byte
    Dim byTaille        As Byte
    Dim byDimension     As Byte
    byTaille = 1
    byDimension = 1
    GetSizeOfArray = 0
    
    If IsArray(arr) Then
        On Error GoTo ErrorHandler
        Do While True
            byTaille = byTaille * (UBound(arr, byDimension) - LBound(arr, byDimension) + 1)
            byDimension = byDimension + 1
            GetSizeOfArray = byTaille
        Loop
    Else
        GetSizeOfArray = 0
        Exit Function
    End If
    
ErrorHandler:
    If Err.Number = 13 Then ' Type Mismatch Error
        Err.Raise vbObjectError, "GetSizeOfArray" _
            , "The argument passed to the GetSizeOfArray function is not an array."
    End If
End Function

Public Function GetSizeOfJaggedArray(ByRef arr As Variant) As Byte
    Dim iTailleExter As Byte
    Dim iElement As Byte
    
    iTailleExter = GetSizeOfArray(arr)
    GetSizeOfJaggedArray = 0
    
    iTailleExter = GetSizeOfArray(arr)
    For iElement = 1 To iTailleExter
        GetSizeOfJaggedArray = GetSizeOfJaggedArray + GetSizeOfArray(arr(iElement))
    Next iElement
End Function

Public Function GetArrayDomaines() As String()
    Dim arrDomaines(1 To 8, 1 To 2) As String
    
    ' *** NOM COMPLET ***
    arrDomaines(1, 1) = "Domaine 1: Les langages pour penser et communiquer - " & vbNewLine & _
                        "Composante 1: La langue française"
    arrDomaines(2, 1) = "Domaine 1: Les langages pour penser et communiquer - " & vbNewLine & _
                        "Composante 2: Les langues vivantes étrangères ou régionales"
    arrDomaines(3, 1) = "Domaine 1: Les langages pour penser et communiquer - " & vbNewLine & _
                        "Composante 3: Les langages mathématiques, scientifiques et informatiques"
    arrDomaines(4, 1) = "Domaine 1: Les langages pour penser et communiquer - " & vbNewLine & _
                        "Composante 4: Les langages des arts et du corps"
    arrDomaines(5, 1) = "Domaine 2: Les méthodes et outils pour apprendre"
    arrDomaines(6, 1) = "Domaine 3: La formation de la personne et du citoyen"
    arrDomaines(7, 1) = "Domaine 4: les systèmes naturels et les systèmes techniques"
    arrDomaines(8, 1) = "Domaine 5: Les représentations du monde et de l'activité humaine"
    
    ' *** ABREVIATION ***
    arrDomaines(1, 2) = "D1-1"
    arrDomaines(2, 2) = "D1-2"
    arrDomaines(3, 2) = "D1-3"
    arrDomaines(4, 2) = "D1-4"
    arrDomaines(5, 2) = "D2"
    arrDomaines(6, 2) = "D3"
    arrDomaines(7, 2) = "D4"
    arrDomaines(8, 2) = "D5"
    
    GetArrayDomaines = arrDomaines
End Function

Public Function GetArrayChoixCompetences() As Variant()
    ' *** DECLARATION VARIABLES ***
    Dim arrTamponSrc As Variant
    Dim arrTamponDest() As Variant
    Dim arrCompet() As Variant
    Dim arrChoixCompet(1 To 8) As Variant
    Dim byNbCompetParDomaine As Byte
    Dim byLigTabCompet As Byte
    Dim iCompetTampon As Byte
    Dim iCompetChoisie As Byte
    Dim iDrpValue As Byte
    Dim iDomaine As Byte
    Dim iCompet As Byte

    With ThisWorkbook.Worksheets(strPage1)
        ' *** AFFECTATION VARIABLES ***
        byLigTabCompet = byLigTabClasses + GetNombreClasses + 7
        iDrpValue = .DropDowns("drpChoixCycle").Value
        iCompetTampon = 1
        arrCompet = GetArrayCompetences(iDrpValue + 1)
        arrTamponSrc = .Range(.Cells(byLigTabCompet + 1, byColTabCompet + 1), _
                              .Cells(byLigTabCompet + GetNombreCompetences(iDrpValue + 1), byColTabCompet + 3))
                              
        ' *** BOUCLE SUR TOUS LES DOMAINES ET COMPETENCES ***
        For iDomaine = 1 To 8
            byNbCompetParDomaine = GetSizeOfArray(arrCompet(iDomaine))
            iCompetChoisie = 1
            ReDim arrTamponDest(1 To 2, 1 To 1)
            For iCompet = 1 To byNbCompetParDomaine
            
                ' *** SI 'x' OU 'X' DANS LE TABLEAU, ALORS COPIE ABREVIATION COMPETENCE ***
                If arrTamponSrc(iCompetTampon, 2) = "X" Or arrTamponSrc(iCompetTampon, 2) = "x" Then
                    ReDim Preserve arrTamponDest(1 To 2, 1 To iCompetChoisie)
                    arrTamponDest(1, iCompetChoisie) = arrTamponSrc(iCompetTampon, 3)
                    arrTamponDest(2, iCompetChoisie) = Right(arrTamponSrc(iCompetTampon, 1), Len(arrTamponSrc(iCompetTampon, 1)) - 5)
                    iCompetChoisie = iCompetChoisie + 1
                End If
                iCompetTampon = iCompetTampon + 1
            Next iCompet
            If Not IsEmpty(arrTamponDest(1, 1)) Then
                arrChoixCompet(iDomaine) = Application.WorksheetFunction.Transpose(arrTamponDest)
            End If
        Next iDomaine
    End With
    GetArrayChoixCompetences = arrChoixCompet
End Function

Public Function GetArrayChoixDomaines() As Variant()
    ' *** DECLARATION VARIABLES ***
    Dim byDomaine As Byte
    Dim byDomaine2 As Byte
    Dim arrChoixCompet As Variant
    Dim arrDomaines As Variant
    Dim arrChoixDomaines As Variant
    
    ' *** AFFECTATION VARIABLES ***
    byDomaine2 = 1
    ReDim arrChoixDomaines(1 To 1)
    arrChoixCompet = GetArrayChoixCompetences()
    arrDomaines = GetArrayDomaines()
    
    ' *** CALCUL ARRAY ***
    For byDomaine = 1 To 8
        If GetSizeOfArray(arrChoixCompet(byDomaine)) <> 0 Then
            If byDomaine2 > 1 Then ReDim Preserve arrChoixDomaines(1 To byDomaine2)
            arrChoixDomaines(byDomaine2) = arrDomaines(byDomaine, 2)
            byDomaine2 = byDomaine2 + 1
        End If
    Next byDomaine
    
    ' *** RENVOI ARRAY ***
    GetArrayChoixDomaines = arrChoixDomaines
End Function

Public Function GetIndexInArray(ByRef arrSearch As Variant, ByVal varValue As Variant) As Long
    Dim lIndex As Long
    
    For lIndex = LBound(arrSearch, 1) To UBound(arrSearch, 1)
        If arrSearch(lIndex) = varValue Then GoTo EOF
    Next lIndex
    lIndex = 0
EOF:
    GetIndexInArray = lIndex
End Function


' *******************************************************************************
'                               Procédures publiques
' *******************************************************************************

Public Sub ProtectWorkbook()
    ThisWorkbook.Protect Password:=strPassword, Structure:=True, Windows:=True
End Sub

Public Sub UnprotectWorkbook()
    ThisWorkbook.Unprotect Password:=strPassword
End Sub

Public Sub ProtectWorksheet(ByRef ws As Worksheet)
    ws.Protect Password:=strPassword, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterFaceOnly:=True
    ws.EnableSelection = xlUnlockedCells
End Sub

Public Sub ProtectAllWorksheets()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        ProtectWorksheet ws
    Next ws
End Sub

Public Sub UnprotectWorksheet(ByRef ws As Worksheet)
    ws.Unprotect Password:=strPassword
End Sub

Public Sub EnableUpdates()
    With Application
        .ScreenUpdating = True
        .StatusBar = "Prêt"
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .DisplayAlerts = True
        .PrintCommunication = True
    End With
End Sub

Public Sub DisableUpdates()
    With Application
        .ScreenUpdating = False
        .StatusBar = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .DisplayAlerts = False
        .PrintCommunication = False
    End With
End Sub

Public Sub DisplayTemporaryMessage(ByVal strMessage As String, Optional ByVal byDuration As Byte = 10)
    Application.StatusBar = strMessage
    Application.OnTime Now + TimeSerial(0, 0, byDuration), "ClearStatusBar"
End Sub

Public Sub ClearStatusBar()
    Application.StatusBar = vbNullString
End Sub

Public Sub FreezePanes(ByRef wdw As Window, ByVal byLig As Byte, ByVal byCol As Byte)
    With wdw
        .SplitRow = byLig
        .SplitColumn = byCol
        .FreezePanes = True
    End With
End Sub

Public Sub UnloadAllUserForms()
    Dim uf As UserForm
    
    For Each uf In VBA.UserForms
        If TypeOf uf Is UserForm Then uf.Hide
    Next uf
    
    For Each uf In VBA.UserForms
        If TypeOf uf Is UserForm Then Unload uf
    Next uf
End Sub

Public Sub DeleteAllButtons(ByRef ws As Worksheet)
    ws.Buttons.Delete
End Sub

Public Sub AddWorksheet(ByVal sNom As String)
    If IsWorkbookProtected Then Exit Sub
    
    With ThisWorkbook
        .Worksheets.Add After:=.Worksheets(.Worksheets.Count)
        .Worksheets(.Worksheets.Count).Name = sNom
    
        With .Worksheets(sNom)
            With .Cells
                .Borders.ColorIndex = 2
                .Locked = True
                .HorizontalAlignment = xlHAlignCenter
                .VerticalAlignment = xlVAlignCenter
            End With
            
            If InStr(sNom, "Notes") Or InStr(sNom, "Bilan") Then
                .Columns.ColumnWidth = 10
            Else
                .Columns.ColumnWidth = 20
            End If
        End With
    End With
End Sub

Public Sub SetNombreEleves(ByVal byClasse As Byte, ByVal byNombre As Byte)
    ThisWorkbook.Worksheets(strPage1).Cells(byLigTabClasses + byClasse + 2, byColTabClasses + 1).Value = byNombre
End Sub

Public Sub CreerTableau(ByVal strNomWs As String, ByVal rngCelOrigine As Range, _
                        ByVal iHaut As Byte, ByVal iLarg As Byte, ByVal iOrientation As Byte, _
                        ByRef arrAttribut() As String, ByVal byCouleur As Byte, Optional ByVal bLocked As Boolean = True)
    ' ***  DECLARATION VARIABLES ***
    Dim ws As Worksheet
    Dim bWsNomOK As Boolean
    Dim iAttribut As Byte
    
    ' *** VERIFICATIONS DONNEES ENTREES ***
    bWsNomOK = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = strNomWs Then
            bWsNomOK = True
            Exit For
        End If
    Next ws
    If Not (bWsNomOK) Or Not (rngCelOrigine.Count = 1) Or Not (iHaut >= 1) Or Not (iLarg >= 1) Then Exit Sub
    If Not (iOrientation = 1) And Not (iOrientation = 2) Then Exit Sub
    If (iOrientation = 1) And (Not (iHaut >= 2) Or Not (GetSizeOfArray(arrAttribut) = iLarg)) Then Exit Sub
    If (iOrientation = 2) And (Not (iLarg >= 2) Or Not (GetSizeOfArray(arrAttribut) = iHaut)) Then Exit Sub
    If Not ((byCouleur >= 1) And (byCouleur <= 56)) Then Exit Sub
    
    ' *** CREATION TABLEAU ***
    With ThisWorkbook.Worksheets(strNomWs)
        With .Range(rngCelOrigine.Address, .Cells(rngCelOrigine.Row + iHaut - 1, rngCelOrigine.Column + iLarg - 1))
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.Color = 1
            .VerticalAlignment = xlVAlignCenter
            .HorizontalAlignment = xlHAlignCenter
        End With
        Select Case iOrientation
            ' *** TAB VERTICAL --> ATTRIBUTS EN HAUT ***
        Case 1
            .Range(rngCelOrigine.Address, .Cells(rngCelOrigine.Row, rngCelOrigine.Column + iLarg - 1)).Interior.ColorIndex = byCouleur
            For iAttribut = LBound(arrAttribut) To UBound(arrAttribut)
                .Cells(rngCelOrigine.Row, rngCelOrigine.Column + iAttribut - LBound(arrAttribut)).Value = arrAttribut(iAttribut)
            Next iAttribut
            .Range(.Cells(rngCelOrigine.Row + 1, rngCelOrigine.Column), _
                   .Cells(rngCelOrigine.Row + iHaut - 1, rngCelOrigine.Column + iLarg - 1)).Locked = bLocked
            
            ' *** TAB HORIZONTAL --> ATTRIBUTS SUR LE COTE ***
        Case 2
            .Range(rngCelOrigine.Address, .Cells(rngCelOrigine.Row + iHaut - 1, rngCelOrigine.Column)).Interior.ColorIndex = byCouleur
            For iAttribut = LBound(arrAttribut) To UBound(arrAttribut)
                .Cells(rngCelOrigine.Row + iAttribut - LBound(arrAttribut), rngCelOrigine.Column).Value = arrAttribut(iAttribut)
            Next iAttribut
            .Range(.Cells(rngCelOrigine.Row, rngCelOrigine.Column + 1), _
                   .Cells(rngCelOrigine.Row + iHaut - 1, rngCelOrigine.Column + iLarg - 1)).Locked = bLocked
        End Select
    End With
End Sub

' *******************************************************************************
'                               Fonctions privées
' *******************************************************************************
Private Function IsInfosOK() As Boolean
    IsInfosOK = False
    With ThisWorkbook.Worksheets(strPage1)
        If WorksheetFunction.CountBlank(.Range(.Cells(byLigTabInfos, byColTabInfos + 1), .Cells(byLigTabInfos + 3, byColTabInfos + 1))) = 0 Then IsInfosOK = True
    End With
End Function

Private Function IsNbClassesOK() As Boolean
    IsNbClassesOK = False
    If GetNombreClasses <> 0 Then IsNbClassesOK = True
End Function

Private Function IsNbEleveOK(ByVal varNbEleve As Variant) As Boolean
    IsNbEleveOK = False
    If IsNumeric(varNbEleve) Then
        If varNbEleve > byNbEleves_Min And varNbEleve < byNbEleves_Max Then IsNbEleveOK = True
    End If
End Function

Private Function IsDonneesClassesOK() As Boolean
    ' *** DECLARATION VARIABLES ***
    Dim byNbClasses As Byte
    Dim byLigFinTableauClasses As Byte
    Dim rngTableauClasses As Range
    Dim celDonnee As Range
    
    IsDonneesClassesOK = False
    byNbClasses = GetNombreClasses
    If byNbClasses = -1 Then GoTo EOF
    
    With ThisWorkbook.Worksheets(strPage1)
        ' *** VERIFICATION NOMBRE CLASSES ***
        byLigFinTableauClasses = .Range(Mid(.UsedRange.Address, InStr(1, .UsedRange.Address, ":") + 1)).Row
        If byLigTabClasses + byNbClasses + 2 <> byLigFinTableauClasses Then GoTo EOF
        
        ' *** VERIFICATION VALEUR MANQUANTE ***
        Set rngTableauClasses = .Range(.Cells(byLigTabClasses + 3, byColTabClasses), .Cells(byLigFinTableauClasses, byColTabClasses + 1))
        If WorksheetFunction.CountBlank(rngTableauClasses) <> 0 Then GoTo EOF
        
        ' *** VERIFICATION NOMBRE ELEVES ***
        For Each celDonnee In rngTableauClasses
            If celDonnee.Column = byColTabClasses + 1 Then
                If Not (IsNbEleveOK(celDonnee.Value)) Then GoTo EOF
            End If
        Next celDonnee
    End With
    
    IsDonneesClassesOK = True
    
EOF:
End Function

Private Function GetArrayCompetences(ByVal iCycle As Byte) As Variant()
    Dim arrCompetencesCycleI() As Variant
    
    ReDim arrCompetencesCycleI(1 To 8)
    Select Case iCycle
    
        ' *** COMPETENCES CYCLE 2 ***
    Case 2
        arrCompetencesCycleI(1) = Array("C1 - Comprendre et s'exprimer à l'oral", _
                                        "C2 - Lire et comprendre l'écrit", _
                                        "C3 - Ecrire", _
                                        "C4 - Utiliser à bon escient les régularités qui organisent la langue française " & _
                                        "(dans la limite de celles qui ont été étudiées)")
        arrCompetencesCycleI(2) = Array("C1 - Comprendre à l'oral (et à l'écrit)", _
                                        "C2 - S'exprimer à l'oral")
        arrCompetencesCycleI(3) = Array("C1 - Utiliser les nombres entiers", _
                                        "C2 - Reconnaître des solides usuels et des figures géométriques", _
                                        "C3 - Se repérer et se déplacer")
        arrCompetencesCycleI(4) = Array("C1 - S'exprimer par des activités physiques, sportives ou artistiques, impliquant le corps", _
                                        "C2 - Partager et comprendre les langages artistiques")
        arrCompetencesCycleI(5) = Array("C1 - Organiser son travail personnel", _
                                        "C2 - Coopérer avec des pairs", _
                                        "C3 - Rechercher et trairer l'information au moyen d'outils numériques")
        arrCompetencesCycleI(6) = Array("C1 - S'exprimer (émotions, opinions, préférences) et respecter l'expression d'autrui", _
                                        "C2 - Prendre en compte les règles communes", _
                                        "C3 - Manifester son appartenance à un collectif")
        arrCompetencesCycleI(7) = Array("C1 - Résoudre des problèmes élémentaires", _
                                        "C2 - Mener quelques étapes d'une démarche scientifique" & _
                                        "C3 - Mettre en pratique des comportements simples respectueux des autres, de l'environnement, de sa santé")
        arrCompetencesCycleI(8) = Array("C1 - Se situer dans le temps et l'espace", _
                                        "C2 - Analyser et comprendre des organisations humaines et les représentations du monde", _
                                        "C3 - Imaginer, élaborer, produire")
        
        ' *** COMPETENCES CYCLE 3 ***
    Case 3
        arrCompetencesCycleI(1) = Array("C1 - S'exprimer à l'oral", _
                                        "C2 - Comprendre des énoncés oraux", _
                                        "C3 - Lire et comprendre l'écrit", _
                                        "C4 - Ecrire", _
                                        "C5 - Exploiter les ressources de la langues / Réfléchir sur le système linguistique")
        arrCompetencesCycleI(2) = Array("C1 - Lire et comprendre l'écrit", _
                                        "C2 - Ecrire et réagir à l'écrit", _
                                        "C3 - Ecouter et comprendre", _
                                        "C4 - S'exprimer à l'oral en continu et en interaction")
        arrCompetencesCycleI(3) = Array("C1 - Utiliser les nombres entiers, les nombres décimaux et les fractions simples", _
                                        "C2 - Reconnaître des solides usuels et des figures géométriques", _
                                        "C3 - Se repérer et se déplacer")
        arrCompetencesCycleI(4) = Array("C1 - S'exprimer par des activités physiques sportives ou artistiques", _
                                        "C2 - Pratiquer des arts en mobilisant divers langages artistiques et leurs ressources expressives / " & _
                                        "Prendre du recul sur la pratique artistique individuelle et collective")
        arrCompetencesCycleI(5) = Array("C1 - Se constituter des outils de travail personnel et mettre en place des stratégies pour comprendre et apprendre", _
                                        "C2 - Coopérer et réaliser des projets", _
                                        "C3 - Rechercher et trairer l'information et s'initier aux langages des médias", _
                                        "C4 - Mobiliser des outils numériques pour apprendre, échanger et communiquer")
        arrCompetencesCycleI(6) = Array("C1 - Maîtriser l'expression de sa sensibilité et de ses opinions, respecter celles des autres", _
                                        "C2 - Comprendre et connaître la règle et le droit", _
                                        "C3 - Exercer son esprit critique, faire preuve de réflexion et de discernement")
        arrCompetencesCycleI(7) = Array("C1 - Mener une démarche scientifique ou technologique, résoudre des problèmes simples", _
                                        "C2 - Mettre en pratique des comportements simples respectueux des autres, de l'environnement, de sa santé")
        arrCompetencesCycleI(8) = Array("C1 - Se situer dans le temps et l'espace", _
                                        "C2 - Analyser et comprendre des organisations humaines et les représentations du monde", _
                                        "C3 - Raisonner, imaginer, élaborer, produire")

        ' *** COMPETENCES CYCLE 4 ***
    Case 4
        arrCompetencesCycleI(1) = Array("C1 - S'exprimer à l'oral", _
                                        "C2 - Comprendre des énoncés oraux", _
                                        "C3 - Lire et comprendre l'écrit", _
                                        "C4 - Ecrire", _
                                        "C5 - Exploiter les ressources de la langues / Réfléchir sur le système linguistique")
        arrCompetencesCycleI(2) = Array("C1 - Lire et comprendre l'écrit", _
                                        "C2 - Ecrire et réagir à l'écrit", _
                                        "C3 - Ecouter et comprendre", _
                                        "C4 - S'exprimer à l'oral en continu et en interaction")
        arrCompetencesCycleI(3) = Array("C1 - Utiliser les nombres", _
                                        "C2 - Utiliser un calcul littéral", _
                                        "C3 - Exprimer une grandeur mesurée ou calculée dans une unité adaptée", _
                                        "C4 - Passer d'un langage à un autre", _
                                        "C5 - Utiliser le langage des probabilités", _
                                        "C6 - Utiliser et produire des représentations d'objets", _
                                        "C7 - Utiliser l'algorithmique et la programmation pour créer des applications simples")
        arrCompetencesCycleI(4) = Array("C1 - Pratiquer des activités physiques sportives et     artistiques", _
                                        "C2 - Pratiquer des arts en mobilisant divers langages artistiques et leurs ressources expressives / " & _
                                        "Prendre du recul sur la pratique artistique individuelle et collective")
        arrCompetencesCycleI(5) = Array("C1 - Organiser son travail personnel", _
                                        "C2 - Coopérer et réaliser des projets", _
                                        "C3 - Rechercher et trairer l'information et s'initier aux langages des médias", _
                                        "C4 - Mobiliser des outils numériques pour apprendre, échanger et communiquer")
        arrCompetencesCycleI(6) = Array("C1 - Maîtriser l'expression de sa sensibilité et de ses opinions, respecter celles des autres", _
                                        "C2 - Comprendre et connaître la règle et le droit", _
                                        "C3 - Exercer son esprit critique, faire preuve de réflexion et de discernement", _
                                        "C4 - Faire preuve de responsabilité, respecter les règles de vie collective, s'engager et prendre des initiatives")
        arrCompetencesCycleI(7) = Array("C1 - Mener une démarche scientifique", _
                                        "C2 - Concevoir des objets et systèmes techniques", _
                                        "C3 - Identifier des règles et des principes de responsabilité individuelle et collective " & _
                                        "dans le domaine de la santé, de la sécurité, de l'environnement")
        arrCompetencesCycleI(8) = Array("C1 - Se situer dans le temps et l'espace", _
                                        "C2 - Analyser et comprendre des organisations humaines et les représentations du monde", _
                                        "C3 - Raisonner, imaginer, élaborer, produire")
        
        ' *** VALEUR iCycle NON VALIDE
    Case Else
        ReDim arrCompetencesCycleI(1 To 1)
        arrCompetencesCycleI = Array(vbNullString)
    End Select
    
    GetArrayCompetences = arrCompetencesCycleI

End Function

Private Function GetNombreCompetences(ByVal iCycle As Byte) As Byte
    GetNombreCompetences = GetSizeOfJaggedArray(GetArrayCompetences(iCycle))
End Function

Private Function IsCompetOK() As Boolean
    Dim arrChoixCompet As Variant
    Dim byDomaine As Byte
    Dim byCompet As Byte
    Dim strCompet As Variant
    Dim byNbCompetParDomaine As Byte
    
    arrChoixCompet = GetArrayChoixCompetences
    IsCompetOK = True
    
    ' *** VERIF 2 COMPET MIN ***
    If GetSizeOfJaggedArray(arrChoixCompet) / 2 < 3 Then
        IsCompetOK = False
        Exit Function
    End If
    
    ' *** VERIF NOM COMPET VALIDE ***
    For byDomaine = 1 To 8
        byNbCompetParDomaine = GetSizeOfArray(arrChoixCompet(byDomaine)) / 2
        ' *** 1 COMPET -> SUPPRESSION D'UNE DIMENSION DU SOUS-ARRAY ***
        If byNbCompetParDomaine = 1 Then
            If Len(CStr(arrChoixCompet(byDomaine)(1))) > 7 Then
                IsCompetOK = False
            End If
        ' *** PLUSIEURS COMPET -> JAGGED ARRAY CORRECT ***
        ElseIf byNbCompetParDomaine > 1 Then
            For byCompet = 1 To UBound(arrChoixCompet(byDomaine), 1)
                If Len(CStr(arrChoixCompet(byDomaine)(byCompet, 1))) > 7 Then
                    IsCompetOK = False
                End If
            Next byCompet
        End If
    Next byDomaine
End Function

' *******************************************************************************
'                               Procédures privées
' *******************************************************************************

'@EntryPoint
Private Sub InitNotaComp()
    ' *** DECLARATION VARIABLES ***
    Dim rngBtnDemarrerNotacomp As Range
    Dim shp As Shape
    
    ' *** REFRESH ECRAN OFF ***
    DisableUpdates

    ' *** VERIFICATION PROTECTION WORKBOOK ***
    If IsWorkbookProtected Then
        MsgBox "Ce classeur est protégé. Enlevez la protection avant de continuer."
        GoTo EOP
    End If
        
    With ThisWorkbook
        ' *** VERIFICATION WORKSHEET UNIQUE ***
        If .Worksheets.Count <> 1 Then
            MsgBox "Ce classeur contient plusieurs feuilles. Supprimez toutes les feuilles sauf une avant de continuer."
            GoTo EOP
        End If
        
        ' *** VERIFICATION PROTECTION FEUILLE ***
        If IsWorksheetProtected(.Worksheets(1)) Then
            MsgBox "Cette feuille est protégée. Enlevez la protection avant de continuer."
            GoTo EOP
        End If
            
        With .Worksheets(1)
            ' *** VERIFICATION FEUILLE VIDE ***
            If Not (.UsedRange.Address = "$A$1" And .Range("A1") = vbNullString) Then
                Select Case MsgBox("La feuille n'est pas vide, cliquez sur 'OK' pour supprimer son contenu" & vbNewLine & _
                                   "ou sur 'Annuler' pour revenir en arrière", vbOKCancel)
                Case vbOK
                    .Cells.Clear
                    .EnableSelection = xlUnlockedCells
                    For Each shp In .Shapes
                        shp.Delete
                    Next shp
                Case vbCancel
                    GoTo EOP
                End Select
            End If
            
            ' *** MISE EN PAGE ***
            .Name = strPage1
            .Rows().RowHeight = 20
            .Rows(2).RowHeight = 50
            .Columns().ColumnWidth = 30
            .Columns(1).ColumnWidth = 5
            .Columns(2).ColumnWidth = 50
            .Columns(3).ColumnWidth = 50
            .Columns(4).ColumnWidth = 20
            With .Cells
                .VerticalAlignment = xlVAlignCenter
                .HorizontalAlignment = xlHAlignCenter
                .Borders.ColorIndex = 2
                .Locked = True
            End With
            With .Range("B2")
                .HorizontalAlignment = xlHAlignLeft
                .Font.Size = 40
                .Value = "NotaComp"
            End With
            
            ' *** AJOUT BOUTON 'DEMARRER NOTACOMP' ***
            Set rngBtnDemarrerNotacomp = .Range(.Cells(byLigTabLogiciel, byColTabLogiciel + 3).Address)
            With .Buttons.Add(rngBtnDemarrerNotacomp.Left, rngBtnDemarrerNotacomp.Top, _
                              rngBtnDemarrerNotacomp.Width, rngBtnDemarrerNotacomp.Height)
                .Caption = "Démarrer NotaComp"
                .OnAction = "BtnDemarrerConfiguration_Click"
                .Name = "BtnDemarrerConfiguration"
            End With

            ' *** AJOUT TABLEAU INFOS LOGICIEL ***
            CreerTableauLogiciel
        End With
        ' *** PROTECTION + REFRESH ECRAN ON ***
        ProtectWorksheet .Worksheets(strPage1)
        ProtectWorkbook
    End With
    
EOP:
    EnableUpdates
End Sub

Private Sub CreerTableauLogiciel()
    Dim arrAttributLogiciel(1 To 3) As String
    
    arrAttributLogiciel(1) = "Version de l'outil"
    arrAttributLogiciel(2) = "Classeur vierge, documents et code source"
    arrAttributLogiciel(3) = "Textes officiels - Socle commun"
    
    With ThisWorkbook.Worksheets(strPage1)
        CreerTableau strNomWs:=strPage1, rngCelOrigine:=.Cells(byLigTabLogiciel, byColTabLogiciel), _
        iHaut:=3, iLarg:=2, iOrientation:=2, _
        arrAttribut:=arrAttributLogiciel, byCouleur:=byCouleurLogiciel
        .Cells(byLigTabLogiciel, byColTabLogiciel + 1).Value = strVersion
        .Hyperlinks.Add Anchor:=.Cells(byLigTabLogiciel + 1, byColTabLogiciel + 1), _
        Address:=strLienGithub, TextToDisplay:="Dépôt Github"
        .Hyperlinks.Add Anchor:=.Cells(byLigTabLogiciel + 2, byColTabLogiciel + 1), _
        Address:=strLienSocleCommun, TextToDisplay:="Bulletin Officiel - Décret n° 2015-372 du 31/03/2015"
        .Range(.Cells(byLigTabLogiciel + 1, byColTabLogiciel + 1), .Cells(byLigTabLogiciel + 2, byColTabLogiciel + 1)).Locked = False
    End With
End Sub

'@EntryPoint
Private Sub BtnDemarrerConfiguration_Click()
    ' *** REFRESH ECRAN OFF ***
    DisableUpdates
    
    ' *** MESSAGE INFORMATION ***
    MsgBox "Bienvenue dans la configuration de NotaComp. Prennez le temps nécessaire pour réaliser correctement " & _
           "cette démarche, car il ne sera pas possible de revenir dessus plus tard. " & _
           "Le processus est découpé en plusieurs étapes, et des messages comme celui-ci " & _
           "s'afficheront pour vous donner des indications. Merci de les lire attentivement." & vbNewLine & vbNewLine & _
           "Au fur et à mesure, des informations vous seront demandées concernant:" & vbNewLine & _
           "    - Les classes auxquelles vous enseignez, avec la liste des élèves" & vbNewLine & _
           "    - Les compétences évaluées au cours de l'année (selon la dénomination officielle)" & vbNewLine & vbNewLine & _
           "Pensez à préparer ces éléments au préalable pour faciliter la configuration." & vbNewLine & vbNewLine & _
           "C'est parti !"
    ' *** SUPPRESSION BOUTONS ***
    DeleteAllButtons ThisWorkbook.Worksheets(strPage1)
    
    ' *** APPEL PROCEDURE ***
    CreerTableauInformations
    CreerTableauNombreClasses

    ' *** REFRESH ECRAN ON ***
    EnableUpdates
        
    ' *** MESSAGE INFORMATION ***
    MsgBox "Entrez tout d'abord vos informations dans le tableau d'informations (tableau gris)." & vbNewLine & _
           "Puis entrez le nombre de classes dans la case correspondante (tableau jaune). " & _
           "Cliquez ensuite sur le bouton 'Valider le nombre de classes' pour passer à l'étape suivante." & vbNewLine & vbNewLine & _
           "ATTENTION: ce classeur peut comporter des classes de différents niveaux, mais appartenant toutes au même cycle (2, 3 ou 4). " & _
           "Par exemple, si vous enseignez en collège à des classes de Sixième et de Quatrième, il vous faudra faire deux classeurs séparés."

End Sub

Private Sub CreerTableauInformations()
    Dim arrAttributInfos(1 To 4) As String
    
    arrAttributInfos(1) = "Etablissement"
    arrAttributInfos(2) = "Matière"
    arrAttributInfos(3) = "Professeur"
    arrAttributInfos(4) = "Année scolaire"
    
    With ThisWorkbook.Worksheets(strPage1)
        CreerTableau strNomWs:=strPage1, rngCelOrigine:=.Cells(byLigTabInfos, byColTabInfos), _
        iHaut:=4, iLarg:=2, iOrientation:=2, _
        arrAttribut:=arrAttributInfos, byCouleur:=byCouleurInfos, bLocked:=False
    End With
End Sub

Private Sub CreerTableauNombreClasses()
    ' *** DECLARATION VARIABLES ***
    Dim arrAttributClasses(1 To 1) As String
    Dim rngBtnValiderNbClasses As Range
    
    arrAttributClasses(1) = "Nombre de classes"
    
    With ThisWorkbook.Worksheets(strPage1)
        ' *** CREATION TABLEAU ***
        CreerTableau strNomWs:=strPage1, rngCelOrigine:=.Cells(byLigTabClasses, byColTabClasses), _
        iHaut:=1, iLarg:=2, iOrientation:=2, _
        arrAttribut:=arrAttributClasses, byCouleur:=byCouleurClasse, bLocked:=False

        ' *** AJOUT BOUTON 'VALIDER NOMBRE CLASSES' ***
        Set rngBtnValiderNbClasses = .Range(.Cells(byLigTabClasses, byColTabClasses + 3).Address)
        With .Buttons.Add(rngBtnValiderNbClasses.Left, rngBtnValiderNbClasses.Top, _
                          rngBtnValiderNbClasses.Width, rngBtnValiderNbClasses.Height)
            .Caption = "Valider le nombre de classes"
            .OnAction = "BtnValiderNombreClasses_Click"
            .Name = "BtnValiderNombreClasses"
        End With
        Set rngBtnValiderNbClasses = Nothing
    End With
End Sub

'@EntryPoint
Private Sub BtnValiderNombreClasses_Click()
    ' *** DECLARATION VARIABLES ***
    Dim byNbClasses As Byte
    
    ' *** REFRESH ECRAN OFF ***
    DisableUpdates
    
    ' *** VERIFICATION VALEUR ***
    If Not (IsNbClassesOK) Then
        MsgBox "ATTENTION: le nombre de classes n'est pas valide  (nombre min = " & byNbClasses_Min & ", nombre max = " & byNbClasses_Max & ")."
        GoTo EOP
    Else
        byNbClasses = GetNombreClasses
    End If
    
    ' *** MODIFICATION BOUTON ***
    With ThisWorkbook.Worksheets(strPage1).Buttons("BtnValiderNombreClasses")
        .LockedText = False
        .Caption = "Modifier le nombre de classes"
        .OnAction = "BtnModifierNombreClasses_Click"
        .Name = "BtnModifierNombreClasses"
    End With
    
    ' *** APPEL PROCEDURE ***
    CreerTableauClasses byNbClasses
    
    ' *** REFRESH ECRAN ON ***
    EnableUpdates

    ' *** MESSAGE INFORMATION ***
    MsgBox "Entrez maintenant le nom de chaque classe ainsi que le nombre d'élèves qui s'y trouve. " & _
           "Cela permettra au tableur de générer les listes d'élèves à compléter." & vbNewLine & vbNewLine & _
           "Si besoin, vous pouvez changer le nombre de classes à l'aide du bouton 'Modifier'. " & _
           "Une fois toutes les cases complétées, cliquez sur le bouton 'Valider' pour bloquer les données et passer à l'étape suivante." & vbNewLine & vbNewLine & _
           "INDICATION: Les noms de classes sont limités à 7 caractères. Utilisez donc des noms courts, " & _
           "par exemple '5ème 2', '5ème2', '5e 2' ou encore '5e2' pour désigner la classe de Cinquième 2."
    
    Exit Sub

EOP:
    ' *** REFRESH ECRAN ON ***
    EnableUpdates
End Sub

'@EntryPoint
Private Sub BtnModifierNombreClasses_Click()
    ' *** DECLARATION VARIABLES ***
    Dim byNbClasses As Byte
    
    ' *** REFRESH ECRAN OFF ***
    DisableUpdates
    
    ' *** VERIFICATION VALEUR ***
    If Not (IsNbClassesOK) Then
        MsgBox "ATTENTION: le nombre de classes n'est pas valide  (nombre min = " & byNbClasses_Min & ", nombre max = " & byNbClasses_Max & ")."
        GoTo EOP
    Else
        byNbClasses = GetNombreClasses
    End If
    
    ' *** SUPPRESSION CELLULES ***
    With ThisWorkbook.Worksheets(strPage1)
        .Range(.Cells(byLigTabClasses + 1, byColTabClasses), .Cells(byLigTabClasses + byNbClasses_Max + 10, byColTabClasses + 1)).Delete Shift:=xlUp
        .Buttons("BtnValiderClasses").Delete
    End With

    ' *** APPEL PROCEDURE ***
    CreerTableauClasses byNbClasses

EOP:
    ' *** REFRESH ECRAN ON ***
    EnableUpdates
End Sub

Private Sub CreerTableauClasses(ByVal byNbClasses As Byte)
    ' *** DECLARATION VARIABLES ***
    Dim arrAttributClasses(1 To 2) As String
    Dim rngBtnValiderClasses As Range
    
    ' *** AFFECTATION VARIABLES ***
    arrAttributClasses(1) = "Nom de la classe"
    arrAttributClasses(2) = "Nombre d'élèves"
    
    With ThisWorkbook.Worksheets(strPage1)
        ' *** CREATION TABLEAU ***
        CreerTableau strNomWs:=strPage1, rngCelOrigine:=.Cells(byLigTabClasses + 2, byColTabClasses), _
        iHaut:=byNbClasses + 1, iLarg:=2, iOrientation:=1, _
        arrAttribut:=arrAttributClasses, byCouleur:=byCouleurClasse, bLocked:=False
        
        ' *** AJOUT BOUTON 'VALIDER CLASSES' ***
        Set rngBtnValiderClasses = .Range(.Cells(byLigTabClasses + 2, byColTabClasses + 3).Address)
        With .Buttons.Add(rngBtnValiderClasses.Left, rngBtnValiderClasses.Top, _
                          rngBtnValiderClasses.Width, rngBtnValiderClasses.Height)
            .Caption = "Valider les classes"
            .OnAction = "BtnValiderClasses_Click"
            .Name = "BtnValiderClasses"
        End With
        Set rngBtnValiderClasses = Nothing
    End With
End Sub

'@EntryPoint
Private Sub BtnValiderClasses_Click()
    ' ***REFRESH ECRAN OFF ***
    DisableUpdates
    
    ' *** CONFIRMATION UTILISATEUR
    If Not (MsgBox("Confirmez-vous le nom de classe et le nombre d'élèves indiqués? " & _
                   "Il ne sera pas possible de les modifier par la suite.", vbYesNo) = vbYes) Then GoTo EOP
    
    ' *** VERIFICATION VALEUR ***
    If Not (IsDonneesClassesOK) Then
        MsgBox "ATTENTION: les données entrées pour les classes ne sont pas valides, cela empêche de passer à la prochaine étape." & vbNewLine & _
               "Cela peut provenir de trois éléments:" & vbNewLine & _
               "   - Nombre de classes qui ne correspond pas à la taille du tableau," & vbNewLine & _
               "   - Données manquantes dans le tableau," & vbNewLine & _
               "   - Nombre d'élèves incorrect (nombre min = " & byNbEleves_Min & ", nombre max = " & byNbEleves_Max & ")." & vbNewLine & _
               "Vérifiez ces trois propriétés et corrigez-les pour continuer."
        GoTo EOP
    End If
    
    ' *** SUPPRESSION BOUTONS ***
    With ThisWorkbook.Worksheets(strPage1)
        .Range(.Cells(byLigTabClasses, byColTabClasses), .Cells(byLigTabClasses + GetNombreClasses + 2, byColTabClasses + 1)).Locked = True
        ThisWorkbook.Worksheets(strPage1).Buttons.Delete
    End With

    ' *** APPEL PROCEDURE ***
    CreerDropdownCycle
    
    ' ***REFRESH ECRAN ON ***
    EnableUpdates

    ' *** MESSAGE INFORMATION ***
    MsgBox "Choisissez maintenant le cycle d'étude (2, 3 ou 4). Un tableau de compétences s'affichera alors en-dessous et changera en fonction du cycle choisi. " & vbNewLine & vbNewLine & _
           "Sélectionnez ensuite les compétences que vous évaluerez au cours de l'année (en fonction de votre matière, et pour l'ensemble des classes que vous avez indiqué). " & _
           "Pour cela, écrivez la lettre 'X' dans la case correspondant à la compétence notée." & vbNewLine & _
           "Une fois toutes vos compétences sélectionnées, cliquez sur 'Valider les compétences' pour passer à l'étape suivante." & vbNewLine & vbNewLine & _
           "INDICATION: Afin d'améliorer l'affichage, chaque compétence sera indiquée par une abréviation à la place de son nom complet. " & vbNewLine & _
           "En sélectionnant une compétence, une abréviation par défaut vous sera proposée. Cependant, si vous disposez d'une abréviation qui vous est propre, " & _
           "vous êtes libre de l'écrire dans la case correspondante. Elle sera alors utilisée à la place de l'abréviation par défaut" & vbNewLine & _
           "Comme précédemment, utilisez des abréviations courtes car la taille est limitée à 7 caractères."
    Exit Sub
    
EOP:
    ' ***REFRESH ECRAN ON ***
    EnableUpdates
End Sub

Private Sub CreerDropdownCycle()
    ' *** DECLARATION VARIABLES ***
    Dim byLigChoixCycle As Byte
    Dim rngDrpChoixCycle As Range
    Dim rngBtnValiderCompetences As Range

    ' *** AFFECTATION VARIABLES ***
    byLigChoixCycle = byLigTabClasses + GetNombreClasses + 5
    
    With ThisWorkbook.Worksheets(strPage1)
        ' *** MISE EN FORME ***
        With .Range(.Cells(byLigChoixCycle, byColTabCompet), _
                    .Cells(byLigChoixCycle, byColTabCompet + 1)).Borders
            .ColorIndex = 1
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With .Cells(byLigChoixCycle, byColTabCompet)
            .Interior.ColorIndex = byCouleurCompet_1
            .Value = "Choix du cycle"
        End With
        
        ' *** AJOUT COMBOBOX CHOIX CYCLE ***
        Set rngDrpChoixCycle = .Cells(byLigChoixCycle, byColTabCompet + 1)
        With .DropDowns.Add(rngDrpChoixCycle.Left, rngDrpChoixCycle.Top, _
                            rngDrpChoixCycle.Width, rngDrpChoixCycle.Height)
            .DropDownLines = 3
            .AddItem "Cycle 2", 1
            .AddItem "Cycle 3", 2
            .AddItem "Cycle 4", 3
            .OnAction = "drpChoixCycle_Change"
            .Name = "drpChoixCycle"
        End With
    
        ' *** AJOUT BOUTON 'VALIDER CLASSES' ***
        Set rngBtnValiderCompetences = .Cells(byLigChoixCycle, byColTabCompet + 3)
        With .Buttons.Add(rngBtnValiderCompetences.Left, rngBtnValiderCompetences.Top, _
                          rngBtnValiderCompetences.Width, rngBtnValiderCompetences.Height)
            .Caption = "Valider les compétences"
            .OnAction = "BtnValiderCompetences_Click"
            .Name = "BtnValiderCompetences"
        End With
    End With
End Sub

'@EntryPoint
Public Sub DrpChoixCycle_Change()
    ' *** DECLARATION VARIABLES ***
    Dim iDrpValue As Byte
    Dim byLigTabCompet As Byte
    
    ' *** AFFECTATION VARIABLES ***
    iDrpValue = ThisWorkbook.Worksheets(strPage1).DropDowns("drpChoixCycle").Value
    byLigTabCompet = byLigTabClasses + GetNombreClasses + 7
    
    ' *** REFRESH ECRAN OFF ***
    DisableUpdates
    
    ' *** OPERATION ***
    If iDrpValue = 1 Or iDrpValue = 2 Or iDrpValue = 3 Then
        With ThisWorkbook.Worksheets(strPage1)
            .Range(.Cells(byLigTabCompet, byColTabCompet), .Cells(byLigTabCompet + 100, byColTabCompet)).EntireRow.Delete
        End With
        CreerTableauChoixCompetences iDrpValue
    End If
    
    ' *** REFRESH ECRAN ON ***
    EnableUpdates
End Sub

Private Sub CreerTableauChoixCompetences(ByVal iDrpValue As Byte)
    ' *** DECLARATION VARIABLES ***
    Dim byLigTabCompet As Byte
    Dim byLigDomaine As Byte
    Dim byLigCompetence As Byte
    Dim arrAttributCompet(1 To 4) As String
    Dim arrDomaines As Variant
    Dim arrCompetences As Variant
    Dim iDomaine As Byte
    Dim iCompetence As Byte
    Dim byNbCompetences As Byte
    Dim rowCompetence As Range
    
    ' *** VALEUR ATTRIBUTS ***
    arrAttributCompet(1) = "Domaines / Composantes"
    arrAttributCompet(2) = "Compétences"
    arrAttributCompet(3) = "Sélection"
    arrAttributCompet(4) = "Abréviation"

    ' *** AFFECTATION VARIABLES ***
    arrDomaines = GetArrayDomaines
    arrCompetences = GetArrayCompetences(iDrpValue + 1)
    byLigTabCompet = byLigTabClasses + GetNombreClasses + 7
    byLigCompetence = byLigTabCompet + 1
    byNbCompetences = GetNombreCompetences(iDrpValue + 1)

    With ThisWorkbook.Worksheets(strPage1)
        ' *** AJOUT DOMAINES & COMPETENCES ***
        For iDomaine = 1 To 8
            .Cells(byLigCompetence, byColTabCompet).Value = arrDomaines(iDomaine, 1)
            byLigDomaine = byLigCompetence
            For iCompetence = LBound(arrCompetences(iDomaine)) To UBound(arrCompetences(iDomaine))
                .Cells(byLigCompetence, byColTabCompet + 1).Value = arrCompetences(iDomaine)(iCompetence)
                byLigCompetence = byLigCompetence + 1
            Next iCompetence
            .Range(.Cells(byLigDomaine, byColTabCompet), .Cells(byLigCompetence - 1, byColTabCompet)).MergeCells = True
        Next iDomaine
        
        ' *** MISE EN FORME ***
        CreerTableau strNomWs:=strPage1, rngCelOrigine:=.Cells(byLigTabCompet, byColTabCompet), _
                     iHaut:=byNbCompetences + 1, iLarg:=4, iOrientation:=1, _
                     arrAttribut:=arrAttributCompet, byCouleur:=byCouleurCompet_1
                     
        .Range(.Cells(byLigTabCompet + 1, byColTabCompet + 2), _
               .Cells(byLigTabCompet + byNbCompetences, byColTabCompet + 3)).Locked = False
        With .Range(.Cells(byLigTabCompet + 1, byColTabCompet), _
                    .Cells(byLigTabCompet + byNbCompetences, byColTabCompet + 1))
            .HorizontalAlignment = xlHAlignLeft
            .WrapText = True
            .Rows.AutoFit
            For Each rowCompetence In .Rows
                rowCompetence.RowHeight = 20 / 14.4 * rowCompetence.RowHeight
            Next rowCompetence
        End With
    End With
End Sub

'@EntryPoint
Private Sub BtnValiderCompetences_Click()
    Dim byNbCompetences As Byte
    Dim byLigTabCompet As Byte

    If Not (IsInfosOK) Then
        MsgBox "Il manque des données dans le tableau d'informations (tableau gris). " & _
               "Merci de compléter tous les cases avant de passer à l'étape suivante."
        Exit Sub
    End If
    
    If Not (MsgBox("Confirmez-vous la sélection des compétences à évaluer? " & _
                   "Il ne sera pas possible de la modifier par la suite.", vbYesNo) = vbYes) Then Exit Sub

    If Not (IsCompetOK) Then
        MsgBox "ATTENTION: votre choix de compétences n'est pas valide. Cela peut provenir de deux éléments: " & vbNewLine & _
               "    - Moins de deux compétences sélectionnées pour évaluation" & vbNewLine & _
               "    - Abréviations utilisées trop longues (7 caractères max)" & vbNewLine & _
               "Vérifiez ces deux propriétés et corrigez-les pour continuer."
        Exit Sub
    End If
    
    With ThisWorkbook.Worksheets(strPage1)
        byLigTabCompet = byLigTabClasses + GetNombreClasses + 7
        byNbCompetences = GetNombreCompetences(.Shapes("drpChoixCycle").ControlFormat.Value + 1)
        
        DisableUpdates
        .Cells.Font.Bold = False
        .Range(.Cells(byLigTabLogiciel + 1, byColTabLogiciel + 1).Address).Select
        .Range(.Cells(byLigTabInfos, byColTabInfos + 1), .Cells(byLigTabInfos + 3, byColTabInfos + 1)).Locked = True
        .Range(.Cells(byLigTabCompet + 1, byColTabCompet + 2), .Cells(byLigTabCompet + byNbCompetences, byColTabCompet + 3)).Locked = True
        .Shapes("drpChoixCycle").ControlFormat.Enabled = False
        DeleteAllButtons ThisWorkbook.Worksheets(strPage1)
        UnprotectWorkbook
        AddWorksheet (strPage2)
        InitPage2
        ProtectAllWorksheets
        ProtectWorkbook
        EnableUpdates
    End With
    
    MsgBox "Vous arrivez maintenant sur la page de gestion des listes de classes. " & _
           "Vous pouvez ici entrer le noms et prénom des élèves de chaque classe, dans les colonnes correspondantes. " & vbNewLine & _
           "Une fois terminé, cliquez sur le bouton 'Valider les listes' pour passer à l'étape suivante." & vbNewLine & vbNewLine & _
           "INDICATION: Afin de gérer d'éventuelles modifications de classe en cours d'année (arrivée ou départ d'un élève), " & _
           "il est possible de d'ajouter, de transférer ou de supprimer un élève en cliquant sur le bouton 'Modifier les listes' " & _
           "(qui s'affichera à la place du bouton 'Valider les listes')."
End Sub


