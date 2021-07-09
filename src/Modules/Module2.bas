Attribute VB_Name = "Module2"

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
'                               Module 2 - Listes
' *******************************************************************************
'
'   Fonctions publiques
'       GetIndiceEleve(ByVal strEleve As String, ByVal byClasse As Byte, ByVal bValeurExacte As Boolean) As Byte
'
'   Procédures publiques
'       InitPage2()
'       AjouterEleve(ByVal byClasse As Byte, ByVal byEleve As Byte, ByVal strEleve As String)
'       SupprimerEleve(ByVal byClasse As Byte, ByVal byEleve As Byte)
'       TransfererEleve(ByVal byClasseSource As Byte, ByVal byEleveSource As Byte, ByVal byClasseDest As Byte, ByVal byEleveDest As Byte, ByVal strEleve As String)
'
'   Fonctions privées
'       IsListesOK() As Boolean
'
'   Procédures privées
'       BtnValiderListes_Click()
'       BtnModifierListes_Click()
'
' *******************************************************************************

Option Explicit

' *******************************************************************************
'                               Fonctions publiques
' *******************************************************************************

Public Function GetIndiceEleve(ByVal strEleve As String, ByVal byClasse As Byte, ByVal bValeurExacte As Boolean) As Byte
    ' *** DECLARATION VARIABLES ***
    Dim byNbEleves As Byte
    Dim byEleve As Byte
    
    ' *** AFFECTATION VARIABLES ***
    byNbEleves = GetNombreEleves(byClasse)
    GetIndiceEleve = -1
    
    With ThisWorkbook
        ' *** RECHERCHE INDICE EXACT ***
        If bValeurExacte Then
            For byEleve = 1 To byNbEleves
                If StrComp(strEleve, .Worksheets(strPage2).Cells(byLigListePage2 + byEleve, 2 * byClasse - 1).Value) = 0 Then
                    GetIndiceEleve = byEleve
                    Exit For
                End If
            Next byEleve
            
            ' *** RECHERCHE INDICE POUR INSERTION ***
        Else
            For byEleve = 1 To byNbEleves
                If StrComp(strEleve, .Worksheets(strPage2).Cells(byLigListePage2 + byEleve, 2 * byClasse - 1).Value) = -1 Then
                    GetIndiceEleve = byEleve
                    Exit For
                ElseIf byEleve = byNbEleves Then GetIndiceEleve = byNbEleves + 1
                End If
            Next byEleve
        End If
    End With
End Function

' *******************************************************************************
'                               Procédures publiques
' *******************************************************************************
Public Sub InitPage2()
    ' *** DECLARATION VARIABLES ***
    Dim byClasse As Byte
    Dim byNbClasses As Byte
    Dim byNbEleves As Byte
    Dim byColClasse As Byte
    Dim arrNomClasse(1 To 1) As String
    Dim rngBtnValiderListes As Range
    
    ' *** AFFECTATION VARIABLES ***
    byNbClasses = GetNombreClasses
        
    With ThisWorkbook.Worksheets(strPage2)
        For byClasse = 1 To byNbClasses
            byNbEleves = GetNombreEleves(byClasse)
            byColClasse = 2 * byClasse - 1
            arrNomClasse(1) = GetNomClasse(byClasse)
            
            CreerTableau strNomWs:=strPage2, rngCelOrigine:=.Cells(byLigListePage2, byColClasse), _
                         iHaut:=byNbEleves + 1, iLarg:=1, iOrientation:=1, _
                         arrAttribut:=arrNomClasse, byCouleur:=byCouleurClasse, bLocked:=False
                         
            With .Range(.Cells(byLigListePage2 + 1, byColClasse), .Cells(byLigListePage2 + byNbEleves, byColClasse))
                .HorizontalAlignment = xlHAlignLeft
                .EntireColumn.ColumnWidth = 40
            End With
            .Columns(byColClasse + 1).ColumnWidth = 5
        Next byClasse
        
        ' *** CREATION BOUTON 'VALIDER LISTE' ***
        Set rngBtnValiderListes = .Cells(1, byColClasse + 2)
        With .Buttons.Add(rngBtnValiderListes.Left, rngBtnValiderListes.Top, _
                          rngBtnValiderListes.Width, rngBtnValiderListes.Height)
            .Caption = "Valider les listes"
            .OnAction = "BtnValiderListes_Click"
            .Name = "BtnValiderListes"
        End With
    End With
    
    ' *** FORMATAGE PAGE ***
    FreezePanes ActiveWindow, byLigListePage2, 0
End Sub

Public Sub AjouterEleve(ByVal byClasse As Byte, ByVal byEleve As Byte, ByVal strEleve As String)
    ' *** DECLARATION VARIABLES ***
    Dim byNbEleves As Byte
    Dim byColClasse As Byte
    Dim lColMax As Long
    
    ' *** AFFECTATION VARIABLES ***
    byNbEleves = GetNombreEleves(byClasse)
    byColClasse = 2 * byClasse - 1
    
    ' *** UPDATES OFF ***
    DisableUpdates
    
    ' *** MODIFICATION PAGE 1 - ACCUEIL ***
    SetNombreEleves byClasse, byNbEleves + 1
    
    ' *** MODIFICATION PAGE 2 - LISTES ***
    With ThisWorkbook.Worksheets(strPage2)
        Select Case byEleve
        Case Is = 1                     ' Cas 1: élève en début de liste
            .Cells(byLigListePage2 + byEleve + 1, byColClasse).Insert xlShiftDown, xlFormatFromRightOrBelow
            .Cells(byLigListePage2 + byEleve + 1, byColClasse).Value = .Cells(byLigListePage2 + byEleve, byColClasse).Value
        Case 2 To byNbEleves            ' Cas 2: élève au milieu de la liste
            .Cells(byLigListePage2 + byEleve, byColClasse).Insert xlShiftDown, xlFormatFromLeftOrAbove
        Case Is = byNbEleves + 1        ' Cas 3: élève en fin de liste
            .Cells(byLigListePage2 + byEleve - 1, byColClasse).Insert xlShiftDown, xlFormatFromLeftOrAbove
            .Cells(byLigListePage2 + byEleve - 1, byColClasse).Value = .Cells(byLigListePage2 + byEleve, byColClasse).Value
        End Select
        .Cells(byLigListePage2 + byEleve, byColClasse).Value = strEleve
        .Cells.Locked = True
    End With
        
    If ThisWorkbook.Worksheets.Count > 3 Then
        ' *** MODIFICATION PAGE 3 - NOTES ***
        With ThisWorkbook.Worksheets(GetNomPage3(byClasse))
            lColMax = .UsedRange.Columns.Count - 1
            Select Case byEleve
            Case 1
                .Range(.Cells(byLigListePage3 + byEleve + 1, 1), .Cells(byLigListePage3 + byEleve + 1, lColMax)).Insert xlShiftDown, xlFormatFromRightOrBelow
                .Range(.Cells(byLigListePage3 + byEleve + 1, 1), .Cells(byLigListePage3 + byEleve + 1, 2)).MergeCells = True
                .Range(.Cells(byLigListePage3 + byEleve + 1, 1), .Cells(byLigListePage3 + byEleve + 1, lColMax)).Value = _
                        .Range(.Cells(byLigListePage3 + byEleve, 1), .Cells(byLigListePage3 + byEleve, lColMax)).Value
                .Range(.Cells(byLigListePage3 + byEleve, 1), .Cells(byLigListePage3 + byEleve, lColMax)).ClearContents
            Case 2 To byNbEleves
                .Range(.Cells(byLigListePage3 + byEleve, 1), .Cells(byLigListePage3 + byEleve, lColMax)).Insert xlShiftDown, xlFormatFromLeftOrAbove
                .Range(.Cells(byLigListePage3 + byEleve, 1), .Cells(byLigListePage3 + byEleve, lColMax)).MergeCells = True
            Case byNbEleves + 1
                .Range(.Cells(byLigListePage3 + byEleve - 1, 1), .Cells(byLigListePage3 + byEleve - 1, lColMax)).Insert xlShiftDown, xlFormatFromRightOrBelow
                .Range(.Cells(byLigListePage3 + byEleve - 1, 1), .Cells(byLigListePage3 + byEleve - 1, 2)).MergeCells = True
                .Range(.Cells(byLigListePage3 + byEleve - 1, 1), .Cells(byLigListePage3 + byEleve - 1, lColMax)).Value = _
                        .Range(.Cells(byLigListePage3 + byEleve, 1), .Cells(byLigListePage3 + byEleve, lColMax)).Value
                .Range(.Cells(byLigListePage3 + byEleve, 1), .Cells(byLigListePage3 + byEleve, lColMax)).ClearContents
            End Select
            .Cells(byLigListePage3 + byEleve, 1).Value = strEleve
        End With
        ' *** MODIFICATION PAGE 4 - BILAN ***
        With ThisWorkbook.Worksheets(GetNomPage4(byClasse))
            lColMax = .UsedRange.Columns.Count - 1
            Select Case byEleve
            Case 1
                .Range(.Cells(byLigListePage4 + byEleve + 1, 1), .Cells(byLigListePage4 + byEleve + 1, lColMax)).Insert xlShiftDown, xlFormatFromRightOrBelow
                .Range(.Cells(byLigListePage4 + byEleve + 1, 1), .Cells(byLigListePage4 + byEleve + 1, lColMax)).Value = _
                        .Range(.Cells(byLigListePage4 + byEleve, 1), .Cells(byLigListePage4 + byEleve, lColMax)).Value
                .Range(.Cells(byLigListePage4 + byEleve, 1), .Cells(byLigListePage4 + byEleve, lColMax)).ClearContents
            Case 2 To byNbEleves
                .Range(.Cells(byLigListePage4 + byEleve, 1), .Cells(byLigListePage4 + byEleve, lColMax)).Insert xlShiftDown, xlFormatFromLeftOrAbove
            Case byNbEleves + 1
                .Range(.Cells(byLigListePage4 + byEleve - 1, 1), .Cells(byLigListePage4 + byEleve - 1, lColMax)).Insert xlShiftDown, xlFormatFromRightOrBelow
                .Range(.Cells(byLigListePage4 + byEleve - 1, 1), .Cells(byLigListePage4 + byEleve - 1, lColMax)).Value = _
                        .Range(.Cells(byLigListePage4 + byEleve, 1), .Cells(byLigListePage4 + byEleve, lColMax)).Value
                .Range(.Cells(byLigListePage4 + byEleve, 1), .Cells(byLigListePage4 + byEleve, lColMax)).ClearContents
            End Select
            .Cells(byLigListePage4 + byEleve, 1).Value = strEleve
        End With
    End If
    
    ' *** UPDATES ON ***
    EnableUpdates
    
    ' *** MESSAGE INFORMATION ***
    MsgBox "Élève ajouté.", vbInformation, "Ajout d'élève"
    
End Sub

Public Sub SupprimerEleve(ByVal byClasse As Byte, ByVal byEleve As Byte)
    
    ' *** DECLARATION VARIABLES ***
    Dim byNbEleves As Byte
    Dim byColClasse As Byte
    Dim lColMax As Long
    
    ' *** AFFECTATION VARIABLES ***
    byNbEleves = GetNombreEleves(byClasse)
    byColClasse = 2 * byClasse - 1
    
    ' *** UPDATES OFF ***
    DisableUpdates
    
    ' *** MODIFICATION PAGE 1 - ACCUEIL ***
    SetNombreEleves byClasse, byNbEleves - 1
    
    ' *** MODIFICATION PAGE 2 - LISTES ***
    With ThisWorkbook.Worksheets(strPage2)
        Select Case byEleve
        Case Is = 1                     ' Cas 1: élève en début de liste
            .Cells(byLigListePage2 + byEleve, byColClasse).Value = .Cells(byLigListePage2 + byEleve + 1, byColClasse).Value
            .Cells(byLigListePage2 + byEleve + 1, byColClasse).Delete xlShiftUp
        Case 2 To byNbEleves - 1        ' Cas 2: élève au milieu de la liste
            .Cells(byLigListePage2 + byEleve, byColClasse).Delete xlShiftUp
        Case Is = byNbEleves            ' Cas 3: élève en fin de liste
            .Cells(byLigListePage2 + byEleve, byColClasse).Value = .Cells(byLigListePage2 + byEleve - 1, byColClasse).Value
            .Cells(byLigListePage2 + byEleve - 1, byColClasse).Delete xlShiftUp
        End Select
        .Cells.Locked = True
    End With
        
    If ThisWorkbook.Worksheets.Count > 3 Then
        ' *** MODIFICATION PAGE 3 - NOTES ***
        With ThisWorkbook.Worksheets(GetNomPage3(byClasse))
            lColMax = .UsedRange.Columns.Count - 1
            Select Case byEleve
            Case Is = 1                     ' Cas 1: élève en début de liste
                .Range(.Cells(byLigListePage3 + byEleve, 1), .Cells(byLigListePage3 + byEleve, lColMax)).Value = _
                        .Range(.Cells(byLigListePage3 + byEleve + 1, 1), .Cells(byLigListePage3 + byEleve + 1, lColMax)).Value
                .Range(.Cells(byLigListePage3 + byEleve + 1, 1), .Cells(byLigListePage3 + byEleve + 1, lColMax)).Delete xlShiftUp
            Case 2 To byNbEleves - 1        ' Cas 2: élève au milieu de la liste
                .Range(.Cells(byLigListePage3 + byEleve, 1), .Cells(byLigListePage3 + byEleve, lColMax)).Delete xlShiftUp
            Case Is = byNbEleves            ' Cas 3: élève en fin de liste
                .Range(.Cells(byLigListePage3 + byEleve, 1), .Cells(byLigListePage3 + byEleve, lColMax)).Value = _
                        .Range(.Cells(byLigListePage3 + byEleve - 1, 1), .Cells(byLigListePage3 + byEleve - 1, lColMax)).Value
                .Range(.Cells(byLigListePage3 + byEleve - 1, 1), .Cells(byLigListePage3 + byEleve - 1, lColMax)).Delete xlShiftUp
            End Select
        End With
        ' *** MODIFICATION PAGE 4 - BILAN ***
        With ThisWorkbook.Worksheets(GetNomPage4(byClasse))
            lColMax = .UsedRange.Columns.Count - 1
            Select Case byEleve
            Case Is = 1                     ' Cas 1: élève en début de liste
                .Range(.Cells(byLigListePage4 + byEleve, 1), .Cells(byLigListePage4 + byEleve, lColMax)).Value = _
                        .Range(.Cells(byLigListePage4 + byEleve + 1, 1), .Cells(byLigListePage4 + byEleve + 1, lColMax)).Value
                .Range(.Cells(byLigListePage4 + byEleve + 1, 1), .Cells(byLigListePage4 + byEleve + 1, lColMax)).Delete xlShiftUp
            Case 2 To byNbEleves - 1        ' Cas 2: élève au milieu de la liste
                .Range(.Cells(byLigListePage4 + byEleve, 1), .Cells(byLigListePage4 + byEleve, lColMax)).Delete xlShiftUp
            Case Is = byNbEleves            ' Cas 3: élève en fin de liste
                .Range(.Cells(byLigListePage4 + byEleve, 1), .Cells(byLigListePage4 + byEleve, lColMax)).Value = _
                        .Range(.Cells(byLigListePage4 + byEleve - 1, 1), .Cells(byLigListePage4 + byEleve - 1, lColMax)).Value
                .Range(.Cells(byLigListePage4 + byEleve - 1, 1), .Cells(byLigListePage4 + byEleve - 1, lColMax)).Delete xlShiftUp
            End Select
        End With
    End If
    
    ' *** UPDATES ON ***
    EnableUpdates
    
    ' *** MESSAGE INFORMATION ***
    MsgBox "Élève supprimé.", vbInformation, "Suppression d'élève"
End Sub

Public Sub TransfererEleve(ByVal byClasseSource As Byte, ByVal byEleveSource As Byte, ByVal byClasseDest As Byte, ByVal byEleveDest As Byte, ByVal strEleve As String)
     ' *** OPERATION 1 : AJOUT DANS NOUVELLE CLASSE ***
     AjouterEleve byClasseDest, byEleveDest, strEleve
     
     ' Il n'y a pas de transfert de notes car rien ne garanti l'homogénéïté des évaluations entre plusieurs classes.
    
     ' *** OPERATION 2 : SUPPRESSION DANS ANCIENNE CLASSE ***
     SupprimerEleve byClasseSource, byEleveSource
End Sub

' *******************************************************************************
'                               Fonctions privées
' *******************************************************************************

Private Function IsListesOK() As Boolean
    Dim byClasse As Byte
    Dim byNbClasses As Byte
    Dim byColClasse As Byte
    Dim byNbEleves As Byte
    
    IsListesOK = False
    byNbClasses = GetNombreClasses
    
    With ThisWorkbook.Worksheets(strPage2)
        For byClasse = 1 To byNbClasses
            byColClasse = 2 * byClasse - 1
            byNbEleves = GetNombreEleves(byClasse)
            If Application.WorksheetFunction.CountA(.Range(.Cells(byLigListePage2 + 1, byColClasse), _
                                                           .Cells(byLigListePage2 + byNbEleves, byColClasse))) <> byNbEleves Then Exit Function
        Next byClasse
    End With
    
    IsListesOK = True
End Function

' *******************************************************************************
'                               Procédures privées
' *******************************************************************************

'@EntryPoint
Private Sub BtnValiderListes_Click()
    Dim byClasse            As Byte
    Dim byNbEleves          As Byte
    Dim byNbClasses         As Byte
    Dim byAvancement        As Byte
    Dim byAvancementTotal   As Byte
    
    ' *** VERIFICATION LISTES VALIDES ***
    If Not (IsListesOK) Then
        MsgBox "Les listes de classes sont incomplètes. Merci de compléter tous les cases avant de passer à l'étape suivante." & vbNewLine & vbNewLine & _
               "INDICATION: si le nombre d'élève réel ne correspond pas au nombre de cases disponibles, complétez tout de même la totalité des listes " & _
               "et cliquez sur le bouton 'Valider les listes'. Un bouton 'Modifier les listes' apparaitra alors à la place du précédent, " & _
               "et vous permettra d'ajouter les élèves manquants ou de supprimer les élèves en trop."
        Exit Sub
    End If
    ' *** DEMANDE CONFIRMATION UTILISATEUR ***
    If Not (MsgBox("Confirmez-vous les listes de classes indiquées ? " & _
                   "Si besoin vous pourrez toujours modifier les élèves un par un, " & _
                   "mais il sera impossible de re-générer intégralement les listes.", vbYesNo) = vbYes) Then Exit Sub

    ' *** REFRESH ECRAN OFF ***
    UserForm5.Show vbModeless
    UnprotectWorkbook
    DisableUpdates
    
    ' *** AFFECTATION VARIABLES ***
    byNbClasses = GetNombreClasses
    byAvancement = 0
    byAvancementTotal = 2 * byNbClasses
    
    For byClasse = 1 To byNbClasses
        byNbEleves = GetNombreEleves(byClasse)
        
        ' *** AJOUT PAGE 3 ***
        AddWorksheet (GetNomPage3(byClasse))
        InitPage3 byClasse, byNbEleves
        byAvancement = byAvancement + 1
        UserForm5.updateAvancement byAvancement, byAvancementTotal

        ' *** AJOUT PAGE 4 ***
        AddWorksheet (GetNomPage4(byClasse))
        InitPage4 byClasse, byNbEleves
        byAvancement = byAvancement + 1
        UserForm5.updateAvancement byAvancement, byAvancementTotal
    Next byClasse
    
    ' *** MODIFICATION BOUTON ***
    With ThisWorkbook.Worksheets(strPage2)
        .Activate
        '        With .Buttons("BtnValiderListes")
        '            .LockedText = False
        '            .Caption = "Modifier les listes"
        '            .OnAction = "BtnModifierListes_Click"
        '            '.Name = "BtnModifierListes"
        '        End With
    End With
    
    ' *** BLOQUAGE PAGE 2 ***
    ThisWorkbook.Worksheets(strPage2).Cells.Locked = True
    
    ' *** PROTECTION + REFRESH ECRAN ON ***
    ProtectAllWorksheets
    ProtectWorkbook
    Unload UserForm5
    EnableUpdates
    
    ' *** MESSAGE INFORMATION ***
    MsgBox "Tableau 'Notes' et 'Bilan' créés avec succès !"
        
End Sub

'@EntryPoint
Private Sub BtnModifierListes_Click()
    UserForm1.Show
End Sub


