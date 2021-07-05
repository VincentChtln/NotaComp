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
'
'   Fonctions publiques
'
'   Procédures publiques
'
'   Fonctions privées
'
'   Procédures privées
'
' *******************************************************************************

Option Explicit

' *******************************************************************************
'                               Fonctions publiques
' *******************************************************************************

Public Function getIndiceEleve(ByVal strEleve As String, ByVal byClasse As Byte, ByVal bValeurExacte As Boolean) As Byte
    ' *** DECLARATION VARIABLES ***
    Dim byNbEleves As Byte
    Dim byEleve As Byte
    
    ' *** AFFECTATION VARIABLES ***
    byNbEleves = getNombreEleves(byClasse)
    getIndiceEleve = -1
    
    With ThisWorkbook
        ' *** RECHERCHE INDICE EXACT ***
        If bValeurExacte Then
            For byEleve = 1 To byNbEleves
                If StrComp(strEleve, .Worksheets(strPage2).Cells(byLigListePage2 + byEleve, 2 * byClasse - 1).Value) = 0 Then
                    getIndiceEleve = byEleve
                    Exit For
                End If
            Next byEleve
            
            ' *** RECHERCHE INDICE POUR INSERTION ***
        Else
            For byEleve = 1 To byNbEleves
                If StrComp(strEleve, .Worksheets(strPage2).Cells(byLigListePage2 + byEleve, 2 * byClasse - 1).Value) = -1 Then
                    getIndiceEleve = byEleve
                    Exit For
                ElseIf byEleve = byNbEleves Then getIndiceEleve = byNbEleves + 1
                End If
            Next byEleve
        End If
    End With
End Function

' *******************************************************************************
'                               Procédures publiques
' *******************************************************************************
Public Sub creerTableauListeClasses()
    ' *** DECLARATION VARIABLES ***
    Dim byClasse As Byte
    Dim byNbClasses As Byte
    Dim byNbEleves As Byte
    Dim byColClasse As Byte
    Dim arrNomClasse(1 To 1) As String
    Dim rngBtnValiderListes As Range
    Dim btnValiderListes As Variant
    
    ' *** AFFECTATION VARIABLES ***
    byNbClasses = getNombreClasses
        
    With ThisWorkbook.Worksheets(strPage2)
        For byClasse = 1 To byNbClasses
            byNbEleves = getNombreEleves(byClasse)
            byColClasse = 2 * byClasse - 1
            arrNomClasse(1) = getNomClasse(byClasse)
            
            creerTableau strNomWs:=strPage2, rngCelOrigine:=.Cells(byLigListePage2, byColClasse), _
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
            .OnAction = "btnValiderListes_Click"
            .Name = "btnValiderListes"
        End With
    End With
    
    ' *** FORMATAGE PAGE ***
    freezePanes ActiveWindow, byLigListePage2, 0
End Sub

Public Sub ajouterEleve(ByVal byClasse As Byte, ByVal byEleve As Byte, ByVal strEleve As String)

End Sub

Public Sub supprimerEleve(ByVal byClasse As Byte, ByVal byEleve As Byte)

End Sub

Public Sub transfereEleve(ByVal byClasseSource As Byte, ByVal byEleveSource As Byte, ByVal byClasseDest As Byte, ByVal byEleveDest As Byte)

End Sub

' *******************************************************************************
'                               Fonctions privées
' *******************************************************************************

Private Function isListesOK() As Boolean
    Dim byClasse As Byte
    Dim byNbClasses As Byte
    Dim byColClasse As Byte
    Dim byNbEleves As Byte
    
    isListesOK = False
    byNbClasses = getNombreClasses
    
    With ThisWorkbook.Worksheets(strPage2)
        For byClasse = 1 To byNbClasses
            byColClasse = 2 * byClasse - 1
            byNbEleves = getNombreEleves(byClasse)
            If Application.WorksheetFunction.CountA(.Range(.Cells(byLigListePage2 + 1, byColClasse), _
                                                           .Cells(byLigListePage2 + byNbEleves, byColClasse))) <> byNbEleves Then Exit Function
        Next byClasse
    End With
    
    isListesOK = True
End Function

' *******************************************************************************
'                               Procédures privées
' *******************************************************************************

Private Sub btnValiderListes_Click()
    Dim byClasse            As Byte
    Dim byNbEleves          As Byte
    Dim byNbClasses         As Byte
    Dim byAvancement        As Byte
    Dim byAvancementTotal   As Byte
    
    ' *** VERIFICATION LISTES VALIDES ***
    If Not (isListesOK) Then
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
    unprotectWorkbook
    disableUpdates
    
    ' *** AFFECTATION VARIABLES ***
    byNbClasses = getNombreClasses
    byAvancement = 0
    byAvancementTotal = 2 * byNbClasses
    
    For byClasse = 1 To byNbClasses
        byNbEleves = getNombreEleves(byClasse)
        
        ' *** AJOUT PAGE 3 ***
        addWorksheet (getNomPage3(byClasse))
        initPage3 byClasse, byNbEleves
        byAvancement = byAvancement + 1
        UserForm5.updateAvancement byAvancement, byAvancementTotal

        ' *** AJOUT PAGE 4 ***
        addWorksheet (getNomPage4(byClasse))
        initPage4 byClasse, byNbEleves
        byAvancement = byAvancement + 1
        UserForm5.updateAvancement byAvancement, byAvancementTotal
    Next byClasse
    
    ' *** MODIFICATION BOUTON ***
    With ThisWorkbook.Worksheets(strPage2)
        .Activate
        '        With .Buttons("btnValiderListes")
        '            .LockedText = False
        '            .Caption = "Modifier les listes"
        '            .OnAction = "btnModifierListes_Click"
        '            '.Name = "btnModifierListes"
        '        End With
    End With
    
    ' *** BLOQUAGE PAGE 2 ***
    ThisWorkbook.Worksheets(strPage2).Cells.Locked = True
    
    ' *** PROTECTION + REFRESH ECRAN ON ***
    protectAllWorksheets
    protectWorkbook
    Unload UserForm5
    enableUpdates
    
    ' *** MESSAGE INFORMATION ***
    MsgBox "Tableau 'Notes' et 'Bilan' créés avec succès !"
        
End Sub

Private Sub btnModifierListes_Click()
    UserForm1.Show
End Sub


