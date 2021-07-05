Attribute VB_Name = "Module4"

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

Sub test()
    Const byClasse As Byte = 1
    Dim byNbEleves As Byte
    disableUpdates
    byNbEleves = getNombreEleves(byClasse)
    initPage4 byClasse, byNbEleves
    enableUpdates
End Sub

' *******************************************************************************
'                               Procédures publiques
' *******************************************************************************

Public Sub initPage4(byClasse As Byte, byNbEleves As Byte)
    Dim wsPage2                     As Worksheet
    Dim wsPage4                     As Worksheet
    Dim rngBtnActualiserResultats   As Range
    Dim arrDomaines                 As Variant
    Dim arrChoixCompet              As Variant
    Dim arrTampon                   As Variant
    Dim byDomaine                   As Byte
    Dim byColArray                  As Byte
    
    Set wsPage4 = ThisWorkbook.Worksheets(getNomPage4(byClasse))
    
    With wsPage4
        ' *** FORMATAGE LIGNES / COLONNES ***
        .Rows.RowHeight = 15
        Union(.Rows(1), .Rows(2), .Rows(3)).RowHeight = 25
        .Columns.ColumnWidth = 7
        .Columns(1).ColumnWidth = 40
        
        ' *** BOUTON 'ACTUALISER RESULTATS' ***
        Set rngBtnActualiserResultats = .Range("A1")
        With .Buttons.Add(rngBtnActualiserResultats.Left, rngBtnActualiserResultats.Top, _
                          rngBtnActualiserResultats.Width, rngBtnActualiserResultats.Height)
            .Caption = "Actualiser résultats"
            .OnAction = "btnActualiserResultats_Click"
            .Name = "btnActualiserResultats"
        End With
        
        ' *** NOM CLASSE ***
        With .Range("A2")
            .Value = getNomClasse(byClasse)
            .Interior.ColorIndex = byCouleurClasse
        End With
        .Range("A2:A3").MergeCells = True
        
        ' *** LISTE CLASSE ***
        Set wsPage2 = ThisWorkbook.Worksheets(strPage2)
        With .Range(.Cells(byLigListePage4 + 1, 1), .Cells(byLigListePage4 + byNbEleves, 1))
            .Value = wsPage2.Range(wsPage2.Cells(byLigListePage2 + 1, 2 * byClasse - 1), wsPage2.Cells(byLigListePage2 + byNbEleves, 2 * byClasse - 1)).Value
            .HorizontalAlignment = xlHAlignLeft
            .Borders(xlInsideHorizontal).Weight = xlThin
            .Borders.ColorIndex = xlColorIndexAutomatic
        End With

        ' *** EN-TETE ***
        arrChoixCompet = getArrayChoixCompetences
        arrDomaines = getArrayDomaines
        ReDim arrTampon(1 To 3, 1 To 45)
        byColArray = 1
        
        For byDomaine = 1 To 9
        
            ' *** AJOUT TRIMESTRE ET ANNEE ***
            If arrTampon(3, byColArray) = vbNullString Then
                arrTampon(3, byColArray) = "1e tri"
                arrTampon(3, byColArray + 1) = "2e tri"
                arrTampon(3, byColArray + 2) = "3e tri"
                arrTampon(3, byColArray + 3) = "Année"
                
                .Range(.Cells(2, byColArray + 1), .Cells(2, byColArray + 4)).Merge
                
                With .Range(.Cells(2, byColArray + 1), .Cells(byLigListePage4 + byNbEleves, byColArray + 4))
                    .Borders.ColorIndex = xlColorIndexAutomatic
                    .Borders.LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).Weight = xlMedium
                    .Borders(xlEdgeTop).Weight = xlMedium
                    .Borders(xlEdgeLeft).Weight = xlMedium
                    .Borders(xlEdgeRight).Weight = xlMedium
                    .Borders(xlInsideHorizontal).Weight = xlThin
                    .Borders(xlInsideVertical).Weight = xlHairline
                End With
            End If
            
            ' *** AJOUT MOYENNE ***
            If byDomaine = 9 Then
                .Range(.Cells(2, byColArray + 1), .Cells(2, byColArray + 4)).Interior.ColorIndex = byCouleurNote_1
                .Range(.Cells(byLigListePage4, byColArray + 4), .Cells(byLigListePage4 + byNbEleves, byColArray + 4)).Interior.ColorIndex = byCouleurNote_2
                arrTampon(2, byColArray) = "Moyenne"
                byColArray = byColArray + 3
            ' *** AJOUT DOMAINE CHOISI ***
            ElseIf getTailleArray(arrChoixCompet(byDomaine)) > 0 Then
                .Range(.Cells(2, byColArray + 1), .Cells(2, byColArray + 4)).Interior.ColorIndex = byCouleurCompet_1
                .Range(.Cells(byLigListePage4, byColArray + 4), .Cells(byLigListePage4 + byNbEleves, byColArray + 4)).Interior.ColorIndex = byCouleurCompet_2
                arrTampon(2, byColArray) = arrDomaines(byDomaine, 2)
                byColArray = byColArray + 4
            End If
        Next byDomaine
        
        With .Range(.Cells(1, 2), .Cells(1, byColArray + 1))
            .Merge
            .Interior.ColorIndex = byCouleurBilan
            arrTampon(1, 1) = "Bilan trimestriel et annuel"
        End With
        
        ReDim Preserve arrTampon(1 To 3, 1 To byColArray)
        .Range(.Cells(1, 2), .Cells(3, byColArray + 1)).Value = arrTampon
        
        .Range(.Cells(1, 2), .Cells(byLigListePage4 + byNbEleves, byColArray + 1)).BorderAround xlDouble, xlThin, xlColorIndexAutomatic
        .Range(.Cells(byLigListePage4 + 1, 1), .Cells(byLigListePage4 + byNbEleves, byColArray + 1)).BorderAround xlDouble, xlThin, xlColorIndexAutomatic
        .Range(.Cells(byLigListePage4 + 1, 2), .Cells(byLigListePage4 + byNbEleves, byColArray + 1)).Cells.Locked = False
        .Cells(byLigListePage4 + 1, 2).Select
    End With

    ' *** FIGEAGE VOLETS ***
    freezePanes ActiveWindow, byLigListePage4, 1
End Sub

' *******************************************************************************
'                               Fonctions privées
' *******************************************************************************

' *******************************************************************************
'                               Procédures privées
' *******************************************************************************

Sub btnActualiserResultats_Click()
    ' *** DECLARATION VARIABLES ***
    Dim byClasse            As Byte
    Dim byDomaine           As Byte
    Dim byNbDomaines        As Byte
    Dim byTrimestre         As Byte
    Dim byEval              As Byte
    Dim byNbEvals           As Byte
    Dim sPage4              As String
    Dim byAvancementTotal   As Byte
    Dim byAvancementActuel  As Byte

    ' *** AFFECTATION VARIABLES ***
    sPage4 = ActiveSheet.Name
    byClasse = getIndiceClasse(sPage4)
    byNbEvals = getNombreEvals(byClasse)
    
    iIndiceClasse = getIndiceClasse(ActiveSheet.Name)
    iNbDomaines = getNombreDomaines
    iNbEvals = getNombreEvals(intIndiceClasse)
    
    ' *** USERFORM 5 - AFFICHAGE AVANCEMENT ***
    iAvancementActuel = 0
    iAvancementTotal = iNbEvals + 4 * iNbDomaines
    UserForm5.Show vbModeless

    ' *** UPDATES OFF ***
    disableUpdates

    ' *** RECALCUL NOTES EVAL ***
    For iIndiceEval = 1 To iNbEvals
        calculNote iIndiceClasse, iIndiceEval
        iAvancementActuel = iAvancementActuel + 1
        UserForm5.updateAvancement iAvancementActuel, iAvancementTotal
    Next iIndiceEval
    
    ' *** CALCUL MOYENNES PAR DOMAINE ET PAR TRIMESTRE ***
    For iIndiceTrimestre = 1 To 4
        For iIndiceDomaine = 1 To iNbDomaines
            calculMoyenneDomaine iIndiceClasse, iIndiceDomaine, iIndiceTrimestre
            iAvancementActuel = iAvancementActuel + 1
            UserForm5.updateAvancement iAvancementActuel, iAvancementTotal
        Next iIndiceDomaine
        calculMoyenneGlobale iIndiceClasse, iIndiceTrimestre
    Next iIndiceTrimestre
    
    ' *** PROTECTION + REFRESH ECRAN ON ***
    UserForm5.Hide
    protectWorksheet iIndiceClasse
    Application.ScreenUpdating = True
    
    ' *** MESSAGE INFORMATION ***
    Worksheets(sPage4).Activate
    MsgBox ("Données mises à jour.")
End Sub

