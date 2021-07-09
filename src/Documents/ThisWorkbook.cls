VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

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
'                               ThisWorkbook
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

Private Sub Workbook_Open()
    ProtectWorkbook
    ProtectAllWorksheets
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal ws As Object, ByVal Target As Excel.Range)
    ' *** DECLARATION VARIABLES ***
    Dim cel As Range
    
    ' *** GESTION SELECTION PAGE 1 ***
    If (InStr(ws.Name, strPage1) <> 0) Then
        DisableUpdates

        With ws.UsedRange.Cells
            .Font.Bold = False
        End With
    
        ' *** MODIF SI SELECTION DANS TAB COMPET ***
        If (Target.Row > byLigTabClasses + GetNombreClasses + 7) Then
            Target.EntireRow.Font.Bold = True
            For Each cel In Target
                ws.Cells(cel.Row, byColTabCompet).MergeArea.Cells(1, 1).Font.Bold = True
            Next cel
        End If
            
        EnableUpdates
    End If

    ' *** GESTION SELECTION PAGE 3 & 4 ***
    If (InStr(ws.Name, "Notes") <> 0 And Target.Row > byLigListePage3) Or (InStr(ws.Name, "Bilan") <> 0 And Target.Row > byLigListePage4) Then
        DisableUpdates
        
        ' *** REFORMATAGE FEUILLE ENTIERE ***
        With ws.Cells
            .Font.Bold = False
            .Font.Size = 11
        End With

        ' *** FORMATAGE LIGNES SELECTIONNEES ***
        With Target.EntireRow
            .Font.Bold = True
            .Font.Size = 12
        End With
        
        EnableUpdates
    End If

End Sub

