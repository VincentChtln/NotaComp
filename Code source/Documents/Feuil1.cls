VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil1"
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
'                               Page 1 - Accueil
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

Private Sub Worksheet_Change(ByVal Target As Range)
    ' *** DECLARATION VARIABLES ***
    Dim arrCelDomaine() As String
    Dim cel As Range
    
    ' *** EVENTS + REFRESH OFF ***
    DisableUpdates

    With ThisWorkbook.Worksheets(strPage1)
        For Each cel In Target
            ' *** GESTION DE CHANGEMENT DE COMPETENCE PAGE 1 ***
            If (cel.Row > byLigTabClasses + GetNombreClasses + 7) Then
                Select Case cel.Column
                    ' *** SELECTION COMPETENCE --> AJOUT ABREVIATION PAR DEFAUT ***
                Case byColTabCompet + 2
                    If cel.Value = "x" Or cel.Value = "X" Then
                        arrCelDomaine = Split(.Cells(cel.Row, byColTabCompet).MergeArea.Cells(1, 1).Value, ":")
                        Select Case GetSizeOfArray(arrCelDomaine)
                        Case 2
                            .Cells(cel.Row, cel.Column + 1).Value = "D" & Right(arrCelDomaine(LBound(arrCelDomaine)), 1) & "/" & _
                                                                                                                         Mid(.Cells(cel.Row, byColTabCompet + 1).Value, 2, 1)
                        Case 3
                            .Cells(cel.Row, cel.Column + 1).Value = "D" & Right(arrCelDomaine(LBound(arrCelDomaine)), 1) & "-" & _
                                                                                                                         Right(arrCelDomaine(LBound(arrCelDomaine) + 1), 1) & "/" & _
                                                                                                                         Mid(.Cells(cel.Row, byColTabCompet + 1).Value, 2, 1)
                        End Select
                    Else
                        .Cells(cel.Row, cel.Column + 1).Value = vbNullString
                    End If
                    ' *** AJOUT ABREVIATION ---> SELECTION COMPETENCE
                Case byColTabCompet + 3
                    If IsEmpty(cel.Value) Then
                        .Cells(cel.Row, cel.Column - 1).Value = vbNullString
                    Else
                        .Cells(cel.Row, cel.Column - 1).Value = "X"
                    End If
                End Select
            End If
        Next cel
    End With

    ' *** EVENTS + REFRESH ON ***
    EnableUpdates
End Sub


