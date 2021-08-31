Attribute VB_Name = "Module3"

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
'                               Module 3 - Notes
' *******************************************************************************
'
'   Fonctions publiques
'       ConvertirLettreEnValeur(ByVal sLettre As String) As Byte
'       ConvertirValeurEnLettre(ByVal dbValeur As Double) As String
'       GetNombreEvals(ByVal byClasse As Byte) As Byte
'
'   Procédures publiques
'       InitPage3(ByVal byClasse As Byte, ByVal byNbEleves As Byte)
'       CalculNote(ByVal byClasse As Byte, ByVal byEval As Byte)
'
'   Fonctions privées
'
'   Procédures privées
'       BtnAjouterEvaluation_Click()
'       AjouterEvaluation(ByVal byClasse As Byte, ByVal byEval As Byte, ByVal byColEval As Byte, ByRef arrCompetEval As Variant)
'       BtnCalculNote_Click()
'       BtnInfosNotation_Click()
'
' *******************************************************************************

Option Explicit

' *******************************************************************************
'                               Fonctions publiques
' *******************************************************************************

Public Function ConvertirLettreEnValeur(ByVal sLettre As String) As Byte
    Dim byAsciiLettre As Byte
    byAsciiLettre = Asc(sLettre)
    If byAsciiLettre > 64 And byAsciiLettre < 70 Then
        ConvertirLettreEnValeur = 69 - byAsciiLettre
    Else
        ConvertirLettreEnValeur = 0
    End If
End Function

Public Function ConvertirValeurEnLettre(ByVal dbValeur As Double) As String
    If dbValeur >= 0 And dbValeur <= 4 Then
        Select Case dbValeur
        Case Is > dblNoteA_Min
            ConvertirValeurEnLettre = "A"
        Case Is > dblNoteB_Min
            ConvertirValeurEnLettre = "B"
        Case Is > dblNoteC_Min
            ConvertirValeurEnLettre = "C"
        Case Is > dblNoteD_Min
            ConvertirValeurEnLettre = "D"
        Case Is = 0
            ConvertirValeurEnLettre = "E"
        End Select
    Else
        ConvertirValeurEnLettre = "Z"
    End If
End Function

Public Function GetNombreEvals(ByVal byClasse As Byte) As Byte
    GetNombreEvals = ThisWorkbook.Worksheets(GetNomPage3(byClasse)).Buttons.Count - 2
End Function

Public Function GetArrayEvals(ByVal byClasse As Byte) As Variant()
    ' *** VARIABLES ***
    Dim byNbEvals As Byte
    Dim byEval As Byte
    Dim byCompet As Long
    Dim arrEvals As Variant
    Dim arrCompetEval As Variant
    
    ' *** AFFECTATION VARIABLES ***
    byNbEvals = GetNombreEvals(byClasse)
    ReDim arrEvals(1 To byNbEvals, 1 To 6)
    
    ' *** CALCUL VALEURS ***
    With ThisWorkbook.Worksheets(GetNomPage3(byClasse))
        For byEval = 1 To byNbEvals
            ' *** COLONNE DEBUT EVAL ***
            If Not (byEval = 1) Then
                arrEvals(byEval, 4) = .Buttons("BtnCalculNote_Classe" & byClasse & "_Eval" & byEval - 1).TopLeftCell.Column + 1
            Else
                arrEvals(byEval, 4) = 3
            End If
            
            ' *** NOM / TRIMESTRE / COEFF / NB COMPETS ***
            arrEvals(byEval, 1) = .Cells(1, arrEvals(byEval, 4)).Value
            arrEvals(byEval, 2) = .Cells(2, arrEvals(byEval, 4)).Value
            arrEvals(byEval, 3) = .Cells(3, arrEvals(byEval, 4)).Value
            arrEvals(byEval, 5) = .Buttons("BtnCalculNote_Classe" & byClasse & "_Eval" & byEval).TopLeftCell.Column - arrEvals(byEval, 4)
            
            ' *** ARRAY COMPET ***
            ReDim arrCompetEval(1 To arrEvals(byEval, 5))
            For byCompet = 1 To arrEvals(byEval, 5)
                arrCompetEval(byCompet) = .Cells(5, arrEvals(byEval, 4) + byCompet - 1)
            Next byCompet
            arrEvals(byEval, 6) = arrCompetEval
        Next byEval
    End With
    
    ' *** RENVOI VALEUR ***
    GetArrayEvals = arrEvals
End Function

Public Function GetNombreCompetsEval(ByVal byClasse As Byte, ByVal byEval As Byte) As Byte
    ' *** VARIABLES ***
    Dim lColDebutEval As Long
    Dim lColFinEval As Long
    
    ' *** CALCUL COL DEBUT ET FIN ***
    If Not (byEval = 1) Then
        lColDebutEval = .Buttons("BtnCalculNote_Classe" & byClasse & "_Eval" & byEval - 1).TopLeftCell.Column + 1
    Else
        lColDebutEval = 3
    End If
    lColFinEval = .Buttons("BtnCalculNote_Classe" & byClasse & "_Eval" & byEval).TopLeftCell.Column
    
    ' *** RENVOI VALEUR ***
    GetNombreCompetsEval = lColFinEval - lColDebutEval
End Function

' *******************************************************************************
'                               Procédures publiques
' *******************************************************************************

Public Sub InitPage3(ByVal byClasse As Byte, ByVal byNbEleves As Byte)
    ' *** DECLARATION VARIABLES ***
    Dim byEleve                                     As Byte
    Dim rngBtnAjouterEval                           As Range
    Dim rngBtnInfosNotation                         As Range
    Dim arrHeaderEval(1 To byLigListePage3, 1 To 1) As String
    Dim wsPage3                                     As Worksheet
    Dim arrCompetEval                               As Variant
    
    Set wsPage3 = ThisWorkbook.Worksheets(GetNomPage3(byClasse))
    
    With wsPage3
        ' *** FORMATAGE TAILLE LIGNES + COLONNES ***
        .Rows.RowHeight = 15
        .Range(.Cells(1, 1), .Cells(byLigListePage3, 1)).EntireRow.RowHeight = 25
        .Rows(5).RowHeight = 45
        .Columns.ColumnWidth = 5
        Union(.Columns(1), .Columns(2)).ColumnWidth = 25
        
        ' *** BOUTON 'AJOUTER EVAL' ***
        Set rngBtnAjouterEval = .Range("A1:A2")
        With .Buttons.Add(rngBtnAjouterEval.Left, rngBtnAjouterEval.Top, _
                          rngBtnAjouterEval.Width, rngBtnAjouterEval.Height)
            .Caption = "Gérer les évaluations"
            .OnAction = "BtnGererEvaluation_Click"
            .Name = "BtnGererEval"
        End With
        Set rngBtnAjouterEval = Nothing
        
        ' *** BOUTON 'INFOS NOTATION' ***
        Set rngBtnInfosNotation = .Range("A3:A4")
        With .Buttons.Add(rngBtnInfosNotation.Left, rngBtnInfosNotation.Top, _
                          rngBtnInfosNotation.Width, rngBtnInfosNotation.Height)
            .Caption = "Informations notation"
            .OnAction = "BtnInfosNotation_Click"
            .Name = "BtnInfosNotation"
        End With
        Set rngBtnInfosNotation = Nothing
        
        ' *** LEGENDE ***
        With .Range("A5:A6")
            .Interior.ColorIndex = byCouleurClasse
            .Borders.ColorIndex = xlColorIndexAutomatic
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlMedium
            .MergeCells = True
            .Value = GetNomClasse(byClasse)
        End With
        arrHeaderEval(1, 1) = "Nom de l'évaluation"
        arrHeaderEval(2, 1) = "Trimestre"
        arrHeaderEval(3, 1) = "Coefficient évaluation"
        arrHeaderEval(4, 1) = "Domaine"
        arrHeaderEval(5, 1) = "Compétence"
        arrHeaderEval(6, 1) = "Coefficient compétence"
        With .Range("B1:B3,B4:B6")
            .Borders.ColorIndex = xlColorIndexAutomatic
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders(xlEdgeLeft).Weight = xlMedium
            .Borders(xlEdgeRight).Weight = xlMedium
            .Borders(xlEdgeTop).Weight = xlMedium
            .Borders(xlEdgeBottom).Weight = xlMedium
        End With
        .Range("B1:B6").Value = arrHeaderEval
        .Range("B1").Interior.ColorIndex = byCouleurEval_1
        .Range("B2").Interior.ColorIndex = byCouleurEval_2
        .Range("B4").Interior.ColorIndex = byCouleurCompet_1
        .Range("B5").Interior.ColorIndex = byCouleurCompet_2
        
        ' *** LISTE ELEVES ***
        For byEleve = 1 To byNbEleves
            .Cells(byLigListePage3 + byEleve, 1).Value = ThisWorkbook.Worksheets(strPage2).Cells(byLigListePage2 + byEleve, 2 * byClasse - 1).Value
            .Range(.Cells(byLigListePage3 + byEleve, 1), .Cells(byLigListePage3 + byEleve, 2)).MergeCells = True
        Next byEleve
        With .Range(.Cells(byLigListePage3 + 1, 1), .Cells(byLigListePage3 + byNbEleves, 2))
            .HorizontalAlignment = xlHAlignLeft
            .Borders.ColorIndex = xlColorIndexAutomatic
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
    End With
    
    ' *** FIGEAGE VOLETS ***
    FreezePanes ActiveWindow, byLigListePage3, 2
End Sub

Public Sub CalculNote(ByVal byClasse As Byte, ByVal byEval As Byte)
    Dim byColNote       As Byte         ' Colonne position note
    Dim byColEval       As Byte         ' Colonne début évaluation
    Dim arrCoeff        As Variant      ' Array des coeff compétences
    Dim arrLettres      As Variant      ' Array des lettres de compétence (A/B/C/D/E)
    Dim arrNotes        As Variant      ' Array de sortie des notes (moyenne classe + note pour chaque élèves)
    Dim byEleve         As Byte         ' Index de l'élève
    Dim byNbEleves      As Byte         ' Nombre d'élèves de la classe
    Dim byCompet        As Byte         ' Index de la compétence évaluée
    Dim byNbCompetEval  As Byte         ' Nombre de compétences notées pendant l'évaluation
                                        ' (variable à supprimer après choix des compétences de l'éval par UserForm)
    Dim dbSommeCoeff    As Double       ' Somme des coefficients de chaque compétence
    
    With ThisWorkbook.Worksheets(GetNomPage3(byClasse))
        ' *** COL DEBUT/FIN EVAL ***
        byColNote = .Buttons("BtnCalculNote_Classe" & byClasse & "_Eval" & byEval).TopLeftCell.Column
        If byEval = 1 Then
            byColEval = 3
        Else
            byColEval = .Buttons("BtnCalculNote_Classe" & byClasse & "_Eval" & byEval - 1).TopLeftCell.Column + 1
        End If
        
        ' *** RECUPERATION ARRAYS ***
        arrCoeff = .Range(.Cells(byLigListePage3, byColEval), .Cells(byLigListePage3, byColNote - 1)).Value2
        byNbEleves = GetNombreEleves(byClasse)
        arrLettres = .Range(.Cells(byLigListePage3 + 1, byColEval), .Cells(byLigListePage3 + byNbEleves, byColNote - 1))
        ReDim arrNotes(0 To byNbEleves)
        byNbCompetEval = 0
        dbSommeCoeff = 0#
        arrNotes(0) = 0#
        
        ' *** CALCUL NOMBRE COMPET EVAL ***
        For byCompet = 1 To byColNote - byColEval
            If Not (IsEmpty(arrCoeff(1, byCompet))) And IsNumeric(arrCoeff(1, byCompet)) Then
                byNbCompetEval = byNbCompetEval + 1
                dbSommeCoeff = dbSommeCoeff + arrCoeff(1, byCompet)
            End If
        Next byCompet
        
        ' *** PARCOURS DES ELEVES ***
        If byNbCompetEval > 0 Then
            For byEleve = 1 To byNbEleves
                arrNotes(byEleve) = 0#
                ' *** PARCOURS DES COMPETENCES ***
                For byCompet = 1 To byColNote - byColEval
                    ' *** AJOUT DE LA COMPETENCE SI COEFF <> 0
                    If Not (IsEmpty(arrCoeff(1, byCompet))) And IsNumeric(arrCoeff(1, byCompet)) Then
                        arrNotes(byEleve) = arrNotes(byEleve) + CDbl(ConvertirLettreEnValeur(CStr(arrLettres(byEleve, byCompet)))) * CDbl(arrCoeff(1, byCompet))
                    End If
                Next byCompet
                ' *** CALCUL NOTE FINALE ***
                arrNotes(byEleve) = Format(5# * arrNotes(byEleve) / CDbl(dbSommeCoeff), "Standard")
                arrNotes(0) = arrNotes(0) + arrNotes(byEleve)
                ' *** ACTUALISATION AVANCEMENT ***
                UserForm6.updateAvancement byEleve, byNbEleves
            Next byEleve
            ' *** CALCUL MOYENNE CLASSE ***
            arrNotes(0) = Format(arrNotes(0) / CDbl(byNbEleves), "Standard")
            ' *** ECRITURE DES CELLULES ***
            .Range(.Cells(byLigListePage3, byColNote), .Cells(byLigListePage3 + byNbEleves, byColNote)) = Application.WorksheetFunction.Transpose(arrNotes)
        End If
    End With
End Sub

Public Sub AjouterEvaluation(ByVal byClasse As Byte, ByRef arrCompetEval As Variant, Optional ByVal strEval As Variant, Optional ByVal dbCoeffEval As Variant, Optional ByVal byTrimestre As Variant)
    ' *** DECLARATION VARIABLES ***
    Dim byEval As Byte
    Dim lEval_Col As Long
    Dim arrDomaines As Variant
    Dim byDomaine As Byte
    Dim byCompet As Byte
    Dim byNbCompet As Byte
    Dim byNbCompetParDomaine As Byte
    Dim byCompetGeneral As Byte
    Dim byNbEleves As Byte
    Dim rngBtnCalculNote As Range
    
    If Not IsArray(arrCompetEval) Then
        Call DisplayTemporaryMessage("Erreur Module3 'AjouterEvaluation': argument arrCompetEval incorrect")
        Exit Sub
    End If
    
    With ThisWorkbook.Worksheets(GetNomPage3(byClasse))
        ' *** AFFECTATION VARIABLES ***
        arrDomaines = GetArrayDomaines
        byNbCompet = GetSizeOfJaggedArray(arrCompetEval) / 2
        byNbEleves = GetNombreEleves(byClasse)
        byEval = .Buttons.Count - 1
        If byEval <> 1 Then
            lEval_Col = .UsedRange.Columns.Count
        Else
            lEval_Col = 3
        End If
        
        ' *** FORMATAGE ZONE EVAL ***
        .Range(.Cells(1, lEval_Col), .Cells(1, lEval_Col + byNbCompet - 1)).Interior.ColorIndex = byCouleurEval_1
        .Range(.Cells(1, lEval_Col), .Cells(1, lEval_Col + byNbCompet - 1)).MergeCells = True
        If Not IsMissing(strEval) Then .Range(.Cells(1, lEval_Col), .Cells(1, lEval_Col + byNbCompet - 1)).Value = strEval
        .Range(.Cells(2, lEval_Col), .Cells(2, lEval_Col + byNbCompet - 1)).Interior.ColorIndex = byCouleurEval_2
        .Range(.Cells(2, lEval_Col), .Cells(2, lEval_Col + byNbCompet - 1)).MergeCells = True
        If Not IsMissing(byTrimestre) Then .Range(.Cells(2, lEval_Col), .Cells(2, lEval_Col + byNbCompet - 1)).Value = byTrimestre
        .Range(.Cells(3, lEval_Col), .Cells(3, lEval_Col + byNbCompet - 1)).MergeCells = True
        If Not IsMissing(dbCoeffEval) Then .Range(.Cells(3, lEval_Col), .Cells(3, lEval_Col + byNbCompet - 1)).Value = dbCoeffEval
        With .Range(.Cells(1, lEval_Col), _
                    .Cells(3, lEval_Col + byNbCompet - 1))
            .Borders.ColorIndex = xlColorIndexAutomatic
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders(xlEdgeBottom).Weight = xlMedium
            .Borders(xlEdgeTop).Weight = xlMedium
            .Borders(xlEdgeLeft).Weight = xlMedium
            .Borders(xlEdgeRight).Weight = xlMedium
            .Columns.ColumnWidth = 3
            .Locked = False
        End With
        
        ' *** FORMATAGE ZONE COMPET ***
        byCompetGeneral = 0
        .Range(.Cells(4, lEval_Col), .Cells(4, lEval_Col + byNbCompet - 1)).Interior.ColorIndex = byCouleurCompet_1
        .Range(.Cells(5, lEval_Col), .Cells(5, lEval_Col + byNbCompet - 1)).Interior.ColorIndex = byCouleurCompet_2
        .Range(.Cells(5, lEval_Col), .Cells(5, lEval_Col + byNbCompet - 1)).Orientation = xlUpward
        For byDomaine = 1 To 8
            byNbCompetParDomaine = GetSizeOfArray(arrCompetEval(byDomaine)) / 2
            If byNbCompetParDomaine <> 0 Then
                With .Range(.Cells(4, lEval_Col + byCompetGeneral), _
                            .Cells(6 + byNbEleves, lEval_Col + byCompetGeneral + (byNbCompetParDomaine - 1)))
                    .Borders.ColorIndex = xlColorIndexAutomatic
                    .Borders.LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).Weight = xlMedium
                    .Borders(xlEdgeTop).Weight = xlMedium
                    .Borders(xlEdgeLeft).Weight = xlMedium
                    .Borders(xlEdgeRight).Weight = xlMedium
                    .Borders(xlInsideHorizontal).Weight = xlThin
                    .Borders(xlInsideVertical).Weight = xlHairline
                End With
                With .Range(.Cells(4, lEval_Col + byCompetGeneral), _
                            .Cells(4, lEval_Col + byCompetGeneral + (byNbCompetParDomaine - 1)))
                    .MergeCells = True
                    .Value = arrDomaines(byDomaine, 2)
                End With
                If byNbCompetParDomaine <> 1 Then
                    For byCompet = 1 To byNbCompetParDomaine
                        byCompetGeneral = byCompetGeneral + 1
                        .Cells(5, lEval_Col + (byCompetGeneral - 1)).Value = arrCompetEval(byDomaine)(byCompet, 1)
                    Next byCompet
                ElseIf byNbCompetParDomaine = 1 Then
                    byCompetGeneral = byCompetGeneral + 1
                    .Cells(5, lEval_Col + (byCompetGeneral - 1)).Value = arrCompetEval(byDomaine)(1)
                End If
            End If
        Next byDomaine
        
        ' *** FORMATAGE ZONE NOTE ***
        Set rngBtnCalculNote = .Range(.Cells(1, lEval_Col + byNbCompet), .Cells(3, lEval_Col + byNbCompet))
        With .Buttons.Add(rngBtnCalculNote.Left, rngBtnCalculNote.Top, _
                          rngBtnCalculNote.Width, rngBtnCalculNote.Height)
            .Caption = "Calcul" & vbNewLine & "note"
            .OnAction = "BtnCalculNote_Click"
            .Name = "BtnCalculNote_Classe" & byClasse & "_Eval" & byEval
            .Locked = True
            .LockedText = True
        End With
        rngBtnCalculNote.Locked = True
        With .Range(.Cells(4, lEval_Col + byNbCompet), .Cells(byLigListePage3 + byNbEleves, lEval_Col + byNbCompet))
            .Interior.ColorIndex = byCouleurNote_2
            .Borders.ColorIndex = xlColorIndexAutomatic
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders(xlEdgeBottom).Weight = xlMedium
            .Borders(xlEdgeTop).Weight = xlMedium
            .Borders(xlEdgeLeft).Weight = xlMedium
            .Borders(xlEdgeRight).Weight = xlMedium
            .Borders(xlInsideHorizontal).Weight = xlThin
        End With
        With .Range(.Cells(4, lEval_Col + byNbCompet), _
                    .Cells(5, lEval_Col + byNbCompet))
            .Interior.ColorIndex = byCouleurNote_1
            .MergeCells = True
            .Orientation = xlUpward
            .Value = "Note / 20"
            .Columns.ColumnWidth = 6
        End With
        
        ' *** FORMATAGE ENCADREMENT ***
        .Range(.Cells(1, lEval_Col), .Cells(byLigListePage3 + byNbEleves, lEval_Col + byNbCompet)).BorderAround xlDouble, xlThin, xlColorIndexAutomatic
        .Range(.Cells(byLigListePage3 + 1, 1), .Cells(byLigListePage3 + byNbEleves, lEval_Col + byNbCompet)).BorderAround xlDouble, xlThin, xlColorIndexAutomatic
        .Range(.Cells(byLigListePage3, lEval_Col), .Cells(byLigListePage3 + byNbEleves, lEval_Col + byNbCompet)).Locked = False
    End With
End Sub

Public Sub ModifierEval(ByVal byClasse As Byte, ByVal byEval As Byte, ByVal sEval As String, ByVal byTrimestre As Byte, ByVal dbCoeff As Double, ByVal arrCompetEval As Variant)
    
End Sub

Public Sub SupprimerEval(ByVal byClasse As Byte, ByVal byEvalSuppr As Byte)
    ' *** VARIABLES ***
    Dim arrEvals As Variant
    Dim byNbEvals As Byte
    Dim byEval As Byte
    Dim rngButton As Variant
    
    ' *** AFFECTATION VARIABLES ***
    arrEvals = GetArrayEvals(byClasse)
    byNbEvals = GetSizeOfArray(arrEvals) / 6
    
    With ThisWorkbook.Worksheets(GetNomPage3(byClasse))
        ' *** SUPPRESSION EVAL ***
        .Buttons("BtnCalculNote_Classe" & byClasse & "_Eval" & byEvalSuppr).Delete
        .Range(.Cells(1, arrEvals(byEvalSuppr, 4)), .Cells(1, arrEvals(byEvalSuppr, 4) + arrEvals(byEvalSuppr, 5))).EntireColumn.Delete
        
        ' *** MODIFICATION NOM BOUTON ***
        If byEvalSuppr < byNbEvals Then
            For byEval = byEvalSuppr + 1 To byNbEvals
                .Buttons("BtnCalculNote_Classe" & byClasse & "_Eval" & byEval).Name = "BtnCalculNote_Classe" & byClasse & "_Eval" & byEval - 1
                .Buttons("BtnCalculNote_Classe" & byClasse & "_Eval" & byEval - 1).Locked = True
                .Buttons("BtnCalculNote_Classe" & byClasse & "_Eval" & byEval - 1).LockedText = True
            Next byEval
        End If
    End With
End Sub

' *******************************************************************************
'                               Fonctions privées
' *******************************************************************************

' *******************************************************************************
'                               Procédures privées
' *******************************************************************************

'@EntryPoint
Private Sub BtnGererEvaluation_Click()
    ' *** VARIABLES ***
    Dim byClasse As Byte

    ' *** CALCUL ***
    byClasse = GetIndiceClasse(ActiveSheet.Name)
    Call DisableUpdates
    Call UserForm1.SetUp(byClasse)
    Call UserForm1.Show(vbModeless)
    Call EnableUpdates
    
    ' *** MESSAGE INFORMATION ***
    Call DisplayTemporaryMessage("Evaluation ajoutée à la classe '" & GetNomClasse(byClasse) & "'.", 10)
End Sub

'@EntryPoint
'Private Sub BtnAjouterEvaluation_Click()
'    ' *** DECLARATION VARIABLES ***
'    Dim byClasse As Byte
'    Dim byColEval As Byte
'    Dim byEval As Byte
'    Dim arrCompetEval As Variant
'
'    ' *** AFFECTATION VARIABLES ***
'    byClasse = GetIndiceClasse(ActiveSheet.Name)
'    With ThisWorkbook.Worksheets(GetNomPage3(byClasse))
'        byEval = .Buttons.Count - 1
'        byColEval = .UsedRange.Columns.Count
'    End With
'    arrCompetEval = GetArrayChoixCompetences
'
'    ' *** AJOUT EVAL ***
'    DisableUpdates
'    AjouterEvaluation byClasse, byEval, byColEval, arrCompetEval
'    EnableUpdates
'
'    ' *** MESSAGE INFORMATION ***
'    MsgBox "Nouvelle évaluation ajoutée à la classe '" & GetNomClasse(byClasse) & "'."
'End Sub

'@EntryPoint
Private Sub BtnInfosNotation_Click()
    MsgBox "Infos notation ..."
End Sub

'@EntryPoint
Private Sub BtnCalculNote_Click()
    Dim sEval       As String   ' Nom de l'évaluation
    Dim byEval      As Byte     ' Numéro de l'évaluation
    Dim byClasse    As Byte     ' Numéro de la classe

    ' *** AFFECTATION VARIABLES ***
    sEval = Split(Application.Caller, "_")(2)
    byEval = CByte(Right(sEval, Len(sEval) - 4))
    byClasse = GetIndiceClasse(ActiveSheet.Name)
    
    ' *** REFRESH ECRAN OFF ***
    ' UserForm6.Show vbModeless
    DisableUpdates
    
    ' *** CALCUL NOTE ***
    CalculNote byClasse, byEval
        
    ' *** REFRESH ECRAN ON ***
    EnableUpdates
    Unload UserForm6
End Sub


