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

Public Function convertLettreEnValeur(sLettre As String) As Byte
    Dim byAsciiLettre As Byte
    byAsciiLettre = Asc(sLettre)
    If byAsciiLettre > 64 And byAsciiLettre < 70 Then
        convertLettreEnValeur = 69 - byAsciiLettre
    Else
        convertLettreEnValeur = 0
    End If
End Function

Public Function convertValeurEnLettre(dbValeur As Double) As String
    If dbValeur >= 0 And dbValeur <= 4 Then
        Select Case iValeur
        Case Is > 3.5
            convertValeurEnLettre = "A"
        Case Is > 2.5
            convertValeurEnLettre = "B"
        Case Is > 1.5
            convertValeurEnLettre = "C"
        Case Is > 0
            convertValeurEnLettre = "D"
        Case Is = 0
            convertValeurEnLettre = "E"
        End Select
    Else
        convertValeurEnLettre = "Z"
    End If
End Function

Function getNombreEvals(byClasse As Byte) As Byte
    getNombreEvals = ThisWorkbook.Worksheets(getNomPage3(byClasse)).Buttons.Count - 1
End Function

' *******************************************************************************
'                               Procédures publiques
' *******************************************************************************

Public Sub initPage3(ByVal byClasse As Byte, ByVal byNbEleves As Byte)
    ' *** DECLARATION VARIABLES ***
    Dim byEleve                                     As Byte
    Dim rngBtnAjouterEval                           As Range
    Dim rngBtnInfosNotation                         As Range
    Dim arrHeaderEval(1 To byLigListePage3, 1 To 1) As String
    Dim wsPage3                                     As Worksheet
    Dim arrCompetEval                               As Variant
    
    Set wsPage3 = ThisWorkbook.Worksheets(getNomPage3(byClasse))
    
    With wsPage3
        ' *** FORMATAGE TAILLE LIGNES + COLONNES ***
        .Rows.RowHeight = 15
        .Range(.Cells(1, 1), .Cells(byLigListePage3, 1)).EntireRow.RowHeight = 25
        .Rows(5).RowHeight = 40
        .Columns.ColumnWidth = 5
        Union(.Columns(1), .Columns(2)).ColumnWidth = 25
        
        ' *** BOUTON 'AJOUTER EVAL' ***
        Set rngBtnAjouterEval = .Range("A1:A2")
        With .Buttons.Add(rngBtnAjouterEval.Left, rngBtnAjouterEval.Top, _
                          rngBtnAjouterEval.Width, rngBtnAjouterEval.Height)
            .Caption = "Ajouter une évaluation"
            .OnAction = "btnAjouterEvaluation_Click"
            .Name = "btnAjouterEval"
        End With
        Set rngBtnAjouterEval = Nothing
        
        ' *** BOUTON 'INFOS NOTATION' ***
        Set rngBtnInfosNotation = .Range("A3:A4")
        With .Buttons.Add(rngBtnInfosNotation.Left, rngBtnInfosNotation.Top, _
                          rngBtnInfosNotation.Width, rngBtnInfosNotation.Height)
            .Caption = "Informations notation"
            .OnAction = "btnInfosNotation_Click"
            .Name = "btnInfosNotation"
        End With
        Set rngBtnInfosNotation = Nothing
        
        ' *** LEGENDE ***
        With .Range("A5:A6")
            .Interior.ColorIndex = byCouleurClasse
            .Borders.ColorIndex = xlColorIndexAutomatic
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlMedium
            .MergeCells = True
            .Value = getNomClasse(byClasse)
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
    freezePanes ActiveWindow, byLigListePage3, 2
    
    ' *** AJOUT 1e EVALUATION ***
    arrCompetEval = getArrayChoixCompetences
    ajouterEvaluation byClasse, 1, 3, arrCompetEval
End Sub

' *******************************************************************************
'                               Fonctions privées
' *******************************************************************************

' *******************************************************************************
'                               Procédures privées
' *******************************************************************************

Private Sub btnAjouterEvaluation_Click()
    ' *** DECLARATION VARIABLES ***
    Dim byClasse As Byte
    Dim byColEval As Byte
    Dim byEval As Byte
    Dim arrCompetEval As Variant
    
    ' *** AFFECTATION VARIABLES ***
    byClasse = getIndiceClasse(ActiveSheet.Name)
    With ThisWorkbook.Worksheets(getNomPage3(byClasse))
        byEval = .Buttons.Count - 1
        byColEval = .UsedRange.Columns.Count
    End With
    arrCompetEval = getArrayChoixCompetences
    
    ' *** AJOUT EVAL ***
    disableUpdates
    ajouterEvaluation byClasse, byEval, byColEval, arrCompetEval
    enableUpdates
    
    ' *** MESSAGE INFORMATION ***
    MsgBox "Nouvelle évaluation ajoutée à la classe '" & getNomClasse(byClasse) & "'."
End Sub

Private Sub ajouterEvaluation(ByVal byClasse As Byte, ByVal byEval As Byte, ByVal byColEval As Byte, ByRef arrCompetEval As Variant)
    ' *** DECLARATION VARIABLES ***
    Dim arrDomaines As Variant
    Dim iDomaine As Byte
    Dim iCompet As Byte
    Dim byNbCompet As Byte
    Dim byNbCompetParDomaine As Byte
    Dim iCompetGeneral As Byte
    Dim byNbEleves As Byte
    Dim rngBtnCalculNote As Range
    
    With ThisWorkbook.Worksheets(getNomPage3(byClasse))
        ' *** AFFECTATION VARIABLES ***
        arrDomaines = getArrayDomaines
        byNbCompet = getTailleJaggedArray(arrCompetEval)
        byNbEleves = getNombreEleves(byClasse)
        
        ' *** FORMATAGE ZONE EVAL ***
        .Range(.Cells(1, byColEval), .Cells(1, byColEval + byNbCompet - 1)).Interior.ColorIndex = byCouleurEval_1
        .Range(.Cells(1, byColEval), .Cells(1, byColEval + byNbCompet - 1)).MergeCells = True
        .Range(.Cells(2, byColEval), .Cells(2, byColEval + byNbCompet - 1)).Interior.ColorIndex = byCouleurEval_2
        .Range(.Cells(2, byColEval), .Cells(2, byColEval + byNbCompet - 1)).MergeCells = True
        .Range(.Cells(3, byColEval), .Cells(3, byColEval + byNbCompet - 1)).MergeCells = True
        With .Range(.Cells(1, byColEval), _
                    .Cells(3, byColEval + byNbCompet - 1))
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
        iCompetGeneral = 0
        .Range(.Cells(4, byColEval), .Cells(4, byColEval + byNbCompet - 1)).Interior.ColorIndex = byCouleurCompet_1
        .Range(.Cells(5, byColEval), .Cells(5, byColEval + byNbCompet - 1)).Interior.ColorIndex = byCouleurCompet_2
        .Range(.Cells(5, byColEval), .Cells(5, byColEval + byNbCompet - 1)).Orientation = xlUpward
        For iDomaine = 1 To 8
            byNbCompetParDomaine = getTailleArray(arrCompetEval(iDomaine))
            If byNbCompetParDomaine <> 0 Then
                With .Range(.Cells(4, byColEval + iCompetGeneral), _
                            .Cells(6 + byNbEleves, byColEval + iCompetGeneral + (byNbCompetParDomaine - 1)))
                    .Borders.ColorIndex = xlColorIndexAutomatic
                    .Borders.LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).Weight = xlMedium
                    .Borders(xlEdgeTop).Weight = xlMedium
                    .Borders(xlEdgeLeft).Weight = xlMedium
                    .Borders(xlEdgeRight).Weight = xlMedium
                    .Borders(xlInsideHorizontal).Weight = xlThin
                    .Borders(xlInsideVertical).Weight = xlHairline
                End With
                With .Range(.Cells(4, byColEval + iCompetGeneral), _
                            .Cells(4, byColEval + iCompetGeneral + (byNbCompetParDomaine - 1)))
                    .MergeCells = True
                    .Value = arrDomaines(iDomaine, 2)
                End With
                For iCompet = 1 To byNbCompetParDomaine
                    iCompetGeneral = iCompetGeneral + 1
                    .Cells(5, byColEval + (iCompetGeneral - 1)).Value = arrCompetEval(iDomaine)(iCompet)
                Next iCompet
            End If
        Next iDomaine
        
        ' *** FORMATAGE ZONE NOTE ***
        Set rngBtnCalculNote = .Range(.Cells(1, byColEval + byNbCompet), .Cells(3, byColEval + byNbCompet))
        With .Buttons.Add(rngBtnCalculNote.Left, rngBtnCalculNote.Top, _
                          rngBtnCalculNote.Width, rngBtnCalculNote.Height)
            .Caption = "Calcul" & vbNewLine & "note"
            .OnAction = "btnCalculNote_Click"
            .Name = "btnCalculNote_Classe" & byClasse & "_Eval" & byEval
        End With
        With .Range(.Cells(4, byColEval + byNbCompet), .Cells(byLigListePage3 + byNbEleves, byColEval + byNbCompet))
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
        With .Range(.Cells(4, byColEval + byNbCompet), _
                    .Cells(5, byColEval + byNbCompet))
            .Interior.ColorIndex = byCouleurNote_1
            .MergeCells = True
            .Orientation = xlUpward
            .Value = "Note / 20"
            .Columns.ColumnWidth = 6
        End With
        
        ' *** FORMATAGE ENCADREMENT ***
        .Range(.Cells(1, byColEval), .Cells(byLigListePage3 + byNbEleves, byColEval + byNbCompet)).BorderAround xlDouble, xlThin, xlColorIndexAutomatic
        .Range(.Cells(byLigListePage3 + 1, 1), .Cells(byLigListePage3 + byNbEleves, byColEval + byNbCompet)).BorderAround xlDouble, xlThin, xlColorIndexAutomatic
        .Range(.Cells(byLigListePage3, byColEval), .Cells(byLigListePage3 + byNbEleves, byColEval + byNbCompet)).Locked = False
    End With
End Sub

Private Sub btnInfosNotation_Click()
    MsgBox "Infos notation ..."
End Sub

Private Sub btnCalculNote_Click()
    Dim sEval       As String   ' Nom de l'évaluation
    Dim byEval      As Byte     ' Numéro de l'évaluation
    Dim byClasse    As Byte     ' Numéro de la classe

    ' *** AFFECTATION VARIABLES ***
    sEval = Split(Application.Caller, "_")(2)
    byEval = CByte(Right(sEval, Len(sEval) - 4))
    byClasse = getIndiceClasse(ActiveSheet.Name)
    
    ' *** REFRESH ECRAN OFF ***
    UserForm5.Show vbModeless
    disableUpdates
    
    ' *** CALCUL NOTE ***
    calculNote byClasse, byEval
        
    ' *** REFRESH ECRAN ON ***
    enableUpdates
    Unload UserForm5
End Sub

Private Sub calculNote(ByVal byClasse As Byte, ByVal byEval As Byte)
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
    
    With ThisWorkbook.Worksheets(getNomPage3(byClasse))
        ' *** COL DEBUT/FIN EVAL ***
        byColNote = .Buttons("btnCalculNote_Classe" & byClasse & "_Eval" & byEval).TopLeftCell.Column
        If byEval = 1 Then
            byColEval = 3
        Else
            byColEval = .Buttons("btnCalculNote_Classe" & byClasse & "_Eval" & byEval - 1).TopLeftCell.Column + 1
        End If
        
        ' *** RECUPERATION ARRAYS ***
        arrCoeff = .Range(.Cells(byLigListePage3, byColEval), .Cells(byLigListePage3, byColNote - 1)).Value2
        byNbEleves = getNombreEleves(byClasse)
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
                        arrNotes(byEleve) = arrNotes(byEleve) + CDbl(convertLettreEnValeur(CStr(arrLettres(byEleve, byCompet)))) * CDbl(arrCoeff(1, byCompet))
                    End If
                Next byCompet
                ' *** CALCUL NOTE FINALE ***
                arrNotes(byEleve) = Format(5# * arrNotes(byEleve) / CDbl(dbSommeCoeff), "Standard")
                arrNotes(0) = arrNotes(0) + arrNotes(byEleve)
                ' *** ACTUALISATION AVANCEMENT ***
                UserForm5.updateAvancement byEleve, byNbEleves
            Next byEleve
            ' *** CALCUL MOYENNE CLASSE ***
            arrNotes(0) = Format(arrNotes(0) / CDbl(byNbEleves), "Standard")
            ' *** ECRITURE DES CELLULES ***
            .Range(.Cells(byLigListePage3, byColNote), .Cells(byLigListePage3 + byNbEleves, byColNote)) = Application.WorksheetFunction.Transpose(arrNotes)
        End If
    End With
End Sub


