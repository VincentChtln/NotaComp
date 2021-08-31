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
'                               Module 4 - Bilan
' *******************************************************************************
'
'   Fonctions publiques
'
'   Procédures publiques
'       InitPage4(byClasse As Byte, byNbEleves As Byte)
'
'   Fonctions privées
'
'   Procédures privées
'       BtnActualiserResultats_Click()
'       CalculMoyenneDomainesEtAnnee(ByVal byClasse As Byte)
'
' *******************************************************************************

Option Explicit

' *******************************************************************************
'                               Fonctions publiques
' *******************************************************************************

' *******************************************************************************
'                               Procédures publiques
' *******************************************************************************

Public Sub InitPage4(byClasse As Byte, byNbEleves As Byte)
    Dim wsPage2                     As Worksheet
    Dim wsPage4                     As Worksheet
    Dim rngBtnActualiserResultats   As Range
    Dim arrDomaines                 As Variant
    Dim arrChoixCompet              As Variant
    Dim arrTampon                   As Variant
    Dim byDomaine                   As Byte
    Dim byColArray                  As Byte
    
    Set wsPage4 = ThisWorkbook.Worksheets(GetNomPage4(byClasse))
    
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
            .OnAction = "BtnActualiserResultats_Click"
            .Name = "BtnActualiserResultats"
        End With
        
        ' *** NOM CLASSE ***
        With .Range("A2")
            .Value = GetNomClasse(byClasse)
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
        arrChoixCompet = GetArrayChoixCompetences()
        arrDomaines = GetArrayDomaines()
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
            ElseIf GetSizeOfArray(arrChoixCompet(byDomaine)) > 0 Then
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
    FreezePanes ActiveWindow, byLigListePage4, 1
End Sub

' *******************************************************************************
'                               Fonctions privées
' *******************************************************************************

' *******************************************************************************
'                               Procédures privées
' *******************************************************************************

'@EntryPoint
Private Sub BtnActualiserResultats_Click()
    ' *** DECLARATION VARIABLES ***
    Dim byClasse            As Byte
    Dim byEval              As Byte
    Dim byNbEvals           As Byte
    Dim sPage4              As String
    Dim byAvancementTotal   As Byte
    Dim byAvancementActuel  As Byte

    ' *** AFFECTATION VARIABLES ***
    sPage4 = ActiveSheet.Name
    byClasse = GetIndiceClasse(sPage4)
    byNbEvals = GetNombreEvals(byClasse)
    
    ' *** VERIFICATION PRESENCE EVALUATIONS ***
    If byNbEvals = 0 Then
        Call DisplayTemporaryMessage("Erreur Module4 'BtnActualiserResultats_Click': Aucune évaluation à calculer.")
        Exit Sub
    End If
    
    ' *** USERFORM 5 - AFFICHAGE AVANCEMENT ***
    byAvancementActuel = 0
    byAvancementTotal = byNbEvals + 1
    Call UserForm6.Show(vbModeless)

    ' *** UPDATES OFF ***
    Call DisableUpdates

    ' *** RECALCUL NOTES EVAL ***
    For byEval = 1 To byNbEvals
        Call CalculNote(byClasse, byEval)
        byAvancementActuel = byAvancementActuel + 1
        Call UserForm6.updateAvancement(byAvancementActuel, byAvancementTotal)
    Next byEval
    
    ' *** CALCUL MOYENNES PAR DOMAINE ET PAR TRIMESTRE ***
    Call CalculMoyenneDomainesEtAnnee(byClasse)
    byAvancementActuel = byAvancementActuel + 1
    Call UserForm6.updateAvancement(byAvancementActuel, byAvancementTotal)
    
    ' *** UPDATES ON ***
    Call UserForm6.Hide
    Call EnableUpdates
    
    ' *** MESSAGE INFORMATION ***
    Call ThisWorkbook.Worksheets(sPage4).Activate
    Call MsgBox("Données mises à jour.")
End Sub

Private Sub CalculMoyenneDomainesEtAnnee(ByVal byClasse As Byte)
    ' *** DECLARATION VARIABLES ***
    Dim byIndiceDomaine     As Byte         '
    Dim byNbDomaines        As Byte         '
    Dim byEleve             As Byte         '
    Dim byNbEleves          As Byte         '
    Dim byTrimestre         As Byte         '
    Dim bEvalOK             As Boolean      '
    Dim dbCoeffEval         As Double       '
    Dim dbCoeffCompet       As Double       '
    Dim arrChoixDomaines    As Variant      '
    Dim arrSource           As Variant      '
    Dim arrDest             As Variant      '
    Dim arrCoeff            As Variant      '
    Dim lColSource          As Long         '
    Dim lColDest            As Long         '
    
    ' *** AFFECTATION VARIABLES ***
    byNbEleves = GetNombreEleves(byClasse)
    With ThisWorkbook.Worksheets(GetNomPage3(byClasse))
        arrSource = .Range(.Cells(1, 3), .Cells(byLigListePage3 + byNbEleves, .UsedRange.Columns.Count - 1)).Value
    End With
    arrChoixDomaines = GetArrayChoixDomaines
    byNbDomaines = GetSizeOfArray(arrChoixDomaines)
    ReDim arrCoeff(1 To 4 * (byNbDomaines + 1))
    For lColSource = 1 To UBound(arrCoeff, 1)
        arrCoeff(lColSource) = 0#
    Next lColSource
    ReDim arrDest(1 To byNbEleves, 1 To 4 * (byNbDomaines + 1))
    bEvalOK = False
    
    ' *** CALCUL SOMME NOTES & COEFFS ***
    For lColSource = LBound(arrSource, 2) To UBound(arrSource, 2)
    
        ' *** VERIFICATION EVAL OK ***
        If ((arrSource(2, lColSource) <> vbNullString) And (arrSource(3, lColSource) <> vbNullString)) Then
            If IsNumeric(CInt(arrSource(2, lColSource))) And IsNumeric(CDbl(arrSource(3, lColSource))) Then
                byTrimestre = CInt(arrSource(2, lColSource))
                If (byTrimestre > 0) And (byTrimestre < 4) Then
                    bEvalOK = True
                    dbCoeffEval = CDbl(arrSource(3, lColSource))
                End If
            End If
        End If
        
        ' *** AJOUT NOTES PONDEREES ***
        If bEvalOK Then
            If (arrSource(4, lColSource) = vbNullString) Then arrSource(4, lColSource) = arrSource(4, lColSource - 1)
            
            ' *** VERIFICATION COEFF / NOTE EVAL VALIDE ***
            If (arrSource(byLigListePage3, lColSource) <> vbNullString) And IsNumeric(CDbl(arrSource(byLigListePage3, lColSource))) Then
            
                ' *** AJOUT COMPETENCE PONDEREE DANS DOMAINE ***
                If (arrSource(4, lColSource) <> "Note / 20") Then
                    byIndiceDomaine = CByte(GetIndexInArray(arrChoixDomaines, arrSource(4, lColSource)))
                    
                    ' *** VERIFICATION DOMAINE VALIDE (PAR SECURITE) ***
                    If (byIndiceDomaine <> 0) Then
                        dbCoeffCompet = CDbl(arrSource(byLigListePage3, lColSource))
                        arrCoeff(4 * (byIndiceDomaine - 1) + byTrimestre) = arrCoeff(4 * (byIndiceDomaine - 1) + byTrimestre) + CDbl(arrSource(byLigListePage3, lColSource))
                        arrCoeff(4 * byIndiceDomaine) = arrCoeff(4 * byIndiceDomaine) + CDbl(arrSource(byLigListePage3, lColSource))
                        For byEleve = 1 To byNbEleves
                            arrDest(byEleve, 4 * (byIndiceDomaine - 1) + byTrimestre) = arrDest(byEleve, 4 * (byIndiceDomaine - 1) + byTrimestre) + dbCoeffCompet * CDbl(ConvertirLettreEnValeur(arrSource(byLigListePage3 + byEleve, lColSource)))
                            arrDest(byEleve, 4 * byIndiceDomaine) = arrDest(byEleve, 4 * byIndiceDomaine) + dbCoeffCompet * CDbl(ConvertirLettreEnValeur(arrSource(byLigListePage3 + byEleve, lColSource)))
                        Next byEleve
                    End If
                    
                ' *** AJOUT MOYENNE PONDEREE DANS TRIMESTRE ***
                Else
                    arrCoeff(4 * byNbDomaines + byTrimestre) = arrCoeff(4 * byNbDomaines + byTrimestre) + dbCoeffEval
                    arrCoeff(4 * (byNbDomaines + 1)) = arrCoeff(4 * (byNbDomaines + 1)) + dbCoeffEval
                    For byEleve = 1 To byNbEleves
                        arrDest(byEleve, 4 * byNbDomaines + byTrimestre) = arrDest(byEleve, 4 * byNbDomaines + byTrimestre) + dbCoeffEval * CDbl(arrSource(byLigListePage3 + byEleve, lColSource))
                        arrDest(byEleve, 4 * (byNbDomaines + 1)) = arrDest(byEleve, 4 * (byNbDomaines + 1)) + dbCoeffEval * CDbl(arrSource(byLigListePage3 + byEleve, lColSource))
                    Next byEleve
                    bEvalOK = False
                End If
            End If
        End If
    Next lColSource
    
    ' *** CALCUL VALEURS FINALES ***
    For lColDest = LBound(arrDest, 2) To UBound(arrDest, 2)
        For byEleve = 1 To byNbEleves
            If Not (IsEmpty(arrDest(byEleve, lColDest))) And Not (IsEmpty(arrCoeff(lColDest))) Then
            
                ' *** CALCUL MOYENNE DOMAINES ***
                If lColDest <= 4 * byNbDomaines Then
                    arrDest(byEleve, lColDest) = ConvertirValeurEnLettre(arrDest(byEleve, lColDest) / arrCoeff(lColDest))
                    
                ' *** CALCUL MOYENNE TRIMESTRE ***
                Else
                    arrDest(byEleve, lColDest) = arrDest(byEleve, lColDest) / arrCoeff(lColDest)
                End If
            End If
        Next byEleve
    Next lColDest
    
    With ThisWorkbook.Worksheets(GetNomPage4(byClasse))
        .Range(.Cells(byLigListePage4 + 1, 2), .Cells(byLigListePage4 + byNbEleves, 1 + 4 * (byNbDomaines + 1))) = arrDest
    End With
End Sub


