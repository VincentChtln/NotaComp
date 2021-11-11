VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "Ajout/modif éval"
   ClientHeight    =   5040
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4788
   OleObjectBlob   =   "UserForm4.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' *******************************************************************************
'   Copyright (C)
'   Date: 2021
'   Auteur: Vincent Chatelain
' *******************************************************************************
'
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
' *******************************************************************************


' *******************************************************************************
'                   UserForm 4 - Ajout/modification d'évaluation
' *******************************************************************************
'
'   Fonctions publiques
'
'   Procédures publiques
'       SetUp(ByVal byMode As UserFormMode, ByVal byClasse As Byte, Optional ByVal byEval As Byte = 0)
'
'   Fonctions privées
'
'   Procédures privées
'       BtnValider_Click()
'
' *******************************************************************************

Option Explicit

' *******************************************************************************
'                                   Variables
' *******************************************************************************

Dim byModeActuel As UserFormMode
Dim byClasseActuelle As Byte
Dim byEvalActuelle As Byte

' *******************************************************************************
'                               Fonctions publiques
' *******************************************************************************

' *******************************************************************************
'                               Procédures publiques
' *******************************************************************************

Public Sub SetUp(ByVal byMode As UserFormMode, ByVal byClasse As Byte, Optional ByVal byEval As Byte = 0)
    ' *** VARIABLES ***
    Dim arrEvals As Variant
    Dim arrEvalCompets As Variant
    Dim arrChoixCompets As Variant
    Dim arrDomaines As Variant
    Dim byDomaine As Variant
    Dim byCompet As Variant
    Dim byCompetEval As Variant
    Dim byNbCompet As Byte
    
    ' *** AFFECTATION VALEURS ***
    byModeActuel = byMode
    byClasseActuelle = byClasse
    If byEval = 0 Then byEvalActuelle = ThisWorkbook.Worksheets(GetNomPage3(byClasseActuelle)).Buttons.Count - 1
    arrChoixCompets = GetArrayChoixCompetences()
    arrDomaines = GetArrayDomaines()
    
    ' *** AJOUT COMPET DANS LISTE ***
    lbxEval_Compet.ColumnWidths = "30; 200"
    For byDomaine = 1 To 8
        byNbCompet = GetSizeOfArray(arrChoixCompets(byDomaine)) / 2
        If byNbCompet > 1 Then
            For byCompet = LBound(arrChoixCompets(byDomaine), 1) To UBound(arrChoixCompets(byDomaine), 1)
                lbxEval_Compet.AddItem
                lbxEval_Compet.List(lbxEval_Compet.ListCount - 1, 0) = arrChoixCompets(byDomaine)(byCompet, 1)
                lbxEval_Compet.List(lbxEval_Compet.ListCount - 1, 1) = arrChoixCompets(byDomaine)(byCompet, 2)
            Next byCompet
        ElseIf byNbCompet = 1 Then
            lbxEval_Compet.AddItem
            lbxEval_Compet.List(lbxEval_Compet.ListCount - 1, 0) = arrChoixCompets(byDomaine)(1)
            lbxEval_Compet.List(lbxEval_Compet.ListCount - 1, 1) = arrChoixCompets(byDomaine)(2)
        End If
    Next byDomaine
    
    ' *** AJOUT EVALUATION ***
    If byMode = UserFormMode.Ajouter Then
        Me.Caption = "Classe " & GetNomClasse(byClasseActuelle) & " - Ajouter une évaluation"
    
    ' *** MODIFICATION EVALUATION ***
    ElseIf byMode = UserFormMode.Modifier Then
        Me.Caption = "Classe " & GetNomClasse(byClasseActuelle) & " - Modifier une évaluation"
        
        arrEvals = GetArrayEvals(byClasseActuelle)
        arrEvalCompets = arrEvals(byEvalActuelle, 6)
        tbxEval_Nom = arrEvals(byEvalActuelle, 1)
        tbxEval_Trimestre = arrEvals(byEvalActuelle, 2)
        tbxEval_Coeff = arrEvals(byEvalActuelle, 3)
        
        ' *** SELECTION COMPET EVALUEES DANS LISTE ***
        For byCompetEval = 1 To GetSizeOfArray(arrEvalCompets)
            For byCompet = 0 To lbxEval_Compet.ListCount - 1
                If (lbxEval_Compet.List(byCompet, 0) = arrEvalCompets(byCompetEval)) Then
                    lbxEval_Compet.Selected(byCompet) = True
                End If
            Next byCompet
        Next byCompetEval
    End If
End Sub

' *******************************************************************************
'                               Fonctions privées
' *******************************************************************************

Private Function IsCompetEval(ByVal varValue As Variant) As Boolean
    Dim byCompet As Byte
    Dim byNbCompetEval As Byte
    Dim bFound As Boolean
    
    byNbCompetEval = lbxEval_Compet.ListCount
    IsInCompetEval = False
    For byCompet = 1 To byNbCompetEval
        If lbxEval_Compet.Items(byCompet - 1).ToString = varValue Then
            IsInCompetEval = True
        End If
    Next byCompet
End Function

' *******************************************************************************
'                               Procédures privées
' *******************************************************************************

Private Sub BtnValider_Click()
    ' *** VARIABLES ***
    Dim arrEval_CompetsEvaluees As Variant
    Dim arrChoixCompet As Variant
    Dim arrTampon As Variant
    Dim byDomaine As Byte
    Dim byCompetSrc As Byte
    Dim byCompetList As Byte
    Dim byCompetList_Debut As Byte
    Dim byCompetDest As Byte
    Dim byNbCompetSrc As Byte
    Dim byNbCompetListSelect As Byte
    Dim byNbCompetDest As Byte
    Dim strEval_Nom As String
    Dim byEval_Trimestre As Byte
    Dim dbEval_Coeff As Double
    
    ' *** VERIFICATION DONNEES USERFORM - NOM EVAL ***
    If (tbxEval_Nom = "" Or Not Application.WorksheetFunction.IsText(tbxEval_Nom)) Then
        Call MsgBox("Le nom de cette évaluation n'est pas valide.")
        Exit Sub
    End If
    ' *** VERIFICATION DONNEES USERFORM - TRIMESTRE EVAL ***
    If (tbxEval_Trimestre = "" Or Not IsNumeric(tbxEval_Trimestre) Or Not (tbxEval_Trimestre = 1 Or tbxEval_Trimestre = 2 Or tbxEval_Trimestre = 3)) Then
        Call MsgBox("Le trimestre de cette évaluation n'est pas valide.")
        Exit Sub
    End If
    ' *** VERIFICATION DONNEES USERFORM - COEFF EVAL ***
    If (tbxEval_Coeff = "" Or Not IsNumeric(tbxEval_Coeff) Or Not (tbxEval_Coeff > 0)) Then
        Call MsgBox("Le coefficient de cette évaluation n'est pas valide.")
        Exit Sub
    End If
    ' *** VERIFICATION DONNEES USERFORM - COMPETENCES EVAL ***
    byNbCompetListSelect = 0
    For byCompetList = 0 To lbxEval_Compet.ListCount - 1
        If lbxEval_Compet.Selected(byCompetList) Then byNbCompetListSelect = byNbCompetListSelect + 1
    Next byCompetList
    If (byNbCompetListSelect < 1) Then
        Call MsgBox("Veuillez sélectionner au moins 1 compétence pour créer cette évaluation.")
        Exit Sub
    End If
    
    ' *** AFFECTATION VARIABLES ***
    arrChoixCompet = GetArrayChoixCompetences()
    ReDim arrEval_CompetsEvaluees(1 To 8)
    ReDim arrTampon(1 To 2, 1 To 1)
    byCompetList_Debut = 0
    
    ' *** RECONSTRUCTION ARRAY COMPETS EVALUEES ***
    For byDomaine = 1 To 8
        byNbCompetSrc = GetSizeOfArray(arrChoixCompet(byDomaine)) / 2
        byCompetDest = 1
        byNbCompetDest = 0
        ' *** 1 COMPET -> ARRAY 1D ***
        If byNbCompetSrc = 1 Then
            For byCompetList = byCompetList_Debut To lbxEval_Compet.ListCount - 1
                If lbxEval_Compet.Selected(byCompetList) Then
                    If arrChoixCompet(byDomaine)(1) = lbxEval_Compet.List(byCompetList, 0) Then
                        ReDim arrTampon(1 To 2, 1 To byCompetDest)
                        arrTampon(1, byCompetDest) = lbxEval_Compet.List(byCompetList, 0)
                        arrTampon(2, byCompetDest) = lbxEval_Compet.List(byCompetList, 1)
                        byCompetList_Debut = byCompetList_Debut + 1
                        byNbCompetDest = byNbCompetDest + 1
                    End If
                End If
            Next byCompetList
        ' *** PLUSIEURS COMPET -> ARRAY 2D ***
        ElseIf byNbCompetSrc > 1 Then
            For byCompetSrc = 1 To byNbCompetSrc
                For byCompetList = byCompetList_Debut To lbxEval_Compet.ListCount - 1
                    If lbxEval_Compet.Selected(byCompetList) Then
                        If arrChoixCompet(byDomaine)(byCompetSrc, 1) = lbxEval_Compet.List(byCompetList, 0) Then
                            ReDim Preserve arrTampon(1 To 2, 1 To byCompetDest)
                            arrTampon(1, byCompetDest) = lbxEval_Compet.List(byCompetList, 0)
                            arrTampon(2, byCompetDest) = lbxEval_Compet.List(byCompetList, 1)
                            byCompetDest = byCompetDest + 1
                            byNbCompetDest = byNbCompetDest + 1
                            byCompetList_Debut = byCompetList_Debut + 1
                            GoTo NextCompetSrc
                        End If
                    End If
                Next byCompetList
NextCompetSrc:
            Next byCompetSrc
        End If
        If byNbCompetDest > 0 Then
            arrEval_CompetsEvaluees(byDomaine) = Application.WorksheetFunction.Transpose(arrTampon)
        End If
    Next byDomaine
    
    If byModeActuel = Ajouter Then
        Call DisableUpdates
        Call AjouterEvaluation(byClasseActuelle, arrEval_CompetsEvaluees, tbxEval_Nom, tbxEval_Coeff, tbxEval_Trimestre)
        Call EnableUpdates
    ElseIf byModeActuel = Modifier Then
        ' Call ModifierEvaluation(byClasseActuelle, byEval, arrEval_CompetsEvaluees, tbxEval_Nom, tbxEval_Coeff, tbxEval_Trimestre)
    End If
    Call Unload(UserForm4)
End Sub
