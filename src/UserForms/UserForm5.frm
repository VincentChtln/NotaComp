VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "Modifier/Supprimer �valuation"
   ClientHeight    =   5040
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4788
   OleObjectBlob   =   "UserForm5.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm5"
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
'                 UserForm 5 - Modification/Suppression d'�valuation
' *******************************************************************************
'
'   Fonctions publiques
'
'   Proc�dures publiques
'       SetUp(ByVal byMode As UserFormMode, ByVal byClasse As Byte)
'
'   Fonctions priv�es
'
'   Proc�dures priv�es
'       BtnValider_Click()
'
' *******************************************************************************

Option Explicit

' *******************************************************************************
'                                   Variables
' *******************************************************************************

Dim byModeActuel As UserFormMode
Dim byClasseActuelle As Byte

' *******************************************************************************
'                               Fonctions publiques
' *******************************************************************************

' *******************************************************************************
'                               Proc�dures publiques
' *******************************************************************************

Public Sub SetUp(ByVal byMode As UserFormMode, ByVal byClasse As Byte)
    ' *** VARIABLES ***
    Dim arrEvals As Variant
    Dim byEval As Byte
    
    ' *** AFFECTATION VALEUR ***
    byModeActuel = byMode
    byClasseActuelle = byClasse
    arrEvals = GetArrayEvals(byClasseActuelle)
    
    ' *** ECRITURE LABELS ***
    If byModeActuel = Modifier Then
        Me.Caption = "Classe " & GetNomClasse(byClasseActuelle) & " - Modifier une �valuation"
        lblMessage = "Choisissez une �valuation � modifier"
    ElseIf byModeActuel = Supprimer Then
        Me.Caption = "Classe " & GetNomClasse(byClasseActuelle) & " - Supprimer une �valuation"
        lblMessage = "Choisissez une �valuation � supprimer"
    End If
    
    ' *** VERIFICATION ARRAY EVALS ***
    If Not IsArray(arrEvals) Then
        Exit Sub
    End If
    
    ' *** REMPLISSAGE TABLEAU ***
    lbxEvals.ColumnWidths = "30;80;30;30"
    lbxEvals.TextAlign = fmTextAlignCenter
    For byEval = LBound(arrEvals, 1) To UBound(arrEvals, 1)
        lbxEvals.AddItem
        lbxEvals.List(byEval - 1, 0) = byEval
        lbxEvals.List(byEval - 1, 1) = arrEvals(byEval, 1)
        lbxEvals.List(byEval - 1, 2) = arrEvals(byEval, 2)
        lbxEvals.List(byEval - 1, 3) = arrEvals(byEval, 3)
    Next byEval
End Sub

' *******************************************************************************
'                               Fonctions priv�es
' *******************************************************************************

' *******************************************************************************
'                               Proc�dures priv�es
' *******************************************************************************

Private Sub BtnValider_Click()
    If byModeActuel = Modifier Then
        If vbOK = MsgBox("Confirmez-vous la modification de l'�valuation '" & lbxEvals.List(lbxEvals.ListIndex, 1) & "' ?", vbOKCancel, "Demande confirmation") Then
            Call UserForm4.SetUp(byModeActuel, byClasseActuelle, lbxEvals.ListIndex + 1)
            Call UserForm4.Show
            Call Unload(UserForm5)
        End If
    ElseIf byModeActuel = Supprimer Then
        If vbOK = MsgBox("Confirmez-vous la suppression de l'�valuation '" & lbxEvals.List(lbxEvals.ListIndex, 1) & "' ?", vbOKCancel, "Demande confirmation") Then
            Call SupprimerEval(byClasseActuelle, lbxEvals.ListIndex + 1)
            Call Unload(UserForm5)
        End If
    End If
End Sub
