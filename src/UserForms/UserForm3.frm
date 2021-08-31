VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Modifier/Supprimer un �l�ve"
   ClientHeight    =   5040
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   4788
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
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
'                          UserForm 3 - Modification d'�l�ve
' *******************************************************************************
'
'   Fonctions publiques
'
'   Proc�dures publiques
'       SetUp(ByVal byMode As UserFormMode)
'
'   Fonctions priv�es
'
'   Proc�dures priv�es
'       LbxClasseSource_Change()
'       BtnValider_Click()
'
' *******************************************************************************

Option Explicit

' *******************************************************************************
'                                   Variables
' *******************************************************************************

Dim byModeActuel As UserFormMode

' *******************************************************************************
'                               Fonctions publiques
' *******************************************************************************

' *******************************************************************************
'                               Proc�dures publiques
' *******************************************************************************

Public Sub SetUp(ByVal byMode As UserFormMode)
    ' *** VARIABLES ***
    Dim byNbClasses As Byte
    Dim byClasse As Byte
    
    ' *** AFFECTATION VARIABLES ***
    byNbClasses = GetNombreClasses
    byModeActuel = byMode
    
    ' *** AJOUT CLASSE DANS LISTE ***
    For byClasse = 1 To byNbClasses
        lbxClasseSource.AddItem GetNomClasse(byClasse)
    Next byClasse
    
    ' *** MODIFICATION AFFICHAGE ***
    If byModeActuel = Modifier Then
        Me.Caption = "Gestion listes - Transf�rer un �l�ve"
        lblClasseDest.Visible = True
        lbxClasseDest.Visible = True
    ElseIf byModeActuel = Supprimer Then
        Me.Caption = "Gestion listes - Supprimer un �l�ve"
        lblClasseDest.Visible = False
        lbxClasseDest.Visible = False
    Else
        Unload Me
    End If
    
    ' *** INITIALISATION SELECTION ***
    lbxClasseSource.ListIndex = 0
    Call LbxClasseSource_Change
    lbxEleveSource.ListIndex = 0
    If byModeActuel = UserFormMode.Modifier Then lbxClasseDest.ListIndex = 0
End Sub

' *******************************************************************************
'                               Fonctions priv�es
' *******************************************************************************

' *******************************************************************************
'                               Proc�dures priv�es
' *******************************************************************************

' Modification de la liste Eleve en fonction de la classe s�lectionn�e
Private Sub LbxClasseSource_Change()
    ' *** VARIABLES ***
    Dim byEleve As Byte
    Dim byNbEleves As Byte
    Dim byClasse As Byte
    Dim byNbClasses As Byte
    
    ' *** AFFECTATION VARIABLES ***
    lbxEleveSource.Clear
    lbxClasseDest.Clear
    byNbEleves = GetNombreEleves(lbxClasseSource.ListIndex + 1)
    
    ' *** AJOUT ELEVES DANS LISTE ***
    For byEleve = 1 To byNbEleves
        lbxEleveSource.AddItem ThisWorkbook.Worksheets(strPage2).Cells(byLigListePage2 + byEleve, 2 * (lbxClasseSource.ListIndex + 1) - 1).Value
    Next byEleve
    
    ' *** AJOUT CLASSES TRANSFERT DANS LISTE ***
    If byModeActuel = UserFormMode.Modifier Then
        byNbClasses = GetNombreClasses()
        For byClasse = 1 To byNbClasses
            If byClasse <> lbxClasseSource.ListIndex + 1 Then
                lbxClasseDest.AddItem ThisWorkbook.Worksheets(strPage1).Cells(byLigTabClasses + 2 + byClasse, byColTabClasses).Value
            End If
        Next byClasse
    End If
    
    ' *** INITIALISATION INDEX ***
    lbxEleveSource.ListIndex = 0
    If byModeActuel = UserFormMode.Modifier Then lbxClasseDest.ListIndex = 0
End Sub

Private Sub BtnValider_Click()
    ' *** VARIABLES ***
    Dim byClasseSource As Byte
    Dim strClasseSource As String
    Dim byClasseDest As Byte
    Dim strClasseDest As String
    Dim strEleve As String
    Dim byEleveSource As Byte
    Dim byEleveDest As Byte
    
    ' *** AFFECTATION VARIABLES ***
    strClasseSource = lbxClasseSource.Value
    strClasseDest = lbxClasseDest.Value
    byClasseSource = GetIndiceClasse(strClasseSource)
    byClasseDest = GetIndiceClasse(strClasseDest)
    byEleveSource = lbxEleveSource.ListIndex + 1
    strEleve = ThisWorkbook.Worksheets(strPage2).Cells(byLigListePage2 + byEleveSource, 2 * byClasseSource - 1)
    byEleveDest = GetIndiceEleve(strEleve, byClasseDest, False)
    
    ' *** TRANSFERT ***
    If byModeActuel = UserFormMode.Modifier Then
        If vbYes = MsgBox("Vous allez transf�rer '" & strEleve & "' de la classe '" & strClasseSource & "' vers la classe '" & strClasseDest & "'. " & _
                          "Confirmez-vous cette op�ration ?", vbYesNo, "Confirmation de transfert") Then
            Call TransfererEleve(byClasseSource, byEleveSource, byClasseDest, byEleveDest, strEleve)
            Call DisplayTemporaryMessage("�l�ve transf�r�.")
        Else
            Call DisplayTemporaryMessage("Operation annul�e.")
        End If
        
    ' *** SUPPRESSION ***
    ElseIf byModeActuel = UserFormMode.Supprimer Then
        If vbYes = MsgBox("Vous �tes sur le point de supprimer '" & strEleve & "' de la classe de " & strClasseSource & ". " & _
                          "Voulez-vous poursuivre ?", vbYesNo) Then
            Call SupprimerEleve(byClasseSource, byEleveSource)
            Call DisplayTemporaryMessage("�l�ve supprim�.")
        Else
            Call DisplayTemporaryMessage("Operation annul�e.")
        End If
    End If
    
    Unload Me
End Sub


