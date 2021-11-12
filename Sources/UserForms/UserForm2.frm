VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Gestion listes - Ajouter un élève"
   ClientHeight    =   3036
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   4788
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
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
'                             UserForm 2 - Ajout d'élève
' *******************************************************************************
'
'   Fonctions publiques
'
'   Procédures publiques
'       SetUp()
'
'   Fonctions privées
'
'   Procédures privées
'       TbxEleve_Nom_Change()
'       TbxEleve_Prenom_Change()
'       BtnAjouterEleve_Click()
'
' *******************************************************************************

Option Explicit

' *******************************************************************************
'                               Fonctions publiques
' *******************************************************************************

' *******************************************************************************
'                               Procédures publiques
' *******************************************************************************

' Initialisation de l'UF
Public Sub SetUp()
    Dim byNbClasses As Byte
    Dim byClasse As Byte
    Dim strClasse As String

    byNbClasses = GetNombreClasses
    
    For byClasse = 1 To byNbClasses
        strClasse = GetNomClasse(byClasse)
        lbxClasse.AddItem strClasse
    Next byClasse
    
    lbxClasse.ListIndex = 0
End Sub

' *******************************************************************************
'                               Fonctions privées
' *******************************************************************************

' *******************************************************************************
'                               Procédures privées
' *******************************************************************************

' Mise en forme auto du nom de l'élève
Private Sub TbxEleve_Nom_Change()
    tbxEleve_Nom.Value = StrConv(tbxEleve_Nom.Value, vbUpperCase)
End Sub

' Mise en forme auto du prénom de l'élève
Private Sub TbxEleve_Prenom_Change()
    tbxEleve_Prenom.Value = StrConv(tbxEleve_Prenom.Value, vbProperCase)
End Sub

' Demande de confirmation puis ajout d'un nouvel élève
Private Sub BtnAjouterEleve_Click()
    Dim byClasse As Byte
    Dim strClasse As String
    Dim byEleve As Byte
    Dim strEleve As String
    
    byClasse = lbxClasse.ListIndex + 1
    strClasse = GetNomClasse(byClasse)
    strEleve = tbxEleve_Nom.Value & " " & tbxEleve_Prenom.Value
    byEleve = GetIndiceEleve(strEleve, byClasse, False)
    
    If vbYes = MsgBox("Vous êtes sur le point d'ajouter '" & strEleve & "' à la classe de " & strClasse & ". Voulez-vous poursuivre ?", vbYesNo, "Confirmation d'ajout") Then
        AjouterEleve byClasse, byEleve, strEleve
        Call DisplayTemporaryMessage("Élève ajouté")
    Else
        Call DisplayTemporaryMessage("Operation annulée")
    End If
End Sub

