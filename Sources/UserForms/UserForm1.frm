VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Modification classe"
   ClientHeight    =   2640
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4788
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
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
'                               UserForm 1
' *******************************************************************************
'
'   Fonctions publiques
'
'   Proc�dures publiques
'       SetUp(Optional ByVal byClasse As Byte = 0)
'
'   Fonctions priv�es
'
'   Proc�dures priv�es
'       BtnAjouter_Click()
'       BtnModifier_Click()
'       BtnSupprimer_Click()
'
' *******************************************************************************

Option Explicit

' *******************************************************************************
'                                   Variables
' *******************************************************************************

Public Enum UserFormMode
    NoValue = 0
    Ajouter = 1
    Modifier = 2
    Supprimer = 3
End Enum

Dim byClasseActuelle As Byte

' *******************************************************************************
'                               Fonctions publiques
' *******************************************************************************

' *******************************************************************************
'                               Proc�dures publiques
' *******************************************************************************

Public Sub SetUp(Optional ByVal byClasse As Byte = 0)
    ' Modification variable globale commune
    byClasseActuelle = byClasse
    
    ' byClasseActuelle = 0 --> Page "Listes"
    If byClasseActuelle = 0 Then
        Me.Caption = "Gestion listes de classe"
        btnAjouter.Caption = "Ajouter un �l�ve"
        btnModifier.Caption = "Transf�rer un �l�ve"
        btnSupprimer.Caption = "Supprimer un �l�ve"
        
    ' byClasseActuelle > 0 --> Num�ro de la classe
    Else
        Me.Caption = "Classe " & GetNomClasse(byClasseActuelle) & " - Gestion �valuations"
        btnAjouter.Caption = "Ajouter une �valuation"
        btnModifier.Caption = "Modifier une �valuation"
        btnSupprimer.Caption = "Supprimer une �valuation"
        btnModifier.Visible = False
    End If
End Sub

' *******************************************************************************
'                               Fonctions priv�es
' *******************************************************************************

' *******************************************************************************
'                               Proc�dures priv�es
' *******************************************************************************

Private Sub BtnAjouter_Click()
    ' byClasseActuelle = 0 --> Page "Listes"
    If byClasseActuelle = 0 Then
        Call UserForm2.SetUp
        Call UserForm2.Show(vbModeless)
        
    ' byClasseActuelle > 0 --> Num�ro de la classe
    Else
        Call UserForm4.SetUp(UserFormMode.Ajouter, byClasseActuelle, UserFormMode.NoValue)
        Call UserForm4.Show(vbModeless)
    End If
End Sub

Private Sub BtnModifier_Click()
    ' byClasseActuelle = 0 --> Page "Listes"
    If byClasseActuelle = 0 Then
        If Not (vbOK = MsgBox("ATTENTION: vous allez transf�rer un �l�ve entre d'une classe vers une autre." & vbNewLine & _
                              "Puisque rien ne garantit la comptabilit� des �valuations entre deux classes diff�rentes, " & _
                              "les notes de l'�l�ve ne seront pas transf�r�es et seront par cons�quent perdues dans le processus." & vbNewLine & _
                              "Si vous souhaitez les conserver, veillez � les relever dans un document � part. " & _
                              "Il tient ensuite � vous d'adapter les notes pr�c�demment acquises aux �valuations de sa nouvelle classe." & vbNewLine & vbNewLine & _
                              "Si vous souhaitez revenir � la fen�tre pr�c�dente, cliquez sur 'Annuler'.", vbOKCancel, "Message d'alerte")) Then Exit Sub
        Call UserForm3.SetUp(UserFormMode.Modifier)
        Call UserForm3.Show(vbModeless)
        
    ' byClasseActuelle > 0 --> Num�ro de la classe
    Else
        Call UserForm5.SetUp(UserFormMode.Modifier, byClasseActuelle)
        Call UserForm5.Show(vbModeless)
    End If
End Sub

Private Sub BtnSupprimer_Click()
    ' byClasseActuelle = 0 --> Page "Listes"
    If byClasseActuelle = 0 Then
        If Not (vbOK = MsgBox("ATTENTION: vous allez supprimer un �l�ve d'une classe. " & _
                              "Ses notes seront �galement perdues dans le processus." & vbNewLine & _
                              "Cette op�ration est irr�versible. Si vous souhaitez toutefois garder ses notes, " & _
                              "veuillez les enregister dans un document � part." & vbNewLine & vbNewLine & _
                              "Pour revenir � la fen�tre pr�c�dente, cliquez sur 'Annuler'.", _
                              vbOKCancel, "Message d'alerte")) Then Exit Sub
        Call UserForm3.SetUp(UserFormMode.Supprimer)
        Call UserForm3.Show(vbModeless)
        
    ' byClasseActuelle > 0 --> Num�ro de la classe
    Else
        If Not (vbOK = MsgBox("ATTENTION: vous allez supprimer une �valuation. " & _
                               "Les notes associ�es seront �galement perdues dans le processus." & vbNewLine & _
                               "Cette op�ration est irr�versible. Si vous souhaitez toutefois garder ces notes, " & _
                               "veuillez les enregister dans un document � part." & vbNewLine & vbNewLine & _
                               "Pour revenir � la fen�tre pr�c�dente, cliquez sur 'Annuler'.", _
                               vbOKCancel, "Message d'alerte")) Then Exit Sub
        Call UserForm5.SetUp(UserFormMode.Supprimer, byClasseActuelle)
        Call UserForm5.Show(vbModeless)
    End If
End Sub


