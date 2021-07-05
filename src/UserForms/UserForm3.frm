VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Supprimer un élève"
   ClientHeight    =   4440
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
' UF Suppression d'élève
' *******************************************************************************

Option Explicit

' *******************************************************************************
' PROCÉDURES
' *******************************************************************************
' UserForm_Initialize()
' lbxClasse_Change()
' btnSupprimerEleve_Click()
' *******************************************************************************

' Initialisation de l'UF
Private Sub UserForm_Initialize()
    ' *** DECLARATION VARIABLES ***
    Dim byNbClasses As Byte
    Dim byClasse As Byte
    
    ' *** AFFECTATION VARIABLES ***
    byNbClasses = getNombreClasses
    
    ' *** AJOUT CLASSE DANS LISTE ***
    For byClasse = 1 To byNbClasses
        lbxClasse.AddItem getNomClasse(byClasse)
    Next byClasse
    
    ' *** INITIALISATION SELECTION ***
    lbxClasse.ListIndex = 0
    lbxEleve.ListIndex = 0
End Sub

' Modification de la liste Eleve en fonction de la classe sélectionnée
Private Sub lbxClasse_Change()
    ' *** DECLARATION VARIABLES ***
    Dim byNbEleves As Byte
    Dim byEleve As Byte
    
    ' *** AFFECTATION VARIABLES ***
    lbxEleve.Clear
    byNbEleves = getNombreEleves(lbxClasse.ListIndex + 1)
    
    ' *** AJOUT ELEVES DANS LISTE ***
    For byEleve = 1 To byNbEleves
        lbxEleve.AddItem ThisWorkbook.Worksheets(strPage2).Cells(byLigListePage2 + byEleve, 2 * (lbxClasse.ListIndex + 1) - 1).Value
    Next byEleve
    
    ' *** INITIALISATION INDEX ***
    lbxEleve.ListIndex = 0
End Sub

' Demande de confirmation puis appel de la procédure supprimerEleve (Module 2)
Private Sub btnSupprimerEleve_Click()
    Dim byClasse As Byte
    Dim strClasse As String
    Dim byEleve As Byte
    Dim strEleve As String
    
    byClasse = 1 + lbxClasse.ListIndex
    strClasse = lbxClasse.Value
    byEleve = lbxEleve.ListIndex + 1
    strEleve = ThisWorkbook.Worksheets(strPage2).Cells(byLigListePage2 + byEleve, 2 * byClasse - 1)
    
    If vbYes = MsgBox("Vous êtes sur le point de supprimer '" & strEleve & "' de la classe de " & strClasse & ". " & _
                      "Voulez-vous poursuivre ?", vbYesNo) Then
        supprimerEleve byClasse, byEleve
        MsgBox "Élève supprimé"
    Else
        MsgBox "Operation annulée"
    End If
End Sub


