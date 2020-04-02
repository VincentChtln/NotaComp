VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Suppression d'élève"
   ClientHeight    =   5265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3420
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ##################################
' UF Suppression d'élève
' ##################################

Option Explicit

' **********************************
' PROCÉDURES
' **********************************
' UserForm_Initialize()
' listboxSelectionClasse_Change()
' btnSupprimerEleve_Click()
' **********************************

' Initialisation de l'UF
Private Sub UserForm_Initialize()
    Dim intNombreClasse As Integer, intIndiceClasse As Integer, strNomClasse As String
    
    intNombreClasse = getNombreClasses()
    
    For intIndiceClasse = 1 To intNombreClasse
        strNomClasse = getNomClasse(intIndiceClasse)
        listboxSelectionClasse.AddItem strNomClasse
    Next intIndiceClasse
    
    listboxSelectionClasse.ListIndex = 0
    listboxSelectionEleve.ListIndex = 0
End Sub

' Modification de la liste Eleve en fonction de la classe sélectionnée
Private Sub listboxSelectionClasse_Change()
    Dim intIndiceClasse As Integer, intColonneClasse As Integer
    Dim intNombreEleves As Integer, intIndiceEleve As Integer, strNomCompletEleve As String
    
    listboxSelectionEleve.Clear
    intIndiceClasse = listboxSelectionClasse.ListIndex + 1
    intColonneClasse = 2 * intIndiceClasse - 1
    intNombreEleves = getNombreEleves(intIndiceClasse)
    
    For intIndiceEleve = 1 To intNombreEleves
        strNomCompletEleve = Cells(3 + intIndiceEleve, intColonneClasse).Value
        listboxSelectionEleve.AddItem strNomCompletEleve
    Next intIndiceEleve
    
    listboxSelectionEleve.ListIndex = 0
End Sub

' Demande de confirmation puis appel de la procédure supprimerEleve (Module 2)
Private Sub btnSupprimerEleve_Click()
    Dim intIndiceClasse As Integer, strNomClasse As String
    Dim intIndiceEleve As Integer, strNomCompletEleve As String
    
    intIndiceClasse = 1 + listboxSelectionClasse.ListIndex
    strNomClasse = listboxSelectionClasse.Value
    intIndiceEleve = listboxSelectionEleve.ListIndex + 1
    strNomCompletEleve = Cells(3 + intIndiceEleve, 2 * intIndiceClasse - 1)
    
    If vbYes = MsgBox("Vous êtes sur le point de supprimer '" & strNomCompletEleve & "' de la classe de " & strNomClasse & ". Voulez-vous poursuivre ?", vbYesNo, "Confirmation de suppression") Then
        supprimerEleve intIndiceClasse, strNomCompletEleve
        MsgBox "Élève supprimé"
    Else
        MsgBox "Operation annulée"
    End If
End Sub


