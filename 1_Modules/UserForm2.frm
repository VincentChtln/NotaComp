VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Ajouter un élève"
   ClientHeight    =   4032
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   3780
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ##################################
' UF Ajout d'élève
' ##################################

Option Explicit

' **********************************
' PROCÉDURES
' **********************************
' UserForm_Initialize()
' textboxNomEleve_Change()
' textboxPrenomEleve_Change()
' btnAjouterEleve_Click()
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
End Sub

' Mise en forme auto du nom de l'élève
Private Sub textboxNomEleve_Change()
    textboxNomEleve.Value = StrConv(textboxNomEleve.Value, vbUpperCase)
End Sub

' Mise en forme auto du prénom de l'élève
Private Sub textboxPrenomEleve_Change()
    textboxPrenomEleve.Value = StrConv(textboxPrenomEleve.Value, vbProperCase)
End Sub

' Demande de confirmation puis ajout d'un nouvel élève
Private Sub btnAjouterEleve_Click()
    Dim intIndiceClasse As Integer, strNomClasse As String
    Dim intIndiceEleve As Integer, strNomCompletEleve As String
    
    intIndiceClasse = listboxSelectionClasse.ListIndex + 1
    strNomClasse = getNomClasse(intIndiceClasse)
    strNomCompletEleve = textboxNomEleve.Value & " " & textboxPrenomEleve.Value
    intIndiceEleve = getIndiceEleve(strNomCompletEleve, intIndiceClasse, False)
    
    If vbYes = MsgBox("Vous êtes sur le point d'ajouter '" & strNomCompletEleve & "' à la classe de " & strNomClasse & ". Voulez-vous poursuivre ?", vbYesNo, "Confirmation d'ajout") Then
        ajouterEleve intIndiceClasse, intIndiceEleve, strNomCompletEleve
        MsgBox "Élève ajouté"
    Else
        MsgBox "Operation annulée"
    End If
End Sub

