VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Ajouter un �l�ve"
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
' UF Ajout d'�l�ve
' *******************************************************************************

Option Explicit

' *******************************************************************************
' PROC�DURES
' *******************************************************************************
' UserForm_Initialize()
' textboxNomEleve_Change()
' textboxPrenomEleve_Change()
' btnAjouterEleve_Click()
' *******************************************************************************

' Initialisation de l'UF
Private Sub UserForm_Initialize()
    Dim intNbClasse As Integer
    Dim intIndiceClasse As Integer
    Dim strNomClasse As String
    
    intNbClasse = getNombreClasses()
    
    For intIndiceClasse = 1 To intNbClasse
        strNomClasse = getNomClasse(intIndiceClasse)
        listboxSelectionClasse.AddItem strNomClasse
    Next intIndiceClasse
    
    listboxSelectionClasse.ListIndex = 0
End Sub

' Mise en forme auto du nom de l'�l�ve
Private Sub textboxNomEleve_Change()
    textboxNomEleve.Value = StrConv(textboxNomEleve.Value, vbUpperCase)
End Sub

' Mise en forme auto du pr�nom de l'�l�ve
Private Sub textboxPrenomEleve_Change()
    textboxPrenomEleve.Value = StrConv(textboxPrenomEleve.Value, vbProperCase)
End Sub

' Demande de confirmation puis ajout d'un nouvel �l�ve
Private Sub btnAjouterEleve_Click()
    Dim intIndiceClasse As Integer
    Dim strNomClasse As String
    Dim intIndiceEleve As Integer
    Dim strNomCompletEleve As String
    
    intIndiceClasse = listboxSelectionClasse.ListIndex + 1
    strNomClasse = getNomClasse(intIndiceClasse)
    strNomCompletEleve = textboxNomEleve.Value & " " & textboxPrenomEleve.Value
    intIndiceEleve = getIndiceEleve(strNomCompletEleve, intIndiceClasse, False)
    
    If vbYes = MsgBox("Vous �tes sur le point d'ajouter '" & strNomCompletEleve & "' � la classe de " & strNomClasse & ". Voulez-vous poursuivre ?", vbYesNo, "Confirmation d'ajout") Then
        ajouterEleve intIndiceClasse, intIndiceEleve, strNomCompletEleve
        MsgBox "�l�ve ajout�"
    Else
        MsgBox "Operation annul�e"
    End If
End Sub

