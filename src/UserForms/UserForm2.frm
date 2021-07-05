VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Ajouter un élève"
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
' UF Ajout d'élève
' *******************************************************************************

Option Explicit

' *******************************************************************************
' PROCÉDURES
' *******************************************************************************
' UserForm_Initialize()
' tbxNomEleve_Change()
' tbxPrenomEleve_Change()
' btnAjouterEleve_Click()
' *******************************************************************************

' Initialisation de l'UF
Private Sub UserForm_Initialize()
    Dim byNbClasses As Byte
    Dim byClasse As Byte
    Dim strClasse As String
    
    byNbClasses = getNombreClasses
    
    For byClasse = 1 To byNbClasses
        strClasse = getNomClasse(byClasse)
        lbxClasse.AddItem strClasse
    Next byClasse
    
    lbxClasse.ListIndex = 0
End Sub

' Mise en forme auto du nom de l'élève
Private Sub tbxNomEleve_Change()
    tbxNomEleve.Value = StrConv(tbxNomEleve.Value, vbUpperCase)
End Sub

' Mise en forme auto du prénom de l'élève
Private Sub tbxPrenomEleve_Change()
    tbxPrenomEleve.Value = StrConv(tbxPrenomEleve.Value, vbProperCase)
End Sub

' Demande de confirmation puis ajout d'un nouvel élève
Private Sub btnAjouterEleve_Click()
    Dim byClasse As Byte
    Dim strClasse As String
    Dim byEleve As Byte
    Dim strEleve As String
    
    byClasse = lbxClasse.ListIndex + 1
    strClasse = getNomClasse(byClasse)
    strEleve = tbxNomEleve.Value & " " & tbxPrenomEleve.Value
    byEleve = getIndiceEleve(strEleve, byClasse, False)
    
    If vbYes = MsgBox("Vous êtes sur le point d'ajouter '" & strEleve & "' à la classe de " & strClasse & ". Voulez-vous poursuivre ?", vbYesNo, "Confirmation d'ajout") Then
        ajouterEleve byClasse, byEleve, strEleve
        MsgBox "Élève ajouté"
    Else
        MsgBox "Operation annulée"
    End If
End Sub

