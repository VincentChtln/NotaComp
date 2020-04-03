VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Modification classe"
   ClientHeight    =   2640
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3780
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Procédure d'ajout d'un élève
Private Sub btnAjouterEleve_Click()
    UserForm2.Show
End Sub

' Procédure de suppression d'un élève
Private Sub btnSupprimerEleve_Click()
    UserForm3.Show
End Sub

' Procédure de transfert d'un élève
Private Sub btnTransfererEleve_Click()
    UserForm4.Show
End Sub
