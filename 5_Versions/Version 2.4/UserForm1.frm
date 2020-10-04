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

Option Explicit

' Procédure d'ajout d'un élève
Private Sub btnAjouterEleve_Click()
    UserForm2.Show
End Sub

' Procédure de suppression d'un élève
Private Sub btnSupprimerEleve_Click()

    ' *** MESSAGE ALERTE ***
    If vbCancel = MsgBox("ATTENTION: vous allez supprimer un élève d'une classe. Ses notes seront également supprimées dans le processus." & Chr(13) & Chr(10) & "Cette opération est irréversible." & Chr(13) & Chr(10) & "Si vous souhaitez toutefois garder ces notes, veuillez les relever dans un document à part." & Chr(13) & Chr(10) & "Pour revenir à la fenêtre précédente, cliquez sur 'Annuler'.", vbOKCancel, "Message d'alerte") Then Exit Sub
    
    UserForm3.Show
End Sub

' Procédure de transfert d'un élève
Private Sub btnTransfererEleve_Click()

    ' *** MESSAGE ALERTE ***
    If vbCancel = MsgBox("ATTENTION: vous allez transférer un élève entre deux classes différentes. Veillez à bien relever ses notes dans un document à part, car elle ne seront pas transférées lors du processus. Après opération, elles disparaitront également de la classe d'origine." & Chr(13) & Chr(10) & "Cette opération est irréversible, il sera donc impossible de récupérer les notes perdues si elles n'ont pas été notées ailleurs" & Chr(13) & Chr(10) & "Il tient ensuite à vous d'adapter les notes précédemment acquises aux évaluations de sa nouvelle classe." & Chr(13) & Chr(10) & "Si vous souhaitez revenir à la fenêtre précédente, cliquez sur 'Annuler'.", vbOKCancel, "Message d'alerte") Then Exit Sub
     
    UserForm4.Show
End Sub
