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

' Proc�dure d'ajout d'un �l�ve
Private Sub btnAjouterEleve_Click()
    UserForm2.Show
End Sub

' Proc�dure de suppression d'un �l�ve
Private Sub btnSupprimerEleve_Click()

    ' *** MESSAGE ALERTE ***
    If vbCancel = MsgBox("ATTENTION: vous allez supprimer un �l�ve d'une classe. Ses notes seront �galement supprim�es dans le processus." & Chr(13) & Chr(10) & "Cette op�ration est irr�versible." & Chr(13) & Chr(10) & "Si vous souhaitez toutefois garder ces notes, veuillez les relever dans un document � part." & Chr(13) & Chr(10) & "Pour revenir � la fen�tre pr�c�dente, cliquez sur 'Annuler'.", vbOKCancel, "Message d'alerte") Then Exit Sub
    
    UserForm3.Show
End Sub

' Proc�dure de transfert d'un �l�ve
Private Sub btnTransfererEleve_Click()

    ' *** MESSAGE ALERTE ***
    If vbCancel = MsgBox("ATTENTION: vous allez transf�rer un �l�ve entre deux classes diff�rentes. Veillez � bien relever ses notes dans un document � part, car elle ne seront pas transf�r�es lors du processus. Apr�s op�ration, elles disparaitront �galement de la classe d'origine." & Chr(13) & Chr(10) & "Cette op�ration est irr�versible, il sera donc impossible de r�cup�rer les notes perdues si elles n'ont pas �t� not�es ailleurs" & Chr(13) & Chr(10) & "Il tient ensuite � vous d'adapter les notes pr�c�demment acquises aux �valuations de sa nouvelle classe." & Chr(13) & Chr(10) & "Si vous souhaitez revenir � la fen�tre pr�c�dente, cliquez sur 'Annuler'.", vbOKCancel, "Message d'alerte") Then Exit Sub
     
    UserForm4.Show
End Sub
