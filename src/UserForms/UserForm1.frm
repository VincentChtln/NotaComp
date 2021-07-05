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


Option Explicit

' Proc�dure d'ajout d'un �l�ve
Private Sub btnAjouterEleve_Click()
    UserForm2.Show
End Sub

' Proc�dure de suppression d'un �l�ve
Private Sub btnSupprimerEleve_Click()

    ' *** MESSAGE ALERTE ***
    If vbCancel = MsgBox("ATTENTION: vous allez supprimer un �l�ve d'une classe. Ses notes seront �galement perdues dans le processus." & vbNewLine & _
                         "Cette op�ration est irr�versible. Si vous souhaitez toutefois garder ses notes, veuillez les enregister dans un document � part." & vbNewLine & vbNewLine & _
                         "Pour revenir � la fen�tre pr�c�dente, cliquez sur 'Annuler'.", vbOKCancel, "Message d'alerte") Then Exit Sub
    
    UserForm3.Show
End Sub

' Proc�dure de transfert d'un �l�ve
Private Sub btnTransfererEleve_Click()

    ' *** MESSAGE ALERTE ***
    If vbCancel = MsgBox("ATTENTION: vous allez transf�rer un �l�ve entre d'une classe vers une autre." & vbNewLine & _
                         "Puisque rien ne garantit la comptabilit� des �valuations entre deux classes diff�rentes, " & _
                         "les notes de l'�l�ve ne seront pas transf�r�es et seront par cons�quent perdues dans le processus." & vbNewLine & _
                         "Si vous souhaitez les conserver, veillez � les relever dans un document � part. " & _
                         "Il tient ensuite � vous d'adapter les notes pr�c�demment acquises aux �valuations de sa nouvelle classe." & vbNewLine & vbNewLine & _
                         "Si vous souhaitez revenir � la fen�tre pr�c�dente, cliquez sur 'Annuler'.", vbOKCancel, "Message d'alerte") Then Exit Sub
     
    UserForm4.Show
End Sub


