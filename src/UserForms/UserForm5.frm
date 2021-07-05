VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "Opération en cours"
   ClientHeight    =   1104
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4788
   OleObjectBlob   =   "UserForm5.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Function limiterPourcentage(byValeur As Byte) As Byte
    limiterPourcentage = byValeur
    If limiterPourcentage < 0 Then
        limiterPourcentage = 0
    ElseIf limiterPourcentage > 100 Then
        limiterPourcentage = 100
    End If
End Function

Public Sub updateAvancement(byAvancementActuel As Byte, byAvancementTotal As Byte)
    Dim byPourc As Byte
    byPourc = limiterPourcentage(100 * byAvancementActuel / byAvancementTotal)
    UserForm5.lblChargement.Caption = byPourc & "% terminé"
    UserForm5.barChargement.Width = CInt(2.16 * CDbl(byPourc)) ' Actual loading bar width is (object_width - 4) --> 220 - 4 = 216
    UserForm5.Repaint
End Sub

