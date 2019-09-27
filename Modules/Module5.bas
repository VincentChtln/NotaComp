Attribute VB_Name = "Module5"
' **********************************
' Mise à jour des modules - import & export
' **********************************

Const strGitFolder As String = "C:\Users\Utilisateur\Documents\GitHub\OutilNotationCompetence\Modules\"
Const strWBSource As String = "Outil de gestion des notes_Dev.xlsm"

Function isFileOpen(strFilePath As String) As Boolean
    Dim intFileNum As Integer, intErrNum As Integer

    On Error Resume Next        ' Turn error checking off.
    intFileNum = FreeFile()     ' Get a free file number.
    ' Attempt to open the file and lock it.
    Open strFilePath For Input Lock Read As #intFileNum
    Close intFileNum            ' Close the file.
    intErrNum = Err             ' Save the error number that occurred.
    On Error GoTo 0             ' Turn error checking back on.

    ' Check to see which error occurred.
    Select Case intErrNum
        ' No error occurred.
        Case 0
         isFileOpen = False
        ' Error number for "Permission Denied."
        Case 70
            isFileOpen = True
        ' Another error occurred.
        Case Else
            Error intErrNum
    End Select
End Function

Function isWBOpen(strWBName As String) As Boolean
    Dim wb As Workbook
    isWBOpen = False
    For Each wb In Workbooks
        If wb.Name = strWBName Then
            isWBOpen = True
            Exit For
        End If
    Next wb
End Function

Function isVBProjectProtected(wb As Workbook) As Boolean
    isVBProjectProtected = (wb.VBProject.Protection = 1)
End Function

Function isFolderEmpty(strFolderPath As String) As Boolean
    isFolderEmpty = (Dir(correctFolderPath(strFolderPath) & "*.*") = "")
End Function

Function correctFolderPath(strFolderPath As String) As String
    correctFolderPath = strFolderPath
    If Right(correctFolderPath, 1) <> "\" Then
        correctFolderPath = correctFolderPath & "\"
    End If
End Function

Function correctFileName(strFileName As String) As String
    correctFileName = strFileName
    If Right(correctFileName, 5) <> ".xlsm" Then
        correctFileName = correctFileName & ".xlsm"
    End If
End Function

Public Sub deleteModulesInVBProject(wb As Workbook)
    Dim cmpComponent As VBIDE.VBComponent

    ' Vérfication de la présence de modules & confirmation de suppression
    If Not isVBProjectProtected(wb) Then
        vbClean = MsgBox("Voulez-vous supprimer tous les modules VBA du projet ?", vbYesNoCancel, "Supprimer modules ?")
    Else
        MsgBox ("Projet VBA protégé, accès refusé.")
        Exit Sub
    End If
    
    ' Suppression / annulation
    If vbClean = vbYes Then
        For Each cmpComponent In wb.VBProject.VBComponents
            If cmpComponent.Type = vbext_ct_StdModule Then
                wb.VBProject.VBComponents.Remove cmpComponent
            End If
        Next cmpComponent
        MsgBox ("Modules VBA supprimés.")
    ElseIf vbClean = vbCancel Then
        MsgBox ("Opération annulée.")
        Exit Sub
    End If
End Sub

Public Sub deleteModulesInFolder(strFolderPath As String)
    strFolderPath = correctFolderPath(strFolderPath)
    
    ' Vérfication de la présence de modules & confirmation de suppression
    If Not isFolderEmpty(strFolderPath) Then
        vbClean = MsgBox("Voulez-vous supprimer tous les modules VBA du dossier '" & strFolderPath & "' ?", vbYesNoCancel, "Supprimer fichiers ?")
    Else
        MsgBox ("Aucun module VBA présent dans le dossier.")
        Exit Sub
    End If
    
    ' Suppression / annulation
    If vbClean = vbYes Then
        On Error Resume Next
        Kill strFolderPath & "\*.bas"
        On Error GoTo 0
        MsgBox ("Modules VBA supprimés.")
    ElseIf vbClean = vbCancel Then
        MsgBox ("Opération annulée.")
        Exit Sub
    End If
End Sub

Public Sub importModulesToVBProject()
    Dim wbTarget As Workbook
    Dim strWBFolder As String, strWBName As String
    Dim strImportFolder As String, strModulePath As String
    Dim bWBOK As Boolean, bImportFolderOK As Boolean
    Dim FSO As New FileSystemObject
    
    ' Ouverture du WB
    bWBOK = False
    While Not bWBOK
        strWBName = correctFileName(InputBox("Nom du classeur cible ?", "Nom du classeur", strWBName))
        If isWBOpen(strWBName) Then
            bWBOK = True
        Else
            strWBFolder = correctFolderPath(InputBox("Chemin vers le classeur cible", "Chemin du dossier", strWBFolder))
            MsgBox (strWBFolder & strWBName)
            If FSO.FolderExists(strWBFolder) And FSO.FileExists(strWBFolder & strWBName) Then
                bWBOK = True
                MsgBox ("Ouverture du WB '" & strWBFolder & strWBName & "'.")
                Workbooks.Open (strWBFolder & strWBName)
            Else
                If vbCancel = MsgBox("Fichier non trouvé, vérifier l'orthographe", vbOKCancel) Then Exit Sub
            End If
        End If
    Wend
    Set wbTarget = Workbooks(strWBName)
    
    ' Vérifie si le projet VB n'est pas protégé
    If isVBProjectProtected(wbTarget) Then
        MsgBox ("Projet VBA annulé, accès refusé. Opération annulée.")
        Exit Sub
    End If
    
    ' Vérifie le dossier d'import
    bImportFolderOK = False
    While Not bImportFolderOK
        strImportFolder = correctFolderPath(InputBox("Chemin vers le dossier de modules", vbOKCancel, strGitFolder))
        If FSO.FolderExists(strImportFolder) Then
            bImportFolderOK = True
            If isFolderEmpty(strImportFolder) Then
                MsgBox ("Dossier vide, opération annulée.")
                Exit Sub
            End If
        Else
            If vbCancel = MsgBox("Dossier non trouvé, vérifier l'orthographe", vbOKCancel) Then Exit Sub
        End If
    Wend
    
    ' Supprime les anciens modules
    Call deleteModulesInVBProject(wbTarget)
    
    ' Import les nouveaux modules
    For indexModule = 1 To 4
        strModulePath = strImportFolder & "Module" & indexModule & ".bas"
        If FSO.FileExists(strModulePath) Then
            wbTarget.VBProject.VBComponents.Import strModulePath
        Else
            MsgBox ("Import Module" & indexModule & ".bas échoué, fichier non trouvé.")
        End If
    Next indexModule
    
    ' Modifie la date de dernière MAJ
    Application.ScreenUpdating = False
    With wbTarget.Sheets(strPage1)
        .Unprotect strPassword
        .Range("G5").Value = strVersion
        .Range("G6").Value = Format(Now, "dd/mm/yyyy")
        .Protect strPassword
    End With
    Application.ScreenUpdating = True
    
    MsgBox ("Import terminé.")
    Set wbTarget = Nothing
End Sub

Public Sub exportModulesToFolder()
    Dim bValidFolder As Boolean
    Dim strFileName As String
    Dim FSO As New FileSystemObject
    Dim cmpComponent As VBIDE.VBComponent

    ' Vérification des conditions d'export
    If isVBProjectProtected(Workbooks(strWBSource)) Then
        MsgBox ("Projet VB protégé, accès refusé.")
        Exit Sub
    End If
    
    bValidFolder = False
    While Not bValidFolder
        If FSO.FolderExists(strGitFolder) Then
            bValidFolder = True
            deleteModulesInFolder (strGitFolder)
        Else
            MsgBox ("Le dossier indiqué '" & strGitFolder & "' n'existe pas, opération annulée.")
            Exit Sub
        End If
    Wend
    
    ' Export des modules
    For Each cmpComponent In Workbooks(strWBSource).VBProject.VBComponents
        If cmpComponent.Type = vbext_ct_StdModule Then
            strFileName = cmpComponent.Name & ".bas"
            cmpComponent.Export correctFolderPath(strGitFolder) & strFileName
        End If
    Next cmpComponent
    
    MsgBox ("Export terminé.")
End Sub
