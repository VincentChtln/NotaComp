Attribute VB_Name = "Module5"
' ##################################
' MAJ des modules
' ##################################

Option Explicit

' **********************************
' CONSTANTES
' **********************************
Const strGitFolder As String = "C:\Users\Utilisateur\Documents\GitHub\OutilNotationCompetence\1_Modules\"
Const strWBSource As String = "Outil de gestion des notes_Dev.xlsm"

' **********************************
' FONCTIONS
' **********************************
' isWBOpen(strWBName As String) As Boolean
' isVBProjectProtected(wb As Workbook) As Boolean
' isFolderEmpty(strFolderPath As String) As Boolean
' properFolderPath(strFolderPath As String) As String
' properWBName(strFileName As String) As String
' **********************************

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
    isFolderEmpty = (Dir(properFolderPath(strFolderPath) & "*.*") = vbNullString)
End Function

Function properFolderPath(strFolderPath As String) As String
    properFolderPath = strFolderPath
    If Right(properFolderPath, 1) <> "\" Then
        properFolderPath = properFolderPath & "\"
    End If
End Function

Function properWBName(strFileName As String) As String
    properWBName = strFileName
    If Right(properWBName, 5) <> ".xlsm" Then
        properWBName = properWBName & ".xlsm"
    End If
End Function

' **********************************
' PROC�DURES
' **********************************

Public Sub updateVBProject()
    Call exportModulesToFolder
    Call importModulesToVBProject
End Sub

Public Sub exportModulesToFolder()
    Dim bValidFolder As Boolean
    Dim strFileName As String
    Dim FSO As New FileSystemObject
    Dim cmpComponent As VBIDE.VBComponent

    ' V�rification des conditions d'export
    If isVBProjectProtected(Workbooks(strWBSource)) Then
        MsgBox ("Projet VB prot�g�, acc�s refus�.")
        Exit Sub
    End If
    
    bValidFolder = False
    While Not bValidFolder
        If FSO.FolderExists(strGitFolder) Then
            bValidFolder = True
            deleteModulesInFolder (strGitFolder)
        Else
            MsgBox ("Le dossier indiqu� '" & strGitFolder & "' n'existe pas, op�ration annul�e.")
            Exit Sub
        End If
    Wend
    
    ' Export des modules
    For Each cmpComponent In Workbooks(strWBSource).VBProject.VBComponents
        If cmpComponent.Type = vbext_ct_StdModule Then
            strFileName = cmpComponent.Name & ".bas"
            cmpComponent.Export properFolderPath(strGitFolder) & strFileName
        End If
    Next cmpComponent
    
    MsgBox ("Export termin�.")
End Sub

Public Sub importModulesToVBProject()
    Dim wbTarget As Workbook
    Dim strWBFolder As String, strWBName As String
    Dim strImportFolder As String, strModulePath As String
    Dim bWBOK As Boolean, bImportFolderOK As Boolean
    Dim FSO As New FileSystemObject
    Dim indexModule As Integer
    
    ' Ouverture du WB
    bWBOK = False
    While Not bWBOK
        strWBName = properWBName(InputBox("Nom du classeur cible ?", "Nom du classeur", strWBName))
        If isWBOpen(strWBName) Then
            bWBOK = True
        Else
            strWBFolder = properFolderPath(InputBox("Chemin vers le classeur cible", "Chemin du dossier", strWBFolder))
            MsgBox (strWBFolder & strWBName)
            If FSO.FolderExists(strWBFolder) And FSO.FileExists(strWBFolder & strWBName) Then
                bWBOK = True
                MsgBox ("Ouverture du WB '" & strWBFolder & strWBName & "'.")
                Workbooks.Open (strWBFolder & strWBName)
            Else
                If vbCancel = MsgBox("Fichier non trouv�, v�rifier l'orthographe", vbOKCancel) Then Exit Sub
            End If
        End If
    Wend
    Set wbTarget = Workbooks(strWBName)
    
    ' V�rifie si le projet VB n'est pas prot�g�
    If isVBProjectProtected(wbTarget) Then
        MsgBox ("Projet VBA annul�, acc�s refus�. Op�ration annul�e.")
        Exit Sub
    End If
    
    ' V�rifie le dossier d'import
    bImportFolderOK = False
    While Not bImportFolderOK
        strImportFolder = properFolderPath(InputBox("Chemin vers le dossier de modules", vbOKCancel, strGitFolder))
        If FSO.FolderExists(strImportFolder) Then
            bImportFolderOK = True
            If isFolderEmpty(strImportFolder) Then
                MsgBox ("Dossier vide, op�ration annul�e.")
                Exit Sub
            End If
        Else
            If vbCancel = MsgBox("Dossier non trouv�, v�rifier l'orthographe", vbOKCancel) Then Exit Sub
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
            MsgBox ("Import Module" & indexModule & ".bas �chou�, fichier non trouv�.")
        End If
    Next indexModule
    
    ' Modifie la date de derni�re MAJ
    Application.ScreenUpdating = False
    With wbTarget.Sheets(strPage1)
        .Unprotect strPassword
        .Range("G5").Value = strVersion
        .Range("G6").Value = Format(Now, "dd/mm/yyyy")
        .Protect strPassword
    End With
    Application.ScreenUpdating = True
    
    MsgBox ("Import termin�.")
    Set wbTarget = Nothing
End Sub

Public Sub deleteModulesInFolder(strFolderPath As String)
    Dim vbClean As Variant
    strFolderPath = properFolderPath(strFolderPath)
    
    ' V�rfication de la pr�sence de modules & confirmation de suppression
    If Not isFolderEmpty(strFolderPath) Then
        vbClean = MsgBox("Voulez-vous supprimer tous les modules VBA du dossier '" & strFolderPath & "' ?", vbYesNoCancel, "Supprimer fichiers ?")
    Else
        MsgBox ("Aucun module VBA pr�sent dans le dossier.")
        Exit Sub
    End If
    
    ' Suppression / annulation
    If vbClean = vbYes Then
        On Error Resume Next
        Kill strFolderPath & "\*.bas"
        On Error GoTo 0
        MsgBox ("Modules VBA supprim�s.")
    ElseIf vbClean = vbCancel Then
        MsgBox ("Op�ration annul�e.")
        Exit Sub
    End If
End Sub

Public Sub deleteModulesInVBProject(wb As Workbook)
    Dim cmpComponent As VBIDE.VBComponent
    Dim vbClean As Variant

    ' V�rfication de la pr�sence de modules & confirmation de suppression
    If Not isVBProjectProtected(wb) Then
        vbClean = MsgBox("Voulez-vous supprimer tous les modules VBA du projet ?", vbYesNoCancel, "Supprimer modules ?")
    Else
        MsgBox ("Projet VBA prot�g�, acc�s refus�.")
        Exit Sub
    End If
    
    ' Suppression / annulation
    If vbClean = vbYes Then
        For Each cmpComponent In wb.VBProject.VBComponents
            If cmpComponent.Type = vbext_ct_StdModule Then
                wb.VBProject.VBComponents.Remove cmpComponent
            End If
        Next cmpComponent
        MsgBox ("Modules VBA supprim�s.")
    ElseIf vbClean = vbCancel Then
        MsgBox ("Op�ration annul�e.")
        Exit Sub
    End If
End Sub

