Attribute VB_Name = "Module5"
' ##################################
' MAJ des modules
' ##################################

Option Explicit

' ##################################
' CONSTANTES
' ##################################
Const strGitFolder1 As String = "C:\Users\Utilisateur\Documents\GitHub\OutilNotationCompetence\"
Const strGitFolder2 As String = "C:\Users\vincent.chatelain\Documents\GitHub\NotaComp\"
Const strWBSource As String = "Outil de gestion des notes_Dev.xlsm"

' ##################################
' FONCTIONS
' ##################################
' isWBOpen(strWBName As String) As Boolean
' isVBProjectProtected(wb As Workbook) As Boolean
' isFolderEmpty(strFolderPath As String) As Boolean
' properFolderPath(strFolderPath As String) As String
' properWBName(strFileName As String) As String
' getModulesFolder(strUserName As String) As String
' ##################################

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

Function getModulesFolder(strUserName As String) As String
    Select Case strUserName
    Case Is = "Utilisateur"
        getModulesFolder = strGitFolder1
    Case Is = "CHATELAIN Vincent"
        getModulesFolder = strGitFolder2
    End Select
    getModulesFolder = getModulesFolder & "1_Modules\"
End Function

' ##################################
' PROCÉDURES
' ##################################
' userDisplay()
' updateVBProject()
' exportModulesToFolder()
' importModulesToVBProject()
' deleteFilesInFolder(strFolderPath As String)
' deleteModulesInVBProject(wb As Workbook)
' ##################################

Sub userDisplay()
    MsgBox "current user is " & Application.UserName
End Sub

Public Sub updateVBProject()
    exportModulesToFolder
    importModulesToVBProject
End Sub

Public Sub exportModulesToFolder()
    ' *** DECLARATION VARIABLES ***
    Dim strFileName As String
    Dim strExportFolder As String
    Dim bExportFolderOK As Boolean
    Dim FSO As New FileSystemObject
    Dim cpnFile As VBIDE.VBComponent

    ' *** VERIFICATION CONDITIONS EXPORT ***
    If isVBProjectProtected(Workbooks(strWBSource)) Then
        MsgBox ("Projet VB protégé, accès refusé.")
        Exit Sub
    End If
    
    ' *** AFFECTATION VARIABLES ***
    strExportFolder = getModulesFolder(Application.UserName)
    bExportFolderOK = False
    
    ' *** SUPPRESSION MODULES DOSSIER DEST ***
    Do
        If FSO.FolderExists(strExportFolder) Then
            If vbNo = MsgBox("Confirmation export des nouveaux modules ?", vbYesNo) Then GoTo Annulation
            bExportFolderOK = True
        Else
            InputBox "Le dossier indiqué n'existe pas. Modifiez le chemin d'accès ou annulez l'export.", "Export folder path", strExportFolder
            If Len(strExportFolder) = 0 Then GoTo Annulation
        End If
    Loop Until bExportFolderOK
    
    ' *** SUPPRESSION ANCIENS MODULES ***
    deleteFilesInFolder (strExportFolder)
    
    ' *** EXPORT MODULES ***
    For Each cpnFile In Workbooks(strWBSource).VBProject.VBComponents
        If cpnFile.Type = vbext_ct_StdModule Then
            strFileName = cpnFile.Name & ".bas"
            cpnFile.Export properFolderPath(strExportFolder) & strFileName
        ElseIf cpnFile.Type = vbext_ct_MSForm Then
            strFileName = cpnFile.Name & ".frm"
            cpnFile.Export properFolderPath(strExportFolder) & strFileName
        End If
    Next cpnFile
    
    ' *** MESSAGE INFORMATION ***
    MsgBox ("Export terminé avec succès.")
    Exit Sub
    
    ' *** MESSAGE ANNULATION ***
Annulation:
    MsgBox ("Opération annulée.")
End Sub

Public Sub importModulesToVBProject()
    ' *** DECLARATION VARIABLES ***
    Dim bWBOK As Boolean
    Dim wbTarget As Workbook
    Dim strWBFolder As String
    Dim strWBName As String
    Dim strImportFolder As String
    Dim bImportFolderOK As Boolean
    Dim strImportFile As String
    Dim FSO As New FileSystemObject
    
    ' *** VERIFICATION WB + OUVERTURE ***
    bWBOK = False
    Do
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
                If vbCancel = MsgBox("Fichier non trouvé, vérifier l'orthographe", vbOKCancel) Then Exit Sub
            End If
        End If
    Loop Until bWBOK
    Set wbTarget = Workbooks(strWBName)
    
    ' *** VERFICATION PROTECTION VB_PROJECT ***
    If isVBProjectProtected(wbTarget) Then
        MsgBox ("Projet VBA protégé, accès refusé.")
        GoTo Annulation
    End If
    
    ' *** VERIFICATION DOSSIER IMPORT ***
    bImportFolderOK = False
    Do
        strImportFolder = properFolderPath(InputBox("Chemin vers le dossier de modules", vbOKCancel, getModulesFolder(Application.UserName)))
        If Len(strImportFolder) = 0 Then GoTo Annulation
        If FSO.FolderExists(strImportFolder) Then
            If isFolderEmpty(strImportFolder) Then
                MsgBox ("Dossier vide.")
                GoTo Annulation
            End If
            bImportFolderOK = True
        End If
    Loop Until bImportFolderOK
    
    ' *** SUPPRESSION ANCIENS MODULES ***
    deleteModulesInVBProject wbTarget
    
    ' *** IMPORT NOUVEAUX MODULES ***
    strImportFile = "Module1.bas"
    Do Until Len(strImportFile) = 0
        If Right(strImportFile, 4) = ".bas" Or Right(strImportFile, 4) = ".frm" Then
            If strImportFile <> "Module5.bas" Then wbTarget.VBProject.VBComponents.Import strImportFolder & strImportFile
        End If
        strImportFile = Dir()
    Loop
    
    ' *** MODIFICATION DATA MAJ ***
    Application.ScreenUpdating = False
    With wbTarget.Worksheets(strPage1)
        .Unprotect strPassword
        .Range("G5").Value = strVersion
        .Range("G6").Value = Format(Now, "MM/dd/yyyy")      ' Affiche la date en format "dd/MM/yyyy"
        .Protect strPassword
    End With
    Application.ScreenUpdating = True
    
    ' *** MESSAGE INFORMATION ***
    MsgBox "Import terminé avec succès."
    Exit Sub
    
    ' *** MESSAGE ANNULATION ***
Annulation:
    MsgBox "Opération annulée."
End Sub

Public Sub deleteFilesInFolder(strFolderPath As String)
    ' *** FORMATAGE CHEMIN DOSSIER ***
    strFolderPath = properFolderPath(strFolderPath)
    
    ' *** SUPPRESSION FICHIERS DANS DOSSIER EXPORT ***
    On Error Resume Next
    Kill strFolderPath & "\*"   ' Suppression de tous les fichiers du dossier
    On Error GoTo 0
    
    ' *** MESSAGE INFORMATION ***
    MsgBox ("Fichiers VBA supprimés.")
End Sub

Public Sub deleteModulesInVBProject(wb As Workbook)
    ' *** DECLARATION VARIABLES ***
    Dim cpnFile As VBIDE.VBComponent
    
    ' *** SUPPRESSION FICHIERS VBA ***
    For Each cpnFile In wb.VBProject.VBComponents
        If cpnFile.Type = vbext_ct_StdModule Or cpnFile.Type = vbext_ct_MSForm Then
            wb.VBProject.VBComponents.Remove cpnFile
        End If
    Next cpnFile
    
    ' *** MESSAGE INFORMATION ***
    MsgBox ("Modules et UserForm VBA supprimés.")
End Sub

