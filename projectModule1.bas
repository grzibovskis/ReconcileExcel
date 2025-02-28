Sub ExportMacroToGitHub()
    Dim fso As Object
    Dim File As Object
    Dim GitFile As Object
    Dim ModuleCode As String
    Dim GitHubFolder As String
    Dim GitRepo As String
    
    ' Folder where files will be exported
    GitHubFolder = "C:\Users\grzyb\Desktop\project"
    GitRepo = "https://github.com/grzibovskis/ReconcileExcel"
    
    ' Create File System Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Create Folder if not exists
    If Not fso.FolderExists(GitHubFolder) Then
        fso.CreateFolder GitHubFolder
    End If
    
    ' Export VBA Modules
    Dim Component As Object
    
    For Each Component In Application.VBE.VBProjects(1).VBComponents
        If Component.Type = 1 Then ' 1 = Module
            ModuleCode = Component.CodeModule.Lines(1, Component.CodeModule.CountOfLines)
            If ModuleCode <> "" Then
                Set GitFile = fso.CreateTextFile(GitHubFolder & Component.Name & ".bas", True)
                GitFile.Write ModuleCode
                GitFile.Close
            End If
        End If
    Next Component
    
    MsgBox "Export Successful!"

    ' Automate Git Commit & Push ??
    Dim Command As String
    Command = "cd /d " & GitHubFolder & " && git init && git remote add origin " & GitRepo & " && git add . && git commit -m ""VBA Macros Update"" && git push origin main"
    Shell "cmd.exe /c " & Command, vbNormalFocus
End Sub
