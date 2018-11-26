Attribute VB_Name = "modDevTools"
' modDevTools: Tools for helping Access databases interact with git-based source control in a
' semi-rational, semi-manageable sort of way
'
' Derived in part from code posted at https://stackoverflow.com/questions/16948215/exporting-ms-access-forms-and-class-modules-recursively-to-text-files
' and at https://christopherjmcclellan.wordpress.com/2014/10/10/vba-and-git/
'
' @version 2018.1126

Option Explicit

Public Const CFG_GIT_BASH_PATH = "C:\Users\charlesw.johnson\Documents\GitWin64\git-bash.exe"
Public Const CFG_GIT_SCRIPT_PATH = "/m/git-add-commit.bash"

'Derived from code posted at https://stackoverflow.com/questions/16948215/exporting-ms-access-forms-and-class-modules-recursively-to-text-files
'Modified to allow the user to set a desired path
'Requires Tools > References > Microsoft Visual Basic for Applications Extensibility Library for VBComponent
'Requires Tools > References > Microsoft Office ... Object Library for FileDialog
Public Sub ExportAllCode(Optional ByVal Path As String)
    
    Dim Project As VBProject
    Dim Component As VBComponent
    Dim Sfx As String

    Dim sDestinationFolder As String
    Dim dlgDestinationFolder As FileDialog
    
    If Len(Path) > 0 Then
        Let sDestinationFolder = Path
    Else
    
        Set dlgDestinationFolder = Application.FileDialog(msoFileDialogFolderPicker): With dlgDestinationFolder
            .Title = "Export Destination Folder"
            .InitialFileName = IIf(Len(Path) > 0, Path, CurrentProject.Path)
        End With
                
        If dlgDestinationFolder.Show Then
            Let sDestinationFolder = dlgDestinationFolder.SelectedItems(1)
        End If
    End If
    
    If Len(sDestinationFolder) > 0 Then

        For Each Project In Application.VBE.VBProjects
            For Each Component In Project.VBComponents
                ExportCodeToSource Component:=Component, SourcePath:=sDestinationFolder
            Next Component
        Next Project
    
    End If
    
End Sub

Public Sub CommitAllCode(Optional ByVal Path As String)
    Dim sDestinationFolder As String
    Dim dlgDestinationFolder As FileDialog
    
    If Len(Path) > 0 Then
        Let sDestinationFolder = Path
    Else
        Set dlgDestinationFolder = Application.FileDialog(msoFileDialogFolderPicker): With dlgDestinationFolder
            .Title = "Export Destination Folder"
            .InitialFileName = IIf(Len(Path) > 0, Path, CurrentProject.Path)
        End With
                    
        If dlgDestinationFolder.Show Then
            Let sDestinationFolder = dlgDestinationFolder.SelectedItems(1)
        End If
    End If
    
    If Len(sDestinationFolder) > 0 Then
        ExportAllCode Path:=sDestinationFolder
        
        Shell PathName:=CFG_GIT_BASH_PATH & " -c '" & CFG_GIT_SCRIPT_PATH & " """ & sDestinationFolder & """'"
    End If
End Sub

Public Sub ReImport(Module As String, Optional ByVal Path As String)
    Dim sSourceFolder As String
    Dim dlgSourceFolder As FileDialog
    
    If Len(Path) > 0 Then
        Let sSourceFolder = Path
    Else
        Set dlgSourceFolder = Application.FileDialog(msoFileDialogFolderPicker): With dlgSourceFolder
            .Title = "ReImport Source Folder"
            .InitialFileName = IIf(Len(Path) > 0, Path, CurrentProject.Path)
        End With
                    
        If dlgSourceFolder.Show Then
            Let sSourceFolder = dlgSourceFolder.SelectedItems(1)
        End If
    End If
    
    If Len(sSourceFolder) > 0 Then
        Dim Project As VBProject
        Dim Comp As VBComponent
        Dim FSO As New FileSystemObject
        Dim sourceName As String
        
        Set Project = Application.VBE.ActiveVBProject
        For Each Comp In Project.VBComponents
            Dim Ext As String
            Dim ComponentName As String
            
            Let Ext = ExpectedExImFileExtension(Comp, Import:=True)
            If Comp.Name = Module And Len(Ext) > 0 Then
                'Look to see if we have an ex-im file in the version-controlled directory
                Let ComponentName = Comp.Name
                Let sourceName = sSourceFolder & "\" & ComponentName & "." & Ext
                
                If FSO.FileExists(sourceName) Then
                    RemoveCode Project:=Project, Component:=Comp, Repo:=sSourceFolder, FSO:=FSO
                    ImportCodeFromSource Project:=Project, ComponentName:=ComponentName, SourcePath:=sourceName
                Else
                    MsgBox sourceName & " not found in repository."
                End If
            End If
        Next Comp
    End If
End Sub

Public Sub DiffCode(Repo As String, Optional ByRef FSO As FileSystemObject)
    If FSO Is Nothing Then
        Set FSO = New FileSystemObject
    End If

    Dim RecoveryPath As String
    Dim RecoveryUniqId As String
    
    Let RecoveryUniqId = FORMAT(Now, "YYYYmmddHHmmss")
    Let RecoveryPath = Repo & "\.vba_exim"
    If Not FSO.FolderExists(RecoveryPath) Then
        FSO.CreateFolder Path:=RecoveryPath
    End If
    
    Let RecoveryPath = RecoveryPath & "\diff-" & RecoveryUniqId
    If Not FSO.FolderExists(RecoveryPath) Then
        FSO.CreateFolder Path:=RecoveryPath
    End If

    ExportAllCode Path:=RecoveryPath
    
    Dim CFG_DIFF_SCRIPT_PATH As String
    Let CFG_DIFF_SCRIPT_PATH = "clear ; diff --unified --color --exclude=" & ExImSubDirectory & " --exclude=.git --exclude=.gitignore"
    Shell PathName:=CFG_GIT_BASH_PATH & " -c '" & CFG_DIFF_SCRIPT_PATH & " """ & RecoveryPath & """ """ & Repo & """; read -p ""Press any key to continue..."" -n1'"

End Sub

Public Function ExImSubDirectory() As String
    Let ExImSubDirectory = ".vba_exim"
End Function

Public Sub RemoveCode(Project As VBProject, Component As VBComponent, Repo As String, Optional ByRef FSO As FileSystemObject)
    
    If FSO Is Nothing Then
        Set FSO = New FileSystemObject
    End If
    
    Dim RecoveryPath As String
    Dim RecoveryUniqId As String
    
    Let RecoveryUniqId = FORMAT(Now, "YYYYmmdd")
    Let RecoveryPath = Repo & "\" & ExImSubDirectory
    If Not FSO.FolderExists(RecoveryPath) Then
        FSO.CreateFolder Path:=RecoveryPath
    End If
    
    Let RecoveryPath = RecoveryPath & "\" & RecoveryUniqId
    If Not FSO.FolderExists(RecoveryPath) Then
        FSO.CreateFolder Path:=RecoveryPath
    End If
    
    'Export present state of the code module.
    Dim ComponentUniqId As String
    Dim ComponentFileName As String
    
    Let ComponentUniqId = FORMAT(Now, "YYYYmmddHHnnss")
    Let ComponentFileName = Component.Name & "-" & ComponentUniqId & "." & ExpectedExImFileExtension(Component)
    ExportCodeToSource Component:=Component, SourcePath:=RecoveryPath & "\" & ComponentFileName
    
    Project.VBComponents.Remove Component

End Sub

Public Sub ExportCodeToSource(ByRef Component As Variant, ByVal SourcePath As String, Optional ByRef FSO As FileSystemObject)
    If FSO Is Nothing Then
        Set FSO = New FileSystemObject
    End If
    
    Dim oComponent As VBComponent
    Dim vComponent As VBComponent
    
    If TypeName(Component) = "String" Then
        For Each vComponent In Application.VBE.ActiveVBProject.VBComponents
            If vComponent.Name = Component Then
                Set oComponent = vComponent
            End If
        Next vComponent
    Else
        Set oComponent = Component
    End If
    
    If Right(SourcePath, 1) = "\" Then
        Let SourcePath = SourcePath & ExImSourceFileName(oComponent)
    ElseIf FSO.FolderExists(SourcePath) Then
        Let SourcePath = SourcePath & "\" & ExImSourceFileName(oComponent)
    End If
    
    oComponent.Export FileName:=SourcePath
End Sub

Public Sub ImportCodeFromSource(Project As VBProject, ByVal ComponentName As String, ByVal SourcePath As String)
    Project.VBComponents.Import SourcePath
End Sub

Public Function ExImSourceFileName(Component As VBComponent, Optional ByVal Import As Boolean) As String
    Dim Ext As String
    
    Let Ext = ExpectedExImFileExtension(Component, Import)
    If Len(Ext) > 0 Then
        Let ExImSourceFileName = Component.Name & "." & Ext
    End If
End Function

Public Function ExpectedExImFileExtension(v As Variant, Optional ByVal Import As Boolean) As String
    Dim T As vbext_ComponentType
    Dim Ext As String
    
    If TypeName(v) = "vbext_ComponentType" Or TypeName(v) = "Long" Or TypeName(v) = "Integer" Then
        Let T = v
    ElseIf TypeName(v) = "VBComponent" Then
        Let T = v.Type
    Else
        Err.Raise Number:=13, Description:="Type mismatch: Parameter v must be of type VBComponent or vbext_ComponentType"
    End If
    
    Select Case T
    Case vbext_ComponentType.vbext_ct_Document
        If Not Import Then
            Let Ext = "cls"
        End If
    Case vbext_ComponentType.vbext_ct_ClassModule
        Let Ext = "cls"
    Case vbext_ComponentType.vbext_ct_StdModule
        Let Ext = "bas"
    Case vbext_ComponentType.vbext_ct_MSForm
        Let Ext = "frm"
    End Select
    
    Let ExpectedExImFileExtension = Ext
End Function
