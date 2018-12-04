Attribute VB_Name = "modDevTools"
'**
'* modDevTools: Tools for helping Access databases interact with git-based source control in a
'* semi-rational, semi-manageable sort of way
'*
'* Derived in part from code posted at https://stackoverflow.com/questions/16948215/exporting-ms-access-forms-and-class-modules-recursively-to-text-files
'* and at https://christopherjmcclellan.wordpress.com/2014/10/10/vba-and-git/
'*
'* @author C. Johnson
'* @version 2018.1204
'*
'* @uses VBIDE.VBComponent          {0002E157-0000-0000-C000-000000000046}/5.3 # Microsoft Visual Basic for Applications Extensibility Library
'* @uses VBScript_RegExp_55.RegExp  {3F4DACA7-160D-11D2-A8E9-00104B365C9F}/5.5
'* @uses Office.FileDialog          {2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}/2.8
'* @uses modMyVBA                   @file://adahfs3/Data/ADAHusers/CJohnson/Projects/MyVBA
'**

Option Explicit

Public Const CFG_GIT_BASH_PATH = "C:\Users\charlesw.johnson\Documents\GitWin64\git-bash.exe"
Public Const CFG_GIT_SCRIPT_PATH = "/m/bin/git-add-commit.bash"
Public Const CFG_DIFFDIR_SCRIPT_PATH = "/m/bin/diff-vs-git.bash"

Private dCommonRepository As Dictionary
Private oFSO As FileSystemObject

Private Property Get FSO() As FileSystemObject
    If oFSO Is Nothing Then
        Set oFSO = New FileSystemObject
    End If
    Set FSO = oFSO
End Property

Private Property Set FSO(RHS As FileSystemObject)
    Set oFSO = FSO
End Property

Private Property Get CommonRepository(Optional ByVal Context As String) As String
    Dim sCommonRepository As String
    
    If dCommonRepository Is Nothing Then
        Set dCommonRepository = New Dictionary
    End If
    
    If dCommonRepository.Exists(Context) Then
        Let sCommonRepository = dCommonRepository.Item(Context)
    End If
    
    If Len(sCommonRepository) = 0 Then
        Dim sOptionName As String
        
        Let sOptionName = "VBA.Repository"
        If Len(Context) > 0 Then
            Let sOptionName = sOptionName & "." & Context & "|" & sOptionName
        End If
        
        Let sCommonRepository = GetOption(SettingName:=sOptionName, Default:="")
    End If
        
    Let CommonRepository = sCommonRepository
End Property

Private Property Let CommonRepository(Optional ByVal Context As String, RHS As String)
    If dCommonRepository Is Nothing Then
        Set dCommonRepository = New Dictionary
    End If
    
    If Not dCommonRepository.Exists(Context) Then
        dCommonRepository.Add Key:=Context, Item:=RHS
    Else
        Let dCommonRepository.Item(Context) = RHS
    End If
End Property

'**
'* ExportAllCode: export all VBA code modules in the database to a desired repository path
'*
'* Originally derived from code posted at https://stackoverflow.com/questions/16948215/exporting-ms-access-forms-and-class-modules-recursively-to-text-files
'* Now substantially modified to allow the user to set a desired path and to re-use code from a single-file export sub
'*
'* @uses VBIDE.VBComponent {0002E157-0000-0000-C000-000000000046}/5.3 # Microsoft Visual Basic for Applications Extensibility Library
'* @uses InputMaybeRepositoryPath
'* @uses ExportCodeToSource
'*
'* @param String Path Repository path to export to.
'*      If omitted, a FileDialog with an interactive picker will be popped up.
'**
Public Sub ExportAllCode(Optional ByVal Path As String, Optional ByVal Context As String)
    
    Dim sDestinationFolder As String
    
    If Len(Context) = 0 Then Let Context = "Export"
    
    Let sDestinationFolder = InputMaybeRepositoryPath(Path:=Path, Title:="Export Destination Folder", InitialFileName:=CommonRepository(Context:=Context))
    
    If Len(sDestinationFolder) > 0 Then
        
        Dim Project As VBProject
        Dim Component As VBComponent

        For Each Project In Application.VBE.VBProjects
            For Each Component In Project.VBComponents
                ExportCodeToSource Component:=Component, SourcePath:=sDestinationFolder, Context:=Context
            Next Component
        Next Project
    
    End If
    
End Sub

'**
'* CommitAllCode: export all VBA code modules to a desired repository path, then launch a script which should
'* allow you to add and commit any changes relative to the prior working copy in that repository
'*
'* @uses InputMaybeRepositoryPath
'* @uses ExportAllCode
'* @uses CFG_GIT_BASH_PATH provides fully-qualified path to Mingw64 bash shell
'* @uses CFG_GIT_SCRIPT_PATH provides fully-qualified path to a bash script to perform interactive incremental add/commit
'*
'* @param String Path Repository path to export to.
'*      If omitted, a FileDialog with an interactive picker will be popped up.
'**
Public Sub CommitAllCode(Optional ByVal Path As String, Optional ByVal Context As String)
    Dim sDestinationFolder As String
    
    If Len(Context) = 0 Then Let Context = "Export"
    
    Let sDestinationFolder = InputMaybeRepositoryPath(Path:=Path, Title:="Export Destination Folder", InitialFileName:=CommonRepository(Context:=Context))
    
    If Len(sDestinationFolder) > 0 Then
        ExportAllCode Path:=sDestinationFolder
        
        Shell PathName:=CFG_GIT_BASH_PATH & " -c '" & CFG_GIT_SCRIPT_PATH & " """ & ToMinGWPath(sDestinationFolder) & """'"
    End If
End Sub

'**
'* ReImportRepo: re-import all the VBA code modules found in a desired repository path, overwriting present
'* working copies in Access
'*
'* @param String Path Repository path to re-import from.
'*      If omitted, a FileDialog with an interactive picker will be popped up.
'*
'* @uses InputMaybeRepositoryPath
'* @uses FileSystemObject
'* @uses IsExImSourceFileExtension
'* @uses ReImportCodeFromSource
'**
Public Sub ReImportRepo(Optional ByVal Path As String)
    Dim sSourceFolder As String
    
    Let sSourceFolder = InputMaybeRepositoryPath(Path:=Path, Title:="ReImport Source Folder", InitialFileName:=CommonRepository)
    
    If Len(sSourceFolder) > 0 Then
        Dim oFile As File
        Dim oFolder As Folder
        Dim sComponentName As String
        
        If FSO.FolderExists(sSourceFolder) Then
            Set oFolder = FSO.GetFolder(FolderPath:=sSourceFolder)
            For Each oFile In oFolder.Files
            
                If IsExImSourceFileExtension(oFile) Then
                    Let sComponentName = FSO.GetBaseName(oFile)
                    ReImportCodeFromSource Component:=sComponentName, Path:=Path
                End If
                
            Next oFile
        End If
    End If
End Sub

'**
'* ExportCodeToSource: export one or more VBA code modules to source code file(s) in a given repository path
'*
'* @param Variant Component
'* @param String SourcePath
'* @param Boolean SubDirectory
'* @param String Context
'*
'* @uses FSO
'* @uses GetVBComponents
'* @uses ExImSourceFileURL
'* @uses VBIDE.VBComponent
'**
Public Sub ExportCodeToSource(ByRef Component As Variant, Optional ByVal SourcePath As String, Optional ByRef SubDirectory As Boolean, Optional ByVal Context As String)
    
    Dim vComponent As VBComponent
    Dim cComponents As Collection
    Dim FullSourcePath As String
    Dim Project As VBProject
    
    If Len(Context) = 0 Then Let Context = "Export"
    
    Set cComponents = GetVBComponents(Component)
    For Each vComponent In cComponents
        Let FullSourcePath = ExImSourceFileURL(Component:=vComponent, Path:=SourcePath, Context:=Context)
        If Not SubDirectory Then
            Let CommonRepository = FSO.GetParentFolderName(FullSourcePath)
        End If
        vComponent.Export FileName:=FullSourcePath
        
        Set Project = vComponent.Collection.Parent
        ProcessCodeModuleVBDoc Component:=vComponent, Project:=Project, SourcePath:=FullSourcePath, Context:=Context
    
    Next vComponent
End Sub

'**
'* ReImportCodeFromSource: replace the current working copy of a code module in a given project, with a fresh copy
'* of the module from a file-system code repository. Useful to update project code to the newest version.
'*
'* @param VBComponent|String|Collection Component The Component or list of Components to re-import
'*      Can be represented by a live VBComponent object, a String containing the name of the Component
'*      or a Collection containing multiple VBComponent objects and/or strings. If multiple modules
'*      are provided, then each one is replaced in turn, one after another.
'* @param String Path A path to a source code file to import, or to the repo directory where the source
'*      file or files are located. If a directory is provided, the sub will guess a file name based on
'*      the module name and type. If omitted, the repo directory in CommonRepository will be assumed.
'*
'* @uses FSO
'* @uses RemoveCode
'* @uses ImportCodeFromSource
'* @uses GetVBComponents
'* @uses ExImSourceFileExtension
'* @uses ExImSourceFileURL
'* @uses VBIDE.VBComponent
'* @uses VBIDE.VBProject
'**
Public Sub ReImportCodeFromSource(Component As Variant, Optional ByVal Path As String)
    
    Dim Project As VBProject
    Dim SourceName As String
        
    Dim cComponents As Collection
    Dim vComponent As VBComponent
    
    Set cComponents = GetVBComponents(Component)

    For Each vComponent In cComponents
        Dim Ext As String
        Dim ComponentName As String
            
        Let Ext = ExImSourceFileExtension(vComponent, Import:=True)
        If Len(Ext) > 0 Then
            'Look to see if we have an ex-im file in the version-controlled directory
            Let ComponentName = vComponent.Name
                
            Let SourceName = ExImSourceFileURL(Component:=vComponent, Path:=Path)
            If FSO.FileExists(SourceName) Then
                
                'Located an ex-im file in the version-controlled directory. Remove the old and busted code,
                'and then import the new hotness.
                Set Project = vComponent.Collection.Parent
                
                RemoveCode Project:=Project, Component:=vComponent, Repo:=FSO.GetParentFolderName(SourceName)
                ImportCodeFromSource Project:=Project, ComponentName:=ComponentName, SourcePath:=SourceName
                    
            Else
                
                MsgBox SourceName & " not found in repository."
                        
            End If
        End If
    Next vComponent
End Sub

Public Sub ImportCodeFromSource(Project As VBProject, Component As Variant, ByVal SourcePath As String)
    
    Dim cComponentNames As Collection
    Dim vComponentName As Variant
    
    Dim oComponent As VBComponent
    
    If TypeName(Component) = "String" Then
        Set cComponentNames = New Collection
        cComponentNames.Add Component
    ElseIf TypeName(Component) = "Collection" Then
        Set cComponentNames = Component
    Else
        Err.Raise Number:=13, Description:="Type mismatch: Parameter Component must be of type String or Collection"
    End If
    
    DoEvents
    For Each vComponentName In cComponentNames
        Dim FullSourcePath As String
        
        Let FullSourcePath = ExImSourceFileURL(Component:=CStr(vComponentName), Path:=SourcePath)
        Set oComponent = Project.VBComponents.Import(FileName:=FullSourcePath)
        If Not oComponent Is Nothing Then
            Let oComponent.Name = vComponentName
        End If
        
        DoEvents
        If Not oComponent.Saved Then
            DoCmd.Save ObjectType:=acModule, ObjectName:=oComponent.Name
        End If
    
        ProcessCodeModuleVBDoc Component:=oComponent, Project:=Project, SourcePath:=SourcePath
    Next vComponentName
    
    
End Sub

Public Sub RemoveCode(Component As Variant, Repo As String)
       
    Dim Project As VBProject
    Dim cComponents As Collection
    Dim vComponent As VBComponent
    
    Set cComponents = GetVBComponents(Component)
    For Each vComponent In cComponents
        ExportCodeToBackupSubdirectory Component:=vComponent, SourcePath:=Repo
    
        Set Project = vComponent.Collection.Parent
        Project.VBComponents.Remove vComponent
    Next vComponent
End Sub

Public Sub ExportCodeToBackupSubdirectory(Component As Variant, SourcePath As String)
   
    '1. --- First, make sure that a subdirectory for ex-im transients and short-term backups exists.
    Dim RecoveryPath As String
    Dim RecoveryUniqId As String
    
    Let RecoveryPath = SourcePath & "\" & ExImSubDirectory
    If Not FSO.FolderExists(RecoveryPath) Then
        FSO.CreateFolder Path:=RecoveryPath
    End If
    
    '2. --- Next, make sure that the transients/short-term backups directory includes a subdirectory for to-day.
    Let RecoveryUniqId = FORMAT(Now, "YYYYmmdd")
    Let RecoveryPath = RecoveryPath & "\" & RecoveryUniqId
    If Not FSO.FolderExists(RecoveryPath) Then
        FSO.CreateFolder Path:=RecoveryPath
    End If
    
    '3. --- Export present state of the code module to a short-term backup.
    Dim ComponentUniqId As String
    Dim ComponentFileName As String
    
    Let ComponentUniqId = FORMAT(Now, "YYYYmmddHHnnss")
    Let ComponentFileName = Component.Name & "-" & ComponentUniqId & "." & ExImSourceFileExtension(Component)
    
    ExportCodeToSource Component:=Component, _
        SourcePath:=RecoveryPath & "\" & ComponentFileName, _
        SubDirectory:=True

End Sub

'**
'* ProcessCodeModuleVBDoc: process a code module being exported or imported to find DocBlock-style directives.
'* Some directives will be handed off to procedures (for example, @uses or @doc)
'*
'* @param VBProject Project
'* @param VBComponent Component
'* @param String SourcePath
'* @param String Context Can be "Import" or "Export"
'*
'* @uses ProcessCodeModuleVBDocDependency
'* @uses ProcessCodeModuleVBDocFileOutput
'* @uses VBScript_RegExp_55.RegExp
'**
Public Sub ProcessCodeModuleVBDoc(Project As VBProject, Component As VBComponent, ByVal SourcePath As String, Optional ByVal Context As String)
    Dim vbDocRegex As New RegExp
    
    Dim I As Integer, N As Integer
    Dim Line As String
    Dim vbDocRef As Object
    
    If Len(Context) = 0 Then
        Let Context = "Import"
    End If
    
    If FSO.FileExists(SourcePath) Then
        Let SourcePath = FSO.GetParentFolderName(SourcePath)
    End If
    
    Let I = 1
    Let N = Component.CodeModule.CountOfLines
    Do Until I > N
        DoEvents
        
        Let Line = Component.CodeModule.Lines(I, 1)
        
        ' N.B.: It's best to Roll Your Own on the regular expression matching here
        ' and not rely on anything in modRegularExpressionFunctions or modMyVBA,
        ' which would normally be a more attractive way to implement. Why? Well,
        ' because we might be re-importing modRegularExpressionFunctions or modMyVBA
        ' and in those cases, trying to use any function from either would cause VBA
        ' to pitch a fit.
        With vbDocRegex
            .IgnoreCase = True
            .Pattern = "'[*]?\s*@([A-Za-z]+)\s+(.*)\s*$"
            Set vbDocRef = .Execute(sourceString:=Line)
        End With
        
        If vbDocRef.Count > 0 Then
            Dim Match
            
            For Each Match In vbDocRef
                Dim Key
                Dim Value
                
                Let Key = Match.SubMatches(0)
                Let Value = Match.SubMatches(1)
                
                If Key = "version" Then
                    Debug.Print "Version: ", Value
                ElseIf (Key = "uses") Then
                    ProcessCodeModuleVBDocDependency Project:=Project, Component:=Component, Line:=Line, Key:=Key, Value:=Value, SourcePath:=SourcePath, Context:=Context
                ElseIf (Key = "doc") Then
                    ProcessCodeModuleVBDocFileOutput Project:=Project, Component:=Component, LineNumber:=I, LastLine:=N, Key:=Key, Value:=Value, SourcePath:=SourcePath, Context:=Context
                End If
            Next Match
            
        End If
        
        I = I + 1
    Loop
    
End Sub

'**
'* ProcessCodeModuleVBDocFileOutput: process a @doc directive, which allows text files to be generated from comment blocks
'* (for example, you might want to use this to create a portable README or a .gitignore)
'*
'* @param VBProject Project
'* @param VBComponent Component
'* @param Integer LineNumber
'* @param Integer LastLine
'* @param String Key
'* @param String Value
'* @param String SourcePath
'* @param String Context
'*
'* @uses VBScript_RegExp_55.RegExp
'**
Public Sub ProcessCodeModuleVBDocFileOutput(Project As VBProject, Component As VBComponent, ByRef LineNumber As Integer, ByVal LastLine As Integer, ByVal Key As String, ByVal Value As String, ByVal SourcePath As String, ByVal Context As String)
    Dim Line As String
    Dim Code As String
    Dim FullPath As String
    Dim hFile As Integer
    Dim bEOF As Boolean
    Dim vbDocRegex As New RegExp
    Dim vbDocRef As Object
    
    Dim sComment As String
    Dim sBlockCloser As String
    Dim bInBlock As Boolean
    Dim sStdPrefix As String
    
    Let sComment = "'"
    Let sBlockCloser = sComment & "**"
    
    Dim vValueMatch
    Dim sFileName As String
    Dim sDesiredContext As String
                    
    Let sDesiredContext = "Export" ' By default
    
    With vbDocRegex
        .IgnoreCase = True
        .Pattern = "^(\S+)(\s+(.*))?$"
        Set vbDocRef = .Execute(sourceString:=Line)
    End With

    Set vbDocRef = vbDocRegex.Execute(sourceString:=Value)
    If vbDocRef.Count > 0 Then
        For Each vValueMatch In vbDocRef
            Let sFileName = vValueMatch.SubMatches(0)
            Let sDesiredContext = vValueMatch.SubMatches(2)
        Next vValueMatch
    Else
        Let sFileName = Value
    End If
                    
    If Context <> sDesiredContext Then
        Exit Sub
    End If
    
    Let FullPath = SourcePath & "\" & sFileName
    
    Let hFile = FreeFile
    Open FullPath For Output As hFile
    
    Do
        LineNumber = LineNumber + 1
        Line = Component.CodeModule.Lines(StartLine:=LineNumber, Count:=1)
        
        ' N.B.: It's best to Roll Your Own on the regular expression matching here
        ' and not rely on anything in modRegularExpressionFunctions or modMyVBA,
        ' which would normally be a more attractive way to implement. Why? Well,
        ' because we might be re-importing modRegularExpressionFunctions or modMyVBA
        ' and in those cases, trying to use any function from either would cause VBA
        ' to pitch a fit.
        With vbDocRegex
            .IgnoreCase = True
            .Pattern = "'[*]?\s*@([A-Za-z]+)(\s+(.*))?\s*$"
            Set vbDocRef = .Execute(sourceString:=Line)
        End With
        
        Let Code = ""
        If vbDocRef.Count > 0 Then
            Dim Match
            
            For Each Match In vbDocRef
                Let Code = Match.SubMatches(0)
            Next Match
        End If
        
        Let bEOF = ((Left(Line, Len(sComment)) <> sComment) Or (Left(Line, Len(sBlockCloser)) = sBlockCloser) Or (Code = "eof"))
        If Not bEOF Then
            Let Line = Right(Line, Len(Line) - Len(sComment))
            
            If Not bInBlock Then
                
                If Left(Line, 2) = "* " Then
                    Let sStdPrefix = Left(Line, 2)
                ElseIf Left(Line, 1) = "*" Then
                    Let sStdPrefix = Left(Line, 1)
                End If
                
                Let bInBlock = True
            
            End If

            If Len(sStdPrefix) > 0 Then
            
                If Left(Line, Len(sStdPrefix)) = sStdPrefix Then
                    Let Line = Right(Line, Len(Line) - Len(sStdPrefix))
                End If
                
            End If
                        
            Print #hFile, Line
        End If
    Loop Until (LineNumber > LastLine) Or bEOF
    
    Close hFile
    
    Let LineNumber = LineNumber - 1
End Sub

Public Sub ProcessCodeModuleVBDocDependency(Project As VBProject, Component As VBComponent, ByVal Line As String, ByVal Key As String, ByVal Value As String, ByVal SourcePath As String, ByVal Context As String)
    Dim Matches As Object
    Dim Match As Object
    
    Dim vbDocRegex As New RegExp: With vbDocRegex
        .Pattern = "^([^{.\s]*)(\s*[.]\s*(\S*))?\s*([{][^}]+[}])?([/@]([0-9]+)[/.]([0-9]+)|@(\S+))?(\s.*)?$"
        .IgnoreCase = True
        Set Matches = vbDocRegex.Execute(sourceString:=Value)
    End With
    
    If (Context <> "Import") Or (Matches.Count = 0) Then
        Debug.Print "Uses: ", Value
    Else
        For Each Match In Matches
            Dim Wotsit
            Dim Parent, ParentGuid, ParentMajor, ParentMinor
            Dim RepositoryURL
            
            With Match
                Let Parent = .SubMatches(0)
                Let Wotsit = .SubMatches(2)
                
                Let ParentGuid = .SubMatches(3)
                Let ParentMajor = .SubMatches(5)
                Let ParentMinor = .SubMatches(6)
                
                Let RepositoryURL = .SubMatches(7)
            End With
            
            Dim Ref
            Dim MyRef As Object
            
            Let Parent = Trim(Parent)
            Set MyRef = Nothing
            For Each Ref In Project.References
                If UCase(Parent) = UCase(Ref.Name) Then
                    Set MyRef = Ref
                    Exit For
                ElseIf ParentGuid = Ref.guid Then
                    Set MyRef = Ref
                    Exit For
                End If
            Next Ref
                            
            If Not MyRef Is Nothing Then
                Debug.Print "Uses: ", Wotsit, Parent, "Fulfilled by: ", MyRef.guid & "/" & MyRef.Major & "." & MyRef.Minor
            ElseIf Len(ParentGuid) > 0 And Len(ParentMajor) > 0 And Len(ParentMinor) > 0 Then
                'Try: add a Reference using the GUID, Major and Minor version documented in the comment
                On Error GoTo AddFromGuid_Borked:
                Project.References.AddFromGuid guid:=ParentGuid, Major:=Val(ParentMajor), Minor:=Val(ParentMinor)
                On Error GoTo 0

                If Len(ParentGuid) > 0 And Len(ParentMajor) > 0 And Len(ParentMinor) > 0 Then
                    Debug.Print "USES(+): ", Wotsit, Parent, "Fulfilled by: ", ParentGuid & "/" & ParentMajor & "/" & ParentMinor
                Else
                    MsgBox "Add a Reference to: " & Parent & Chr$(13) & Chr$(10) & "To provide: " & Wotsit, Title:="Warning: Reference Missing"
                End If
            ElseIf Len(RepositoryURL) > 0 Then
                ' Check to see if we have the needed Component
                If GetVBComponents(Parent).Count > 0 Then
                    Debug.Print "Uses: ", Parent, "(in VBE)"
                Else
                    ImportCodeFromSource Project:=Project, Component:=Parent, SourcePath:=RepositoryURL
                    If GetVBComponents(Parent).Count > 0 Then
                        Debug.Print "USES(+): ", Parent, "from", RepositoryURL
                    Else
                        MsgBox "Import: " & RepositoryURL & Chr$(13) & Chr$(10) & "To provide: " & Parent & "." & Wotsit, Title:="Warning: Module Missing"
                    End If
                End If
            End If
        Next Match
    End If
    
    Exit Sub
    
AddFromGuid_Borked:
    'Blank GUID is used to indicate a failure to add reference by GUID, and trigger the warning for manual action.
    Let ParentGuid = ""
    
    'Censure and MoveOn
    Resume Next

End Sub

'**
'* DiffCode: display a diff between the working copy of code modules in this database
'* and the working copy most recently exported to a version-controlled repository directory
'*
'* @param String Repo
'*
'* @uses FSO
'* @uses CommonRepository
'**
Public Sub DiffCode(Optional ByVal Repo As String)

    Dim RecoveryPath As String
    Dim RecoveryUniqId As String
    
    If Len(Repo) = 0 Then Let Repo = CommonRepository
        
    If Len(Repo) > 0 Then
        Let RecoveryUniqId = FORMAT(Now, "YYYYmmddHHmmss")
        Let RecoveryPath = Repo & "\" & ExImSubDirectory
        If Not FSO.FolderExists(RecoveryPath) Then
            FSO.CreateFolder Path:=RecoveryPath
        End If
    
        Let RecoveryPath = RecoveryPath & "\diff-" & RecoveryUniqId
        If Not FSO.FolderExists(RecoveryPath) Then
            FSO.CreateFolder Path:=RecoveryPath
        End If

        ExportAllCode Path:=RecoveryPath

        Let CommonRepository = Repo
        Shell PathName:=CFG_GIT_BASH_PATH & " -c '" & CFG_DIFFDIR_SCRIPT_PATH & " \'" & ToMinGWPath(Repo) & "\' \'" & ToMinGWPath(RecoveryPath) & "\' \'" & ExImSubDirectory & "\' \'*.xsd\''"
    Else
        MsgBox "Repository directory required."
    End If
End Sub

Public Function ExImSourceFileURL(Component As Variant, ByVal Path As String, Optional ByVal Context) As String
    Dim FQFN As Boolean
    Dim URL As String
    Dim URLPrefix As String
    
    Let URLPrefix = "file:"
    If Left(Path, Len(URLPrefix)) = URLPrefix Then
        Let Path = Right(Path, Len(Path) - Len(URLPrefix))
        Let Path = Replace(Path, Find:="/", Replace:="\")
    End If
    
    If Len(Path) = 0 Then
        If Len(CommonRepository(Context:=Context)) > 0 Then
            If FSO.FolderExists(CommonRepository(Context:=Context)) Then
                Let Path = CommonRepository(Context:=Context)
            End If
        End If
    End If
    
    ' A. Directory, not fully-qualified file name: ends in backslash
    If Right(Path, 1) = "\" Then
        Let FQFN = False
        
    ' B. Directory, not fully-qualified file name: points to existing directory
    ElseIf FSO.FolderExists(Path) Then
        Let FQFN = False
        Let Path = Path & "\"
        
    ' Z. Seems to be a fully-qualified file name
    Else
        Let FQFN = True
    
    End If

    Let URL = Path
    If Not FQFN Then
        Let URL = URL & ExImSourceFileName(Component:=Component, Path:=Path)
    End If
    
    Let ExImSourceFileURL = URL
End Function

Public Function ExImSourceFileName(Component As Variant, Optional ByVal Path As String, Optional ByVal Import As Boolean) As String
    
    Dim Ext As Variant
    
    If TypeName(Component) = "VBComponent" Then
        Let Ext = ExImSourceFileExtension(Component, Import)
        If Len(Ext) > 0 Then
            Let ExImSourceFileName = Component.Name & "." & Ext
        End If

    ElseIf TypeName(Component) = "String" Then
        Dim Exts(1 To 3) As String
        Let Exts(1) = ".bas"
        Let Exts(2) = ".cls"
        Let Exts(3) = ".frm"
        
        For Each Ext In Exts
            If FSO.FileExists(Path & "\" & Component & Ext) Then
                Let ExImSourceFileName = Component & Ext
                Exit For
            End If
        Next Ext
    End If

End Function

Public Function ExImSourceFileExtension(v As Variant, Optional ByVal Import As Boolean) As String
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
    
    Let ExImSourceFileExtension = Ext
End Function

Public Function IsExImSourceFileExtension(v As Variant, Optional ByVal Import As Boolean) As Boolean
    Dim Ext As String

    If TypeName(v) = "String" Then
        Let Ext = v
    ElseIf TypeName(v) = "File" Then
        Let Ext = FSO.GetExtensionName(v)
    End If
    
    Select Case Ext
    Case "cls"
        Let IsExImSourceFileExtension = True
    Case "bas"
        Let IsExImSourceFileExtension = True
    Case "frm"
        Let IsExImSourceFileExtension = True
    End Select

End Function

'**
'* ExImSubDirectory: get a path relative to a given repo parent directory where you can store transient
'* and short-term backup files when carrying on ex-im operations for code modules.
'*
'* @param String Path The repository parent directory you are exporting to and importing from
'* @return String A path relative to the repository parent directory
'**
Public Function ExImSubDirectory(Optional ByVal Path As String) As String
    Let ExImSubDirectory = ".\.vba_exim"
End Function

'* @uses Office.FileDialog      {2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}/2.8
Private Function InputMaybeRepositoryPath(ByVal Path As String, Optional ByVal Title, Optional ByVal InitialFileName, Optional ByVal Context As String) As String
    Dim sFolder As String
    Dim dlgFolder As Office.FileDialog
    
    If Len(Path) > 0 Then
    
        Let sFolder = Path
        
    Else
        Set dlgFolder = Application.FileDialog(msoFileDialogFolderPicker): With dlgFolder
            .Title = IIf(Len(Title) > 0, Title, "Repository Folder")
            .InitialFileName = IIf(Len(InitialFileName) > 0, InitialFileName, CurrentProject.Path) & "\"
        End With
                
        If dlgFolder.Show Then
            Let sFolder = dlgFolder.SelectedItems(1)
        End If
        
    End If
    
    Let InputMaybeRepositoryPath = sFolder
End Function

Public Function GetVBComponents(ByRef Component As Variant) As Collection
    Dim cComponents As New Collection
    Dim oProject As VBProject
    Dim oComponent As VBComponent
    
    If TypeName(Component) = "String" Then
        For Each oProject In Application.VBE.VBProjects
            Set oComponent = Nothing
            On Error Resume Next
            Set oComponent = oProject.VBComponents(Component)
            On Error GoTo 0
            
            If Not oComponent Is Nothing Then
                cComponents.Add oComponent
            End If
            
        Next oProject
    ElseIf TypeName(Component) = "VBComponent" Then
        cComponents.Add Component
    
    ElseIf TypeName(Component) = "VBProject" Then
        For Each oComponent In Component.VBComponents
            cComponents.Add oComponent
        Next oComponent
    
    Else
        Dim Item As Variant
        Dim SubItem As Variant
        Dim Result As Collection
        
        'Try to enumerate it.
        For Each Item In Component
            Set Result = GetVBComponents(Item)
            For Each SubItem In Result
                cComponents.Add SubItem
            Next SubItem
        Next Item
    End If

    Set GetVBComponents = cComponents
End Function

Public Function ToMinGWPath(ByVal Path As String) As String
    Dim WorkingPath As String
    
    Let WorkingPath = Path
    Let WorkingPath = Replace(Path, Find:="\", Replace:="/")
    If Mid(WorkingPath, 2, 1) = ":" Then
        Let WorkingPath = "/" & LCase(Left(WorkingPath, 1)) & Right(WorkingPath, Len(WorkingPath) - 2)
    End If
    Let ToMinGWPath = WorkingPath
End Function
