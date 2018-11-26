Attribute VB_Name = "modMyVBA"
Option Explicit

' No VT_GUID available so must declare type GUID
Private Type GUID_TYPE
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (guid As GUID_TYPE) As LongPtr
Private Declare PtrSafe Function StringFromGUID2 Lib "ole32.dll" (guid As GUID_TYPE, ByVal lpStrGuid As LongPtr, ByVal cbMax As Long) As LongPtr

Public Const EX_NOHOOK = (vbObjectError + 255)
Public Const EX_RENAMEFAILED = (vbObjectError + 254)
Public Const EX_ACCNSCAN_DUPLICATE = (vbObjectError + 253)

Public Const EX_FILEALREADYEXISTS = 58
Public Const EX_FILEPERMISSIONDENIED = 75
Public Const EX_INVALIDPROPERTYVALUE = 380
Public Const EX_DUPLICATE_KEY_VALUE = 3022

Public Const COLOR_ALERT As Long = &H10FFFF         'Bright yellow
Public Const COLOR_DISABLED As Long = &HC0C0C0      'Light grey
Public Const COLOR_UNMARKED As Long = &HFFFFFF      'White
Public Const COLOR_MARKEDERROR As Long = &HC0C0FF   'Light red

'**
'* DebugDump: Utility Function mainly for use in the Immediate pane to more easily display a
'* bunch of different kinds of objects and collections of objects in VBA
'*
'* @param Variant v The object to print out a representation of in the Immediate pane
'**
Sub DebugDump(v As Variant)
    Dim vScalar As Variant
    If IsArray(v) Or TypeName(v) = "Collection" Or TypeName(v) = "ISubMatches" Or TypeName(v) = "Fields" Then
        If IsArray(v) Then
            Debug.Print TypeName(v), LBound(v), UBound(v)
        Else
            Debug.Print TypeName(v), v.Count
        End If
        
        For Each vScalar In v
            DebugDump (vScalar)
        Next vScalar
    ElseIf TypeName(v) = "Dictionary" Then
        Debug.Print TypeName(v)
        For Each vScalar In v.Keys
            Debug.Print vScalar & ":"
            DebugDump v.Item(vScalar)
        Next vScalar
    ElseIf TypeName(v) = "Nothing" Then
        Debug.Print TypeName(v)
    Else
        Debug.Print TypeName(v), v
    End If
End Sub

Public Function Merge(ParamArray Value() As Variant) As Collection
    Dim I As Long
    Dim Item As Variant
    
    Set Merge = New Collection
    
    For I = LBound(Value) To UBound(Value)
        For Each Item In Value(I)
            Merge.Add Item
        Next Item
    Next I
End Function

Public Function Coll(ParamArray Value() As Variant) As Collection
    Dim I As Long
    
    Set Coll = New Collection
    For I = LBound(Value) To UBound(Value)
        Coll.Add Item:=Value(I)
    Next I
End Function

Public Function Assoc(ParamArray KeyValue() As Variant) As Dictionary
    Dim I As Long
    Dim Step As Integer
    Dim vKey As Variant
    Dim vValue As Variant
    
    Set Assoc = New Dictionary
    Let I = LBound(KeyValue)
    Do Until I > UBound(KeyValue)
        Let vKey = Null
        Let vValue = Null
        
        If TypeName(KeyValue(I)) = "Collection" Then
            Let Step = 1
            Let vKey = KeyValue(I).Item(1)
            Let vValue = KeyValue(I).Item(2)
        Else
            Let Step = 2
            Let vKey = KeyValue(I)
            If I + 1 <= UBound(KeyValue) Then
                Let vValue = KeyValue(I + 1)
            End If
        End If
        
        Assoc.Add Key:=vKey, Item:=vValue
        
        Let I = I + Step
    Loop
End Function

Public Function ListOfItems(List As Variant) As Variant
    Dim vItem As Variant
    
    Select Case TypeName(List)
    Case "Collection":
        Set ListOfItems = List 'Collection
    Case "Dictionary":
        Set ListOfItems = New Collection
        For Each vItem In List.Items 'Variant()
            ListOfItems.Add Item:=vItem
        Next vItem
    Case Else:
        If IsObject(List) Then
            Set ListOfItems = List
        ElseIf IsArray(List) Then
            Set ListOfItems = New Collection
            For Each vItem In List 'Variant()
                ListOfItems.Add Item:=vItem
            Next vItem
        Else
            Let ListOfItems = List
        End If
    End Select
End Function

Public Function ListSum(List As Variant, Optional ByVal Initial As Variant) As Variant
    Dim vSum As Variant
    Dim vItem As Variant
    Dim vList As Variant
    
    If IsMissing(Initial) Then
        Let Initial = 0#
    End If
    
    Set vList = ListOfItems(List)
    Let vSum = Initial
    For Each vItem In vList
        Let vSum = vSum + vItem
    Next vItem
    
    Let ListSum = vSum
End Function

Public Sub BubbleSortList(ByRef List As Variant)
    Dim Swapped As Boolean
    Dim vSwap As Variant
    Dim I As Integer, J As Integer
    
    If IsArray(List) Or TypeName(List) = "Collection" Then
        Do
            Let Swapped = False
            For I = LBound(List) To UBound(List) - 1
                If List(I + 1) < List(I) Then
                    Let vSwap = List(I)
                    Let List(I) = List(I + 1)
                    Let List(I + 1) = vSwap
                    
                    Let Swapped = True
                End If
            Next I
        Loop Until Not Swapped
    End If
End Sub

'**
'* Join: join an iterable list into a string, with items separated by a given delimiter
'* (for example ["One", "Two", "Red", "Blue"] => "One;Two;Red;Blue")
'*
'* @param String Delimiter The characters used to separate items from the list (e.g.: ", ")
'* @param Variant List The items to join into a single string (can be any list iterable with For Each) (e.g. Array of ("One", "Two", "Red", "Blue"))
'* @return String containing all the items from List, separated by Delimiter (e.g. "One, Two, Red, Blue")
'**
Public Function Join(ByVal Delimiter As String, List As Variant) As String
    Dim First As Boolean
    Dim vItem As Variant
    Dim sConjunction As String
    
    First = True
    For Each vItem In List
        If Not First Then
            sConjunction = sConjunction & Delimiter
        End If
        
        sConjunction = sConjunction & vItem
        
        First = False
    Next vItem
    
    Join = sConjunction
End Function

Public Function camelCase(ByVal Words As Variant, Optional ByVal FilterOut As String, Optional ByVal InitialLower As Boolean) As String
    Dim s As String
    Dim Word As Variant
    Dim Item As Variant
    Dim WordList As Variant
    
    If TypeName(Words) = "String" Then
        Let WordList = RegexSplit(Text:=Words, Pattern:="\s+")
    Else
        Set WordList = New Collection
        For Each Item In Words
            For Each Word In RegexSplit(Text:=Item, Pattern:="\s+")
                WordList.Add Word
            Next Word
        Next Item
    End If
    
    Let s = ""
    For Each Word In WordList
        If Len(FilterOut) > 0 Then
            Let Word = RegexReplace(Value:=Word, Pattern:=FilterOut, Replace:="")
        End If
        Let Word = TitleCase(Text:=Word, ForceLower:=True)
        Let s = s & Word
    Next Word
    
    If InitialLower Then
        Let s = LCase(Left(s, 1)) & Right(s, Len(s) - 1)
    End If
    
    If TypeName(WordList) = "Collection" Then
        Set WordList = Nothing
    End If
    
    Let camelCase = s
End Function

'**
'* camelCaseSplitString: split a CamelCase string into its apparent component words
'* as marked by the shifts in case
'*
'* @param String s The camelCase text to split into words
'*
'* @return Collection of String items for each word
'**
Public Function camelCaseSplitString(ByVal s As String) As Collection
    Dim isAlpha As New RegExp
    Dim isUpper As New RegExp
    Dim isLower As New RegExp
    Dim isUpperLower As New RegExp
    Dim isWhiteSpace As New RegExp
    
    With isUpper
        .IgnoreCase = False
        .Pattern = "^([A-Z])$"
    End With
    
    With isLower
        .IgnoreCase = False
        .Pattern = "^([a-z])$"
    End With
    
    With isAlpha
        .IgnoreCase = False
        .Pattern = "^([A-Za-z]+)$"
    End With

    With isUpperLower
        .IgnoreCase = False
        .Pattern = "^([A-Z][a-z])$"
    End With
    
    With isWhiteSpace
        .IgnoreCase = False
        .Pattern = "^((\s|[_])+)$"
    End With

    
    Dim cWords As New Collection
    Dim c0 As String, C As String, c2 As String
    Dim I0 As Integer, I As Integer
    Dim Anchor As Integer
    Dim State As Integer
    
    Anchor = 0
    I = 1
    GoTo NextWord
    
    'Finite State Machine
NextWord:
    If I > Len(s) Then
        GoTo ExitMachine
    End If
    
    c0 = Mid(s, I, 1)
    If isUpper.Test(c0) Then
        Anchor = I
        GoTo WordBeginsOnUpper
    ElseIf isLower.Test(c0) Then
        Anchor = I
        GoTo WordBeginsOnLower
    ElseIf isWhiteSpace.Test(c0) Then
        Anchor = I
        Let I = I + Len(C)
        GoTo NextWord
    Else
        Anchor = I
        GoTo FromOtherToNextWord
    End If

WordBeginsOnLower:
    If I > Len(s) Then GoTo NextWord
    Let c0 = C: Let C = Mid(s, I, 1)
    
    Let I = I + Len(C)
    GoTo ContinueWordToUpperBreak

ContinueWordToUpperBreak:
    If I > Len(s) Then GoTo NextWord
    Let c0 = C: Let C = Mid(s, I, 1)
    
    If isUpper.Test(C) Or isWhiteSpace.Test(C) Then
        GoTo ClipWord
    Else
        I = I + Len(C)
    End If
    GoTo ContinueWordToUpperBreak
    
WordBeginsOnUpper:
    If I > Len(s) Then GoTo NextWord
    Let c0 = C: C = Mid(s, I, 1)

    'Move ahead to the next character
    'UPPERCase: two uppers in a row
    'MixedCase: one upper, one lower
    Let I = I + 1
    Let c0 = C: C = Mid(s, I, 1)
    
    If isLower.Test(C) Then
        Let I = I + Len(C)
        GoTo ContinueWordToUpperBreak
    ElseIf isUpper.Test(C) Then
        GoTo ContinueWordToUpperLowerBreak
    ElseIf Not isWhiteSpace.Test(C) Then
        GoTo ContinueWordToUpperLowerBreak
    Else
        GoTo ClipWord
    End If
    
ContinueWordToUpperLowerBreak:
    If I > Len(s) Then GoTo NextWord
    Let c0 = C: C = Mid(s, I, 1): c2 = Mid(s, I, 2)

    If isUpperLower.Test(c2) Or isWhiteSpace.Test(C) Then
        GoTo ClipWord
    ElseIf isAlpha.Test(C) Then
        I = I + 1
        GoTo ContinueWordToUpperLowerBreak
    Else
        I = I + 1
        GoTo ContinueWordToUpperLowerBreak
    End If
    
FromOtherToNextWord:
    If I > Len(s) Then GoTo NextWord
    C = Mid(s, I, 1)
    
    If isUpper.Test(C) Or isLower.Test(C) Or isWhiteSpace.Test(C) Then
        GoTo ClipWord
    Else
        I = I + 1
        GoTo FromOtherToNextWord
    End If

ClipWord:
    If Anchor < I Then
        cWords.Add Mid(s, Anchor, I - Anchor)
    End If
    GoTo NextWord

ExitMachine:
    If Anchor > 0 Then
        cWords.Add Mid(s, Anchor, I - Anchor)
    End If
       
    Set camelCaseSplitString = cWords
End Function

'**
'* TitleCase: convert a string into Title Case (capitalized first character;
'* preserve case for the rest of the word).
'*
'* @param String s The camelCase text to split into words
'*
'* @return Collection of String items for each word
'**
Public Function TitleCase(ByVal Text As String, Optional ByVal ForceLower As Boolean)
    Dim Word As String
    Dim aWords() As String
    Dim I As Integer
    Dim sOutput As String
    Dim First As String, Rest As String
    
    Let aWords = RegexSplit(Text:=Text, Pattern:="\s", DelimCapture:=True)
    Let I = LBound(aWords)
    Do Until I > UBound(aWords)
        Let Word = aWords(I)
        
        Let First = Left(Word, 1)
        If First >= "a" And First <= "z" Then
            Let First = UCase(First)
        End If
        
        Let Rest = Right(Word, Len(Word) - 1)
        If ForceLower Then
            Let Rest = LCase(Rest)
        End If
        
        Let Word = First & Rest
        
        Let sOutput = sOutput & Word
        
        'Next!
        Let I = I + 1
    Loop
    
    Let TitleCase = sOutput
End Function

Public Function FindSlugInDirectory(Directory As String, Slug As String, Optional ByVal MaxDepth As Variant) As String
    Dim sPath As String, sFoundFile As String
    Dim vSubDirectory As Variant
    Dim cSubDirectories As New Collection
    Dim FS As New FileSystemObject
    
    If IsMissing(MaxDepth) Then
        Let MaxDepth = -1
    End If
    
    If FS.FolderExists(Directory & "\" & Slug) Or FS.FileExists(Directory & "\" & Slug) Then
        Let FindSlugInDirectory = Directory & "\" & Slug
    Else
        Let sPath = Dir(PathName:=Directory & "\*.*", Attributes:=vbDirectory)
        Do Until Len(sPath) = 0
            If Not RegexMatch(sPath, "^[.]+$") Then
                If FS.FolderExists(Directory & "\" & sPath) Then
                    cSubDirectories.Add Item:=sPath
                End If
            End If
            Let sPath = Dir(Attributes:=vbDirectory)
        Loop
        
        For Each vSubDirectory In cSubDirectories
            Let sPath = CStr(vSubDirectory)
            If MaxDepth <> 0 Then
                Let sFoundFile = FindSlugInDirectory(Directory:=Directory & "\" & sPath, Slug:=Slug, MaxDepth:=MaxDepth - 1)
                If Len(sFoundFile) > 0 Then
                    Exit For
                End If
            End If
        Next vSubDirectory
        
        Let FindSlugInDirectory = sFoundFile
    End If
End Function

Public Function CreateGuidString()
    Dim guid As GUID_TYPE
    Dim strGuid As String
    Dim retValue As LongPtr
    Const guidLength As Long = 39 'registry GUID format with null terminator {xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}

    retValue = CoCreateGuid(guid)
    If retValue = 0 Then
        strGuid = String$(guidLength, vbNullChar)
        retValue = StringFromGUID2(guid, StrPtr(strGuid), guidLength)
        If retValue = guidLength Then
            ' valid GUID as a string
            CreateGuidString = strGuid
        End If
    End If
End Function

