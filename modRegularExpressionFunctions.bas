Attribute VB_Name = "modRegularExpressionFunctions"
Option Explicit

'**
'* RegexMatch: Return True if the given string value matches the given Regex pattern
'*
'* @param Variant value Value to check for a regular expression match
'* @param String pattern Regular expression pattern
'* @param Boolean MatchCase If True, letters in the pattern must match case (/a/ matches "a", not "A")
'*      If False or omitted, letters in the pattern match across upper/lowercase (/a/ matches "a" or "A")
'
'* @return Boolean True if the given string value matches the given Regex pattern, False otherwise
'**
Public Function RegexMatch(Value As Variant, Pattern As String, Optional ByVal MatchCase As Boolean) As Boolean
    If IsNull(Value) Then Exit Function
    ' Using a static, we avoid re-creating the same regex object for every call '
    Static regex As Object
    ' Initialise the Regex object '
    If regex Is Nothing Then
        Set regex = CreateObject("vbscript.regexp")
        With regex
            .Global = True
            .MultiLine = True
        End With
    End If

    With regex
            .IgnoreCase = Not MatchCase
    End With

    ' Update the regex pattern if it has changed since last time we were called '
    If regex.Pattern <> Pattern Then regex.Pattern = Pattern
    ' Test the value against the pattern '
    RegexMatch = regex.Test(Value)
End Function

'**
'* RegexComponent: get the contents of a numbered back-reference component
'* provided the given string value matches the given Regex pattern
'*
'* @param Variant value Value to check for a regular expression match
'* @param String pattern Regular expression pattern to match it against
'* @param Integer Part Number of the sub-pattern to return the matching contents for, beginning with 1 for the first ($1)
'* @param Boolean MatchCase If True, letters in the pattern must match case (/a/ matches "a", not "A")
'*      If False or omitted, letters in the pattern match across upper/lowercase (/a/ matches "a" or "A")
'*
'* @return String The contents of the matching back-reference, or an empty string if there is no match.
'**
Public Function RegexComponent(Value As Variant, Pattern As String, Part As Variant, Optional ByVal MatchCase As Boolean, Optional NamedGroups As Scripting.Dictionary) As String
    Dim Components As Scripting.Dictionary
    
    Set Components = RegexComponents(Value:=Value, Pattern:=Pattern, Match:=0, MatchCase:=MatchCase, NamedGroups:=NamedGroups)
    
    If Not Components Is Nothing Then
        If Components.Exists(Part) Then
            Let RegexComponent = Components.Item(Part)
        End If
    End If
    
    Set Components = Nothing
End Function

Public Function RegexComponents(Value As Variant, Pattern As String, Optional ByVal Match As Integer, Optional ByVal MatchCase As Boolean, Optional NamedGroups As Scripting.Dictionary) As Scripting.Dictionary
    Dim cMatches As Object
    Dim iMatch As Variant
    Dim iSubMatch As Variant
    Dim Index As Integer
    Dim iMatchIndex As Integer
    Dim dComponents As Scripting.Dictionary
    
    If IsNull(Value) Then Exit Function
    
    ' Using a static, we avoid re-creating the same regex object for every call '
    Static regex As Object
    
    ' Initialise the Regex object '
    If regex Is Nothing Then
        Set regex = CreateObject("vbscript.regexp")
        With regex
            .Global = True
            .MultiLine = True
        End With
    End If
    
    With regex
        .IgnoreCase = Not MatchCase
    End With
    
    ' Update the regex pattern if it has changed since last time we were called '
    If regex.Pattern <> Pattern Then regex.Pattern = Pattern
    ' Test the value against the pattern '
    Set cMatches = regex.Execute(Value)
    
    Let iMatchIndex = 0
    For Each iMatch In cMatches
        If iMatchIndex = Match Then
            Set dComponents = New Scripting.Dictionary
            
            '$0 = whole regex match
            dComponents.Add Item:=iMatch.Value, Key:=0
            
            Let Index = 0
            For Each iSubMatch In iMatch.SubMatches
                Let Index = Index + 1
                
                '$1, $2, $3, ... = components matched
                dComponents.Add Item:=iSubMatch, Key:=Index
                
                If Not NamedGroups Is Nothing Then
                    If NamedGroups.Exists(Index) Then
                        dComponents.Add Item:=iSubMatch, Key:=NamedGroups.Item(Index)
                    End If
                End If
            Next iSubMatch
        End If
        
        Let iMatchIndex = iMatchIndex + 1
    Next iMatch
    
    Set RegexComponents = dComponents
End Function

Public Function RegexReplace(Value As Variant, Pattern As String, Replace As String, Optional ByVal MatchCase As Boolean, Optional ByVal OnlyOne As Boolean) As String
    ' Using a static, we avoid re-creating the same regex object for every call '
    Static regex As RegExp
    
    Dim hasMatch As Boolean
    Dim Result As String
    
    If IsNull(Value) Then Exit Function
    
    ' Initialise the Regex object '
    If regex Is Nothing Then
        Set regex = New RegExp
        With regex
            .MultiLine = True
        End With
    End If
    
    Result = CStr(Value)
    With regex
        .Pattern = Pattern
        .Global = Not OnlyOne
        .IgnoreCase = Not MatchCase
    End With
            
    RegexReplace = regex.Replace(Result, Replace)

End Function

Public Function RegexSplit(ByVal Text As String, ByVal Pattern As String, Optional ByVal MatchCase As Boolean, Optional ByVal DelimCapture As Boolean) As String()
    Dim aWords() As String
    Dim sReplace As String
    
    If DelimCapture Then
        Let sReplace = Chr$(0) & "$1" & Chr$(0)
    Else
        Let sReplace = Chr$(0)
    End If
    
    Let Text = RegexReplace(Value:=Text, Pattern:="(" & Pattern & ")", Replace:=sReplace, MatchCase:=MatchCase)
    Let aWords = Split(Expression:=Text, Delimiter:=Chr$(0))
    Let RegexSplit = aWords
End Function
