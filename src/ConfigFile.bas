Attribute VB_Name = "ConfigFile"
Option Explicit

Public configFullFileName As String

Function GetRowsFromFile() As Boolean
    Dim objStream As Stream
        
    Set objStream = New Stream
    objStream.Charset = "utf-8"
    objStream.Open
    GetRowsFromFile = False
    
    On Error GoTo CreateConfig
    objStream.LoadFromFile configFullFileName
    GoTo SuccessRead

ReadConfigAgain:
    On Error GoTo ExitFunction
    objStream.LoadFromFile configFullFileName
    GoTo SuccessRead
   
SuccessRead:
    ReadRowsFromFile objStream
    GetRowsFromFile = True
ExitFunction:
    objStream.Close
    Set objStream = Nothing
    Exit Function
    
CreateConfig:
    CreateDefaultConfigFile objStream
    GoTo ReadConfigAgain
End Function

Sub ReadRowsFromFile(objStream As Stream)
    Dim line As String
    Dim regexItem As RegExp
    Dim regexMode As RegExp
    Dim header As String
    Dim key As String
    Dim value As String
    Dim properties As Dictionary
    Dim aMatch As MatchCollection
    
    Set regexItem = New RegExp
    regexItem.IgnoreCase = True
    regexItem.Global = True
    regexItem.Pattern = "\s*(\S+)(.*)"
    
    Set regexMode = New RegExp
    regexMode.IgnoreCase = True
    regexMode.Global = True
    regexMode.Pattern = "\s*\[(.*\S.*)\]\s*"
    
    Do Until objStream.EOS
        line = objStream.ReadText(adReadLine)
        
        If regexMode.Test(line) Then
            Set aMatch = regexMode.Execute(line)
            header = Trim(aMatch(0).SubMatches(0))
            AddItemInDict header, New Dictionary, changesProperties
        ElseIf regexItem.Test(line) And changesProperties.Exists(header) Then
            Set aMatch = regexItem.Execute(line)
            key = aMatch(0).SubMatches(0)
            Set properties = changesProperties(header)
            
            If key = UseEnglishNames Then
                AddItemInDict UseEnglishNames, True, properties
            Else
                value = Trim(aMatch(0).SubMatches(1))
                AddItemInDict key, value, properties
            End If
        End If
    Loop
End Sub

Sub AddItemInDict(key As Variant, item As Variant, ByRef dict As Dictionary)
    If dict.Exists(key) Then
        dict(key) = item
    Else
        dict.Add key, item
    End If
End Sub

Sub CreateDefaultConfigFile(objStream As Stream)
    'TODO: check if cannot to create file
    objStream.WriteText _
        "# Свойства: " + pDevel + ", " + pDraft + ", " + pCheck + ", " + pTech + ", " + pNorm + ", " + pAppr + ", " + pFirm + vbNewLine + _
        "# " + UseEnglishNames + " переименовывает виды и листы на английский язык" + vbNewLine + _
        vbNewLine
    objStream.WriteText _
        "[По умолчанию]" + vbNewLine + _
        pCheck + vbNewLine + _
        pTech + vbNewLine + _
        pNorm + vbNewLine + _
        pAppr + vbNewLine + _
        pFirm + vbNewLine + _
         vbNewLine
    objStream.WriteText _
        "[Польша]" + vbNewLine + _
        pFirm + " ООО ""Эко-Инвест""" + vbNewLine + _
         vbNewLine
    objStream.WriteText _
        "[English]" + vbNewLine + _
        UseEnglishNames + vbNewLine + _
        pCheck + " Urikov" + vbNewLine + _
        pTech + " Gumennyj" + vbNewLine + _
        pNorm + " Urikov" + vbNewLine + _
        pAppr + " Gumennyj" + vbNewLine + _
        pFirm + " Ekoton Industrial Group" + vbNewLine
    objStream.SaveToFile configFullFileName
End Sub

