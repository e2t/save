Attribute VB_Name = "Tools"
Option Explicit

Sub SaveSetting2(ByRef key As String, ByRef value As String)
    SaveSetting macroName, macroSection, key, value
End Sub

Sub SaveIntSetting(ByRef key As String, value As Integer)
    SaveSetting2 key, str(value)
End Sub

Sub SaveBoolSetting(ByRef key As String, value As Boolean)
    SaveSetting2 key, BoolToStr(value)
End Sub

Function GetSetting2(ByRef key As String) As String
    GetSetting2 = GetSetting(macroName, macroSection, key, "0")
End Function

Function GetBoolSetting(ByRef key As String) As Boolean
    GetBoolSetting = StrToBool(GetSetting2(key))
End Function

Function GetIntSetting(ByRef key As String) As Integer
    GetIntSetting = StrToInt(GetSetting2(key))
End Function

Function StrToInt(ByRef value As String) As Integer
    If IsNumeric(value) Then
        StrToInt = CInt(value)
    Else
        StrToInt = 0
    End If
End Function

Function StrToBool(ByRef value As String) As Boolean
    If IsNumeric(value) Then
        StrToBool = CInt(value)
    Else
        StrToBool = False
    End If
End Function

Function BoolToStr(value As Boolean) As String
    BoolToStr = str(CInt(value))
End Function

Function IsDrawing(ByRef model As ModelDoc2) As Boolean
    IsDrawing = (model.GetType = swDocDRAWING)
End Function

Function IsPart(ByRef model As ModelDoc2) As Boolean
    IsPart = (model.GetType = swDocPART)
End Function

Function IsAssembly(ByRef model As ModelDoc2) As Boolean
    IsAssembly = (model.GetType = swDocASSEMBLY)
End Function

Function InsertSheetName(ByRef filename As String, ByRef sheetname As String) As String
    Dim posDot As Integer
    
    posDot = InStrRev(filename, ".")
    InsertSheetName = Left(filename, posDot - 1) & " - " & sheetname & "." & _
                      Right(filename, Len(filename) - posDot)
End Function

Function SaveDocAs(ByRef doc As ModelDoc2, filename As String, _
                   Optional ByRef data As ExportPdfData = Nothing) As Boolean
    Dim errors As swFileSaveError_e
    Dim warnings As swFileSaveWarning_e
    
    RemoveOldFile filename
    SaveDocAs = AsBool(doc.Extension.SaveAs(filename, swSaveAsCurrentVersion, _
                       swSaveAsOptions_Copy, data, errors, warnings))
End Function

Function TrySaveDocAs(ByRef doc As ModelDoc2, filename As String, _
                      ByRef data As ExportPdfData, ByRef abort As Boolean) As Boolean
    Dim userCode As Integer
    Dim accepted As Boolean
    
    accepted = False
    TrySaveDocAs = False
    While Not accepted And Not abort
        TrySaveDocAs = SaveDocAs(doc, filename, data)
        accepted = TrySaveDocAs
        If Not TrySaveDocAs Then
            userCode = MsgBox("�� ������� ��������� ��������:" & vbNewLine & vbNewLine & _
                       ShortFsName(filename) & vbNewLine & vbNewLine & _
                       IIf(Len(filename) > 259, "��� ����� ������� �������.", "��������, �� ������."), _
                       vbAbortRetryIgnore)
            Select Case userCode
                Case vbAbort
                    abort = True
                Case vbIgnore
                    accepted = True
            End Select
        Else
            swApp.CloseDoc filename
        End If
    Wend
End Function

Function AsBool(value As Boolean) As Boolean
    AsBool = CInt(value)
End Function

Function ShortFsName(ByRef fullname As String) As String
    ShortFsName = Mid(fullname, InStrRev(fullname, "\") + 1, Len(fullname))
End Function

Function IsFileExists(fullname As String, Optional attr As VbFileAttribute = vbNormal) As Boolean
    IsFileExists = CBool(Len(Dir(fullname, attr)))
End Function

Function GetProperty(ByRef value As String, modelMgr As CustomPropertyManager, propertyName As String) As Boolean
    Dim resolvedValue As String
    Dim wasResolved As Boolean
    
    GetProperty = (modelMgr.Get5(propertyName, False, value, resolvedValue, wasResolved) <> _
                   swCustomInfoGetResult_NotPresent)
End Function

Function GetDrawingProperty(ByRef property As String, drawing As DrawingDoc, _
                            propertyName As String) As CustomPropertyManager
    Set GetDrawingProperty = drawing.Extension.CustomPropertyManager("")
    If Not GetProperty(property, GetDrawingProperty, propertyName) Then
        Set GetDrawingProperty = Nothing
    End If
End Function

Function GetModelProperty(ByRef property As String, model As ModelDoc2, conf As String, _
                          propertyName As String) As CustomPropertyManager
    Set GetModelProperty = model.Extension.CustomPropertyManager(conf)
    If Not GetProperty(property, GetModelProperty, propertyName) Then
        Set GetModelProperty = model.Extension.CustomPropertyManager("")
        If Not GetProperty(property, GetModelProperty, propertyName) Then
            Set GetModelProperty = Nothing
        End If
    End If
End Function

Function IsArrayEmpty(anArray As Variant) As Boolean
    Dim i As Integer
    
    On Error Resume Next
        i = UBound(anArray, 1)
    If err.number = 0 Then
        IsArrayEmpty = False
    Else
        IsArrayEmpty = True
    End If
End Function

Function GetFirstSheet(ByRef drawing As DrawingDoc) As sheet
    Dim sheetNames() As String
    
    sheetNames = drawing.GetSheetNames
    Set GetFirstSheet = drawing.sheet(sheetNames(0))
End Function

Function FindDefaultView(drawing As DrawingDoc) As View
    Dim firstSheet As sheet
    Dim nameDefaultView As String
    Dim firstView As View
    
    Set FindDefaultView = Nothing
    Set firstSheet = GetFirstSheet(drawing)
    nameDefaultView = firstSheet.CustomPropertyView
    Set firstView = drawing.GetFirstView.GetNextView 'firstView is Sheet
    
    If Not firstView Is Nothing Then
        Set FindDefaultView = firstView
        Do While FindDefaultView.GetName2 <> nameDefaultView
            Set FindDefaultView = FindDefaultView.GetNextView
            If FindDefaultView Is Nothing Then
                Set FindDefaultView = firstView
                Exit Do
            End If
        Loop
    End If
End Function

Function GetNumberChanging(drawing As DrawingDoc) As Integer
    Dim str As String
    
    GetNumberChanging = 0
    GetDrawingProperty str, drawing, pChanging
    If IsNumeric(str) Then
        GetNumberChanging = CInt(str)
        If GetNumberChanging < 0 Then
            GetNumberChanging = 0
        End If
    End If
End Function

Function GetReferencedDocument(drawing As DrawingDoc, ByRef conf As String) As ModelDoc2
    Dim defaultView As View
    
    Set GetReferencedDocument = Nothing
    conf = ""
    Set defaultView = FindDefaultView(drawing)
    If Not defaultView Is Nothing Then
        Set GetReferencedDocument = defaultView.ReferencedDocument
        If Not GetReferencedDocument Is Nothing Then
            conf = defaultView.ReferencedConfiguration
        End If
    End If
End Function

Sub SetDrawingProperty(ByRef drawing As DrawingDoc, ByRef propertyName As String, ByRef value As String)
    SetProperty drawing.Extension.CustomPropertyManager(""), propertyName, value
End Sub

Sub SetProperty(ByRef mgr As CustomPropertyManager, ByRef propertyName As String, ByRef value As String)
    mgr.Add3 propertyName, swCustomInfoText, value, swCustomPropertyDeleteAndAdd
End Sub

Function SaveThisDoc(doc As ModelDoc2) As Boolean
    Dim errors As swFileSaveError_e
    Dim warnings As swFileSaveWarning_e
    
    SaveThisDoc = AsBool(doc.Save3(swSaveAsOptions_Silent, errors, warnings))
End Function

Function HaveOpenedDocs(ByRef app As Object) As Boolean
    HaveOpenedDocs = (app.GetDocumentCount > 0)
End Function

Sub ActivateDoc(doc As ModelDoc2)
    Dim error As swActivateDocError_e

    swApp.ActivateDoc3 doc.GetPathName, False, swDontRebuildActiveDoc, error  ' if successfull, error = 0
End Sub

Function OpenThisDoc(filename As String) As ModelDoc2
    Dim error As swFileLoadError_e
    Dim warning As swFileLoadWarning_e
    
    Set OpenThisDoc = swApp.OpenDoc6(filename, GetTypeDocument(filename), swOpenDocOptions_Silent, "", error, warning)
End Function

Function GetTypeDocument(filename As String) As swDocumentTypes_e
    Select Case UCase(gFSO.GetExtensionName(filename))
        Case "SLDASM"
            GetTypeDocument = swDocASSEMBLY
        Case "SLDPRT"
            GetTypeDocument = swDocPART
        Case "SLDDRW"
            GetTypeDocument = swDocDRAWING
        Case Else
            GetTypeDocument = swDocNONE
    End Select
End Function

Function RandomString(number As Integer) As String
    Const symbols As String = "abcdefghijklmnopqrstuvwxyz" & "0123456789"
    
    Randomize
    While number > 0
        RandomString = RandomString + Mid(symbols, Int(Len(symbols) * Rnd() + 1), 1)
        number = number - 1
    Wend
End Function

Function RemoveOldFile(filename As String) As Boolean
    On Error GoTo ErrorKill
    If IsFileExists(filename) Then
        Kill filename
    End If
    RemoveOldFile = True
    Exit Function
ErrorKill:
    RemoveOldFile = False
End Function

Function GetDirOfActiveDoc() As String
    Dim doc As ModelDoc2
    
    Set doc = swApp.ActiveDoc
    If Not doc Is Nothing Then
        GetDirOfActiveDoc = gFSO.GetParentFolderName(doc.GetPathName) + "\"
    End If
End Function

Function CreateBaseDesignation(Designation As String) As String
    Dim LastFullstopPosition As Integer
    Dim FirstHyphenPosition As Integer
    
    CreateBaseDesignation = Designation
    LastFullstopPosition = InStrRev(Designation, ".")
    If LastFullstopPosition > 0 Then
        FirstHyphenPosition = InStr(LastFullstopPosition, Designation, "-")
        If FirstHyphenPosition > 0 Then
            CreateBaseDesignation = Left(Designation, FirstHyphenPosition - 1)
        End If
    End If
End Function

Sub ChangeNumberOfChanging(ByRef drawing As DrawingDoc, incChanging As Boolean, breakChanging As Boolean)
    Dim number As Integer

    number = GetNumberChanging(drawing)
    If incChanging Then
        number = number + 1
    ElseIf breakChanging Then
        number = 0
    End If
    If incChanging Or breakChanging Then
        SetDrawingProperty drawing, pChanging, StrNumberOfChanging(number)
    End If
End Sub

Function GetDrawingNameWOext(drawingName As String) As String
    GetDrawingNameWOext = Left(drawingName, Len(drawingName) - 7)
End Function

Function StrNumberOfChanging(number As Integer)
    StrNumberOfChanging = Format(number, "00")
End Function

Function RevNoUnlessZero(number As Integer) As String
    Const revLabel = "rev"
    
    If number > 0 Then
        RevNoUnlessZero = " (" & revLabel & "." & StrNumberOfChanging(number) & ")"
    Else
        RevNoUnlessZero = ""
    End If
End Function

Sub RenameDrawingViewsAndSheets(drawing As DrawingDoc)
    Dim arraySheets As Variant
    Dim ss As Variant
    Dim vv As Variant
    Dim aView As View
    
    arraySheets = drawing.GetViews
    For Each ss In arraySheets
        For Each vv In ss
            Set aView = vv
            TranslateView aView
        Next
    Next
End Sub

Function SaveDrawingAsPDF(drawing As DrawingDoc, pdfname As String, _
                          data As IExportPdfData, ByRef abort As Boolean) As Boolean
    SaveDrawingAsPDF = TrySaveDocAs(drawing, pdfname, data, abort)
    RemoveMetadataFromPDF pdfname
End Function

