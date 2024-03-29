Attribute VB_Name = "WithGui"
Option Explicit

Enum IncBoxMode
    ShowThisNumber
    ShowNextNumber
    ShowAdding
    NonShow
End Enum

Function GetFileExtensions() As String()
    Dim size As Integer
    Dim i As Integer
    
    size = IIf(MyForm.pdfBox.value, 1, 0) + _
           IIf(MyForm.dwgBox.value, 1, 0) + _
           IIf(MyForm.dxfBox.value, 1, 0) + _
           IIf(MyForm.tifBox.value, 1, 0) + _
           IIf(MyForm.psdBox.value, 1, 0)
    If size > 0 Then
        ReDim GetFileExtensions(size - 1)
        i = 0
        AddExtension GetFileExtensions, i, "PDF", MyForm.pdfBox.value
        AddExtension GetFileExtensions, i, "DWG", MyForm.dwgBox.value
        AddExtension GetFileExtensions, i, "DXF", MyForm.dxfBox.value
        AddExtension GetFileExtensions, i, "TIF", MyForm.tifBox.value
        AddExtension GetFileExtensions, i, "PSD", MyForm.psdBox.value
    End If
End Function

Sub AddExtension(ByRef exts() As String, ByRef index As Integer, ByRef fileExtension As String, _
                 condition As Boolean)
    If condition Then
        exts(index) = fileExtension
        index = index + 1
    End If
End Sub

Function GetChangeMode() As String
    GetChangeMode = MyForm.changeBox.text
End Function

Function HaveTRinThisDoc() As Boolean
    Dim thisDoc As ModelDoc2
    
    HaveTRinThisDoc = False
    Set thisDoc = swApp.ActiveDoc
    If Not thisDoc Is Nothing Then
        If IsDrawing(thisDoc) Then
            HaveTRinThisDoc = Not FindExistingTR(thisDoc, TRname(langRussian)) Is Nothing Or _
                              Not FindExistingTR(thisDoc, TRname(langPoland)) Is Nothing
        End If
    End If
End Function

Function GetNumberChangingOfThisDoc() As Integer  ' mask for button
    Dim thisDoc As ModelDoc2
    
    GetNumberChangingOfThisDoc = 0
    Set thisDoc = swApp.ActiveDoc
    If Not thisDoc Is Nothing Then
        If IsDrawing(thisDoc) Then
            GetNumberChangingOfThisDoc = GetNumberChanging(thisDoc)
        End If
    End If
End Function

Sub SetIncBoxCaption(mode As IncBoxMode)
    Dim text As String
    Dim number As Integer
    
    text = IIf(MyForm.IsForAll, "���������� �������", "���������� ������")
    Select Case mode
        Case ShowThisNumber
            text = text & RevNoUnlessZero(GetNumberChangingOfThisDoc)
        Case ShowNextNumber
            text = text & RevNoUnlessZero(GetNumberChangingOfThisDoc + 1)
        Case ShowAdding
            text = text & " (+1)"
    End Select
    MyForm.incBox.Caption = text
End Sub
                 
Function InitSettings()  'mask for button
    Dim name_ As Variant
    Dim name As String
    Dim rowChangeBox As Integer

    With MyForm
        For Each name_ In changesProperties.Keys
            name = name_
            .changeBox.AddItem name
        Next
        For Each name_ In namesXlsNeedMode
            name = name_
            .xlsNeedBox.AddItem name
        Next
        rowChangeBox = GetIntSetting("changemode")
        If rowChangeBox <= .changeBox.ListCount - 1 Then
            .changeBox.ListIndex = rowChangeBox
        Else
            .changeBox.ListIndex = 0
        End If
        'after creating Form
        .scaleBox.value = AsBool(swApp.GetUserPreferenceIntegerValue(swDxfOutputNoScale))
        .pdfBox.value = GetBoolSetting("pdf")
        .dwgBox.value = GetBoolSetting("dwg")
        .dxfBox.value = GetBoolSetting("dxf")
        .tifBox.value = GetBoolSetting("tif")
        .psdBox.value = GetBoolSetting("psd")
        .closeBox.value = GetBoolSetting("close")
        .openBox.value = GetBoolSetting("open")
        .singlyBox.value = GetBoolSetting("singly")
        .colorBox.value = GetBoolSetting("colorpdf")
        .attachBox.value = GetBoolSetting("attach")
        .xlsNeedBox.ListIndex = GetIntSetting("xlsneed")
        .chkTranslate.value = GetBoolSetting("translate")
        .chkPreview.value = GetBoolSetting("preview")
        Select Case GetIntSetting("exportmodel")
            Case exportNone
                .radExportNone.value = True
            Case exportCurrent
                .radExportCurrent.value = True
            Case exportLiked
                .radExportLiked.value = True
        End Select
        ChangeCaptions
    End With
End Function

Function ChangeCaptions()  'mask for button
    With MyForm
        .breakBox.Caption = IIf(.IsForAll, "����� �������", "����� ������")
        If .incBox.value Then
            SetIncBoxCaption IIf(.IsForAll, ShowAdding, ShowNextNumber)
        ElseIf .breakBox.value Then
            SetIncBoxCaption NonShow
        Else
            SetIncBoxCaption IIf(.IsForAll, NonShow, ShowThisNumber)
        End If
        If Not HaveTRinThisDoc Then
            .chkTranslate.Caption = .chkTranslate.Caption & " (���)"
        End If
        .radExportCurrent.Caption = namesExportMode(exportCurrent)
        .radExportLiked.Caption = namesExportMode(exportLiked)
        .radExportNone.Caption = namesExportMode(exportNone)
    End With
End Function

Function GetCloseAfter() As Boolean
    GetCloseAfter = MyForm.closeBox.value
End Function

Function GetOpenAfter() As Boolean
    GetOpenAfter = MyForm.openBox.value
End Function

Function GetSingly() As Boolean
    GetSingly = MyForm.singlyBox.value
End Function

Function GetIncChanging() As Boolean
    GetIncChanging = MyForm.incBox.value
End Function

Function GetBreakChanging() As Boolean
    GetBreakChanging = MyForm.breakBox.value
End Function

Function GetAttachStep() As Boolean
    GetAttachStep = MyForm.attachBox.value
End Function

Function GetXlsNeed() As XlsNeedMode
    GetXlsNeed = MyForm.xlsNeedBox.ListIndex
End Function

Function GetNeedTranslate() As Boolean
    GetNeedTranslate = MyForm.chkTranslate.value
End Function

Function GetNeedPreview() As Boolean
    GetNeedPreview = MyForm.chkPreview.value
End Function

Function GetIsColoredPdf() As Boolean
   GetIsColoredPdf = MyForm.colorBox.value
End Function

Function GetExportModel() As ExportMode
    If MyForm.radExportCurrent.value Then
        GetExportModel = exportCurrent
    ElseIf MyForm.radExportLiked.value Then
        GetExportModel = exportLiked
    ElseIf MyForm.radExportNone.value Then
        GetExportModel = exportNone
    End If
End Function

Function GetEngViews(properties As Dictionary) As Boolean
    GetEngViews = properties.Exists(UseEnglishNames)
End Function
