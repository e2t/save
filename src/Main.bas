Attribute VB_Name = "Main"
Option Explicit

Public Const macroName As String = "Save3"
Public Const macroSection As String = "Main"

Enum ForAllMode
    forActive = 0
    forAllOpened
    forAllInFolder
End Enum

Enum XlsNeedMode
    xlsNeedForAll = 0
    xlsNeedForNew
    xlsNoNeed  'must be last always
End Enum

Enum ExportMode
    exportNone = 0
    exportCurrent
    exportLiked
End Enum
Public namesExportMode(2) As String

Enum LanguageMode
    langRussian = 0
    langPoland
End Enum
Public TRname(1) As String

Public Const maxPathLength As Integer = 255
Public swApp As Object
Public namesXlsNeedMode(xlsNoNeed) As String
Private useEngNames As Boolean
Private Const pChanging As String = "Изменение"
Public gFSO As FileSystemObject

Sub Main()
    Set swApp = Application.SldWorks
    InitAll
    MyForm.Show
End Sub

Function InitAll()  'mask fot button

    configFullFileName = swApp.GetCurrentMacroPathFolder + "\Modes.ini"

    namesXlsNeedMode(xlsNeedForAll) = "Для всех чертежей"
    namesXlsNeedMode(xlsNeedForNew) = "Только для новых"
    namesXlsNeedMode(xlsNoNeed) = "Не создавать"
    
    TRname(langRussian) = "Technical Requirements"
    TRname(langPoland) = "Technical Requirements Poland"
    
    namesExportMode(exportNone) = "Не копировать"
    namesExportMode(exportCurrent) = "Одну конфигурацию"
    namesExportMode(exportLiked) = "С похожим обозначением"
    
    Set gFSO = New FileSystemObject
    
    InitializeProperties

End Function

Sub ConvertDocs(aForAllMode As ForAllMode)

   Dim doc_ As Variant
   Dim doc As ModelDoc2
   Dim fileExtensions() As String
   Dim mode As String
   Dim closeAfter As Boolean
   Dim openAfter As Boolean
   Dim singly As Boolean
   Dim incChanging As Boolean
   Dim breakChanging As Boolean
   Dim abort As Boolean
   Dim attachStep As Boolean
   Dim xlsNeed As XlsNeedMode
   Dim Translate As Boolean
   Dim preview As Boolean
   Dim Export3D As ExportMode
   Dim oldUseEnglishLang As Boolean
   Dim isColoredPdf As Boolean
   Dim oldIsColoredPdf As Boolean

   If HaveOpenedDocs(swApp) Then
      fileExtensions = GetFileExtensions
      mode = GetChangeMode
      closeAfter = GetCloseAfter
      openAfter = GetOpenAfter
      singly = GetSingly
      incChanging = GetIncChanging
      breakChanging = GetBreakChanging
      abort = False
      attachStep = GetAttachStep
      xlsNeed = GetXlsNeed
      Translate = GetNeedTranslate
      preview = GetNeedPreview
      Export3D = GetExportModel
      useEngNames = GetEngViews(changesProperties(mode))
      isColoredPdf = GetIsColoredPdf
      
      oldIsColoredPdf = swApp.GetUserPreferenceToggle(swPDFExportInColor)
      swApp.SetUserPreferenceToggle swPDFExportInColor, isColoredPdf
      
      If useEngNames Then
         oldUseEnglishLang = swApp.GetUserPreferenceToggle(swUseEnglishLanguageFeatureNames)
         swApp.SetUserPreferenceToggle swUseEnglishLanguageFeatureNames, True
      End If
      
      If aForAllMode = forAllOpened Then
         For Each doc_ In swApp.GetDocuments
            Set doc = doc_
            TryConvertDoc doc, fileExtensions, mode, closeAfter, openAfter, singly, _
                          incChanging, breakChanging, abort, attachStep, xlsNeed, _
                          Translate, preview, Export3D
            If abort Then Exit For
         Next
          
      ElseIf aForAllMode = forAllInFolder Then
         Dim file_ As Variant
         Dim file As Object
         Dim dirOfActiveDoc As String
         Dim folder As Object
         Dim filename As String
         Dim maybeCloseAfter As Boolean
         Dim err As swFileLoadError_e
         Dim wrn As swFileLoadWarning_e
         
         Set folder = gFSO.GetFolder(GetDirOfActiveDoc)
         For Each file_ In folder.Files
            Set file = file_
            filename = LCase(file.Path)
            If InStr(filename, "slddrw") > 0 And InStr(filename, "~$") = 0 Then
               Set doc = swApp.OpenDoc6(filename, swDocDRAWING, swOpenDocOptions_Silent, "", err, wrn)
               If wrn = swFileLoadWarning_AlreadyOpen Then
                  maybeCloseAfter = closeAfter
               Else
                  maybeCloseAfter = True
               End If
               
               ''' job
               TryConvertDoc doc, fileExtensions, mode, maybeCloseAfter, openAfter, singly, _
                             incChanging, breakChanging, abort, attachStep, xlsNeed, _
                             Translate, preview, Export3D
               Set doc = Nothing
               If abort Then Exit For
               ''' end job
            End If
         Next
          
      Else
         TryConvertDoc swApp.ActiveDoc, fileExtensions, mode, closeAfter, openAfter, singly, _
                       incChanging, breakChanging, abort, attachStep, xlsNeed, Translate, _
                       preview, Export3D
      End If
       
      If useEngNames Then
         swApp.SetUserPreferenceToggle swUseEnglishLanguageFeatureNames, oldUseEnglishLang
      End If
      swApp.SetUserPreferenceToggle swPDFExportInColor, oldIsColoredPdf
      Unload MyForm
   Else
      MsgBox "Не открытых документов.", vbCritical
   End If

End Sub

Function CloseNonDrawings() 'mask for button   'MAYBE to remove

    Dim doc_ As Variant
    Dim doc As ModelDoc2

    For Each doc_ In swApp.GetDocuments
        Set doc = doc_
        If Not IsDrawing(doc) Then
            swApp.CloseDoc doc.GetPathName
        End If
    Next

End Function

Sub TryConvertDoc(ByRef doc As ModelDoc2, ByRef fileExtensions() As String, mode As String, _
                  closeAfter As Boolean, openAfter As Boolean, singly As Boolean, _
                  incChanging As Boolean, breakChanging As Boolean, ByRef abort As Boolean, _
                  attachStep As Boolean, xlsNeed As XlsNeedMode, Translate As Boolean, _
                  preview As Boolean, Export3D As ExportMode)
                  
    If Not IsDrawing(doc) Or doc.GetPathName = "" Then Exit Sub
    ActivateDoc doc  'only the active doc converting to dxf/dwg
    ChangeNumberOfChanging doc, incChanging, breakChanging
    
    'if need to translate views
    If useEngNames Then
        RenameDrawingViewsAndSheets doc
    End If
    
    'if need then translate
    If Translate Then
        TryConvertDoc2 doc, fileExtensions, mode, openAfter, singly, _
                       abort, attachStep, True, preview
    End If
    
    'russian always
    If Not abort Then
        TryConvertDoc2 doc, fileExtensions, mode, openAfter, singly, _
                       abort, attachStep, False, preview
    End If
    
    'model copying
    If Not abort And Export3D <> exportNone Then
        ExportModel doc, abort, Export3D
    End If
    
    'xls-specification
    If Not abort And IsNeedXLS(doc, xlsNeed) Then
        SaveBOMtoXLS GetBOM(doc), GetDrawingNameWOext(doc.GetPathName)
    End If
    
    SaveThisDoc doc
    If Not abort And closeAfter Then
        swApp.QuitDoc doc.GetPathName
    End If
End Sub

Sub TryConvertDoc2(ByRef doc As ModelDoc2, ByRef fileExtensions() As String, mode As String, _
                   openAfter As Boolean, singly As Boolean, ByRef abort As Boolean, _
                   attachStep As Boolean, Translate As Boolean, preview As Boolean)
    Dim propMgrs As Dictionary
    Dim oldValues As Dictionary
    Dim tr As Note
    Dim oldTextTR As String
    
    If Translate Then
        ChangeTR doc, tr, oldTextTR, preview, abort
    End If
    If Not abort Then
        ReserveAllProperties propMgrs, oldValues, doc
        ChangeAllProperties propMgrs, mode
        doc.ForceRebuild3 True
        
        MultiSaveDrawing doc, fileExtensions, openAfter, singly, abort, attachStep, Translate
        
        If Translate Then
            RestoreOldTR tr, oldTextTR
        End If
        RestoreAllProperties propMgrs, oldValues
        doc.ForceRebuild3 True
    End If
End Sub

Sub ChangeNumberOfChanging(ByRef drawing As DrawingDoc, incChanging As Boolean, breakChanging As Boolean)

    Dim number As Integer

    number = GetNumberChanging(drawing)
    If incChanging Then
        number = number + 1
    ElseIf breakChanging Then
        number = 0
    End If
    If incChanging Or breakChanging Then
        SetDrawingProperty drawing, pChanging, str(number)
    End If

End Sub

Function MultiSaveDrawing(ByRef drawing As DrawingDoc, ByRef fileExtensions() As String, _
                          openAfter As Boolean, singly As Boolean, ByRef abort As Boolean, _
                          attachStep As Boolean, Translate As Boolean) As Boolean

    Dim ext_ As Variant
    Dim ext As String
    
    MultiSaveDrawing = True
    If Not IsArrayEmpty(fileExtensions) Then
        For Each ext_ In fileExtensions
            ext = ext_
            MultiSaveDrawing = MultiSaveDrawing And SaveDrawingAs(drawing, ext, openAfter, singly, abort, attachStep, Translate)
            If abort Then Exit For
        Next
    End If

End Function

Function IsNeedXLS(drawing As DrawingDoc, xlsNeed As XlsNeedMode) As Boolean
    
   Dim miniSign As String
   
   If xlsNeed = xlsNoNeed Then
      IsNeedXLS = False
   Else
      GetDrawingProperty miniSign, drawing, "Пометка"
      Select Case miniSign
         Case "СБ", "AD", "МЧ", "ID", "РСБ", "УЧ"
            IsNeedXLS = True
         Case Else
            If xlsNeed = xlsNeedForNew Then
               IsNeedXLS = (GetNumberChanging(drawing) = 0)
            Else
               IsNeedXLS = False
            End If
      End Select
   End If

End Function

Function SaveDrawingAs(ByRef drawing As DrawingDoc, ByRef fileExtension As String, _
                       openAfter As Boolean, singly As Boolean, ByRef abort As Boolean, _
                       attachStep As Boolean, Translate As Boolean) As Boolean

    Dim data As IExportPdfData
    Dim sheetname_ As Variant
    Dim sheetname As String
    Dim specifiedSheet(0) As String
    Dim filename As String
    Dim pdfnames() As String
    Dim pdfname_ As Variant
    Dim pdfname As String
    Dim isPdf As Boolean
    Dim i As Integer
    
    SaveDrawingAs = True
    isPdf = (fileExtension = "PDF")
    filename = NewFilename(drawing, fileExtension, Translate)  ''TODO
    If isPdf Then
        Set data = swApp.GetExportFileData(1)
        ReDim pdfnames(drawing.GetSheetCount - 1)
        i = 0
        If singly And drawing.GetSheetCount > 1 Then  'по отдельным листам
            For Each sheetname_ In drawing.GetSheetNames
                sheetname = sheetname_
                specifiedSheet(0) = sheetname
                data.SetSheets swExportData_ExportSpecifiedSheets, specifiedSheet
                SaveDrawingAs = SaveDrawingAs And _
                                TrySaveDocAs(drawing, InsertSheetName(filename, sheetname), data, abort)
                pdfnames(i) = InsertSheetName(filename, sheetname)
                i = i + 1
                If abort Then Exit For
            Next
        Else
            data.SetSheets swExportData_ExportAllSheets, drawing.GetSheetNames
            SaveDrawingAs = TrySaveDocAs(drawing, filename, data, abort)
            pdfnames(0) = filename
        End If
        Set data = Nothing
        
        If attachStep Then
            AttachModelToPDF drawing, pdfnames(0), abort
        End If
        If openAfter Then
            For Each pdfname_ In pdfnames
                pdfname = pdfname_
                CreateObject("WScript.Shell").Run """" & pdfname & """", vbHide, False
            Next
        End If
    Else
        SaveDrawingAs = TrySaveDocAs(drawing, filename, Nothing, abort)
    End If

End Function

Sub AttachModelToPDF(drawing As DrawingDoc, pdfname As String, ByRef abort As Boolean)
    Dim stepName As String
    Dim sheetname_ As Variant
    Dim sheetname As String
    Dim asheet As sheet
    Dim aView_ As Variant
    Dim aView As View
    Dim confs As Dictionary
    Dim pair_ As Variant
    Dim pair As PairModelConf
    Dim key As String
    Dim propSign As String
    Dim propName As String
    
    Set confs = New Dictionary
    For Each sheetname_ In drawing.GetSheetNames
        sheetname = sheetname_
        Set asheet = drawing.sheet(sheetname)
        For Each aView_ In asheet.GetViews
            Set aView = aView_
            If Not aView.ReferencedDocument Is Nothing Then  'OPTIMIZE: просматриваются лишние виды
                key = aView.GetReferencedModelName & "::" & aView.ReferencedConfiguration
                If Not confs.Exists(key) Then
                    Set pair = New PairModelConf
                    pair.conf = aView.ReferencedConfiguration
                    Set pair.model = aView.ReferencedDocument
                    confs.Add key, pair
                End If
            End If
        Next
    Next
    For Each pair_ In confs.Items
        Set pair = pair_
        GetModelProperty propSign, pair.model, pair.conf, "Обозначение"
        GetModelProperty propName, pair.model, pair.conf, "Наименование"
        stepName = propSign & " " & propName & " (" & pair.conf & ").STEP"
        pair.model.ShowConfiguration2 pair.conf
        TrySaveDocAs pair.model, stepName, Nothing, abort
        swApp.CloseDoc pair.model.GetPathName
        AttachFileToPDF pdfname, stepName
        Kill stepName
    Next
End Sub

Sub AttachFileToPDF(pdfname As String, attachname As String)
    CreateObject("WScript.Shell").Run """" & swApp.GetCurrentMacroPathFolder & "\cpdf.exe"" -attach-file """ & _
                                      attachname & """ """ & pdfname & """ -o """ & pdfname & """", _
                                      vbHide, True
End Sub

Function GetDrawingNameWOext(ByRef drawingName As String) As String

    GetDrawingNameWOext = Left(drawingName, Len(drawingName) - 7)

End Function

Function FormatNumberOfChanging(number As Integer) As String
    Dim revLabel As String
    revLabel = IIf(useEngNames, "rev", "изм")
    
    If number > 0 Then
        FormatNumberOfChanging = " (" & revLabel & "." & Format(number, "00") & ")"
    Else
        FormatNumberOfChanging = ""
    End If
End Function

Function NewFilename(ByRef drawing As DrawingDoc, ByRef fileExtension As String, Translate As Boolean) As String

    Dim drawingName As String
    Dim number As Integer
    
    drawingName = drawing.GetPathName
    number = GetNumberChanging(drawing)
    NewFilename = GetDrawingNameWOext(drawingName) & _
                  FormatNumberOfChanging(number) & _
                  IIf(Translate, " - POLAND", "") & _
                  "." & fileExtension
End Function

Function GetNumberChanging(ByRef drawing As DrawingDoc) As Integer
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

Sub RenameDrawingViewsAndSheets(drawing As DrawingDoc)
    Dim arraySheets As Variant
    arraySheets = drawing.GetViews
    Dim ss
    For Each ss In arraySheets
        Dim vv
        For Each vv In ss
            Dim aView As View
            Set aView = vv
            TranslateView aView
        Next
    Next
End Sub

Sub TranslateView(aView As View)
    Dim viewName As String
    viewName = aView.GetName2
    
    'Местные виды и разрезы переименовать нельзя.
    'Они меняют название в зависимости от настроек SolidWorks.
    
    Dim search_words(1) As String
    search_words(0) = "лист"
    search_words(1) = "чертежный вид"
    
    Dim eng_words(1) As String
    eng_words(0) = "Sheet"
    eng_words(1) = "Drawing View"
    
    Dim regex As RegExp
    Set regex = New RegExp
    regex.IgnoreCase = True
    
    Dim i As Integer
    For i = LBound(search_words) To UBound(search_words)
        regex.Pattern = search_words(i)
        If regex.Test(viewName) Then
            Dim newViewName As String
            newViewName = regex.Replace(viewName, eng_words(i)) & " "
            aView.SetName2 newViewName
        End If
    Next
End Sub

Function ExitApp() 'mask

   Unload MyForm
   End

End Function
