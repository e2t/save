Attribute VB_Name = "Main"
Option Explicit

Public Const macroName = "Save3"
Public Const macroSection = "Main"
Public Const pBaseDsg = "Базовое обозначение"
Public Const pDsg = "Обозначение"
Public Const pName = "Наименование"
Public Const pNameEN = "Наименование EN"
Public Const pNamePL = "Наименование PL"
Public Const pNameUA = "Наименование UA"
Public Const pNameLT = "Наименование LT"
Public Const pChanging = "Изменение"
Public Const maxPathLength = 255

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

Public swApp As Object
Public namesXlsNeedMode(xlsNoNeed) As String
Public gFSO As FileSystemObject
Public pAllNames(4) As String

Dim search_words(1) As String
Dim eng_words(1) As String

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
    
    'Местные виды и разрезы переименовать нельзя.
    'Они меняют название в зависимости от настроек SolidWorks.
    search_words(0) = "лист"
    search_words(1) = "чертежный вид"
    
    eng_words(0) = "Sheet"
    eng_words(1) = "Drawing View"
    
    pAllNames(0) = pName
    pAllNames(1) = pNameEN
    pAllNames(2) = pNamePL
    pAllNames(3) = pNameUA
    pAllNames(4) = pNameLT
    
    Set gFSO = New FileSystemObject
    
    InitializeProperties
    
End Function

Sub ConvertDocs(aForAllMode As ForAllMode)

   Dim doc_ As Variant
   Dim doc As ModelDoc2
   Dim abort As Boolean
   Dim oldIsColoredPdf As Boolean
   Dim oldUseEnglishLang As Boolean
   Dim user As UserInput
   Dim docMgr As DocManager

   If HaveOpenedDocs(swApp) Then
      Set user = New UserInput
      abort = False

      oldIsColoredPdf = swApp.GetUserPreferenceToggle(swPDFExportInColor)
      swApp.SetUserPreferenceToggle swPDFExportInColor, user.isColoredPdf
      
      If user.useEngNames Then
         oldUseEnglishLang = swApp.GetUserPreferenceToggle(swUseEnglishLanguageFeatureNames)
         swApp.SetUserPreferenceToggle swUseEnglishLanguageFeatureNames, True
      End If
      
      Set docMgr = New DocManager
      docMgr.Init user
      
      If aForAllMode = forAllOpened Then
         For Each doc_ In swApp.GetDocuments
            Set doc = doc_
            docMgr.TryConvertDoc doc, abort
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
            filename = LCase(file.path)
            If InStr(filename, "slddrw") > 0 And InStr(filename, "~$") = 0 Then
               Set doc = swApp.OpenDoc6(filename, swDocDRAWING, swOpenDocOptions_Silent, "", err, wrn)
               If wrn = swFileLoadWarning_AlreadyOpen Then
                  maybeCloseAfter = user.closeAfter
               Else
                  maybeCloseAfter = True
               End If
               
               ''' job
               docMgr.TryConvertDoc doc, abort
               Set doc = Nothing
               If abort Then Exit For
               ''' end job
            End If
         Next
          
      Else
         docMgr.TryConvertDoc swApp.ActiveDoc, abort
      End If
       
      If user.useEngNames Then
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

Sub TryConvertDoc2(doc As ModelDoc2, fileExtensions As Variant, mode As String, _
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

Function MultiSaveDrawing(drawing As DrawingDoc, fileExtensions As Variant, _
                          openAfter As Boolean, singly As Boolean, ByRef abort As Boolean, _
                          attachStep As Boolean, Translate As Boolean) As Boolean
    Dim ext_ As Variant
    Dim ext As String
    
    MultiSaveDrawing = True
    If Not IsArrayEmpty(fileExtensions) Then
        For Each ext_ In fileExtensions
            ext = ext_
            MultiSaveDrawing = MultiSaveDrawing And SaveDrawingAs( _
                drawing, ext, openAfter, singly, abort, attachStep, Translate)
            If abort Then Exit For
        Next
    End If
End Function

Function SaveDrawingAs(drawing As DrawingDoc, fileExtension As String, _
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
    Dim I As Integer
    
    SaveDrawingAs = True
    isPdf = (fileExtension = "PDF")
    filename = NewFilename(drawing, fileExtension, Translate)  ''TODO
    If isPdf Then
        Set data = swApp.GetExportFileData(1)
        ReDim pdfnames(drawing.GetSheetCount - 1)
        I = 0
        If singly And drawing.GetSheetCount > 1 Then  'по отдельным листам
            For Each sheetname_ In drawing.GetSheetNames
                sheetname = sheetname_
                specifiedSheet(0) = sheetname
                data.SetSheets swExportData_ExportSpecifiedSheets, specifiedSheet
                SaveDrawingAs = SaveDrawingAs And _
                                SaveDrawingAsPDF(drawing, InsertSheetName(filename, sheetname), data, abort)
                pdfnames(I) = InsertSheetName(filename, sheetname)
                I = I + 1
                If abort Then Exit For
            Next
        Else
            data.SetSheets swExportData_ExportAllSheets, drawing.GetSheetNames
            SaveDrawingAs = SaveDrawingAsPDF(drawing, filename, data, abort)
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
        GetModelProperty propSign, pair.model, pair.conf, pDsg
        GetModelProperty propName, pair.model, pair.conf, pName
        stepName = propSign & " " & propName & " (" & pair.conf & ").STEP"
        pair.model.ShowConfiguration2 pair.conf
        TrySaveDocAs pair.model, stepName, Nothing, abort
        swApp.CloseDoc pair.model.GetPathName
        AttachFileToPDF pdfname, stepName
        Kill stepName
    Next
End Sub

Sub AttachFileToPDF(pdfname As String, attachname As String)
    Dim cpdf As String
    Dim rewrite As String
    
    cpdf = """" & swApp.GetCurrentMacroPathFolder & "\cpdf.exe"" "
    rewrite = " -o """ & pdfname & """ """ & pdfname & """"
    
    CreateObject("WScript.Shell").Run cpdf & "-attach-file """ & attachname & """" & rewrite, vbHide, True
End Sub

Sub RemoveMetadataFromPDF(pdfname As String)
    Dim cpdf As String
    Dim rewrite As String
    
    cpdf = """" & swApp.GetCurrentMacroPathFolder & "\cpdf.exe"" "
    rewrite = " -o """ & pdfname & """ """ & pdfname & """"
    
    CreateObject("WScript.Shell").Run cpdf & "-remove-metadata" & rewrite, vbHide, True
    CreateObject("WScript.Shell").Run cpdf & "-set-author """"" & rewrite, vbHide, True
    CreateObject("WScript.Shell").Run cpdf & "-set-creator """"" & rewrite, vbHide, True
    CreateObject("WScript.Shell").Run cpdf & "-set-producer """"" & rewrite, vbHide, True
End Sub

Function NewFilename(drawing As DrawingDoc, fileExtension As String, Translate As Boolean) As String
    Dim drawingName As String
    Dim number As Integer
    
    drawingName = drawing.GetPathName
    number = GetNumberChanging(drawing)
    NewFilename = GetDrawingNameWOext(drawingName) & _
                  RevNoUnlessZero(number) & _
                  IIf(Translate, " - POLAND", "") & _
                  "." & fileExtension
End Function

Sub TranslateView(aView As View)
    Dim viewName As String
    Dim regex As RegExp
    Dim I As Integer
    Dim newViewName As String
    
    viewName = aView.GetName2
    Set regex = New RegExp
    regex.IgnoreCase = True
    For I = LBound(search_words) To UBound(search_words)
        regex.Pattern = search_words(I)
        If regex.Test(viewName) Then
            newViewName = regex.Replace(viewName, eng_words(I)) & " "
            aView.SetName2 newViewName
        End If
    Next
End Sub

Function ExitApp() 'mask
   Unload MyForm
   End
End Function
