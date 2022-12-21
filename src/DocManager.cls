VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DocManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public drawing As ModelDoc2
Public model As ModelDoc2
Public defaultConf As String
Public defaultDsg As String
Public defaultName As String
Public baseDsg As String
Public miniSign As String
Public changeNum As Integer
Public path As String
Public user As UserInput
Public isXLSNeeded As Boolean

Sub Init(aUser As UserInput)
    Set user = aUser
End Sub

Private Sub Link(aDrawing As ModelDoc2)
    Set drawing = aDrawing
    
    Set model = GetReferencedDocument(drawing, defaultConf)
    
    GetModelProperty defaultDsg, model, defaultConf, pDsg
    defaultDsg = Trim(defaultDsg)
    
    GetModelProperty defaultName, model, defaultConf, pName
    defaultName = Trim(defaultName)
    
    GetDrawingProperty baseDsg, drawing, pBaseDsg
    baseDsg = Trim(baseDsg)
    If baseDsg = "" Then
        baseDsg = CreateBaseDesignation(defaultDsg)
    End If
    
    changeNum = GetNumberChanging(drawing)
    
    GetDrawingProperty miniSign, drawing, "Пометка"
    
    path = gFSO.GetParentFolderName(drawing.GetPathName)
    
    If user.xlsNeed = xlsNoNeed Then
        isXLSNeeded = False
    Else
        Select Case miniSign
            Case "СБ", "AD", "МЧ", "ID", "РСБ", "УЧ", ".AD", ".ID"
                isXLSNeeded = True
            Case Else
                isXLSNeeded = IIf(user.xlsNeed = xlsNeedForNew, (changeNum = 0), False)
        End Select
    End If
End Sub

Sub ExportModel(ByRef abort As Boolean)
    If model.GetType = swDocPART Then
        If miniSign <> "СБ" And miniSign <> "ВО" And miniSign <> "МЧ" Then
            If user.Export3D = exportCurrent Then
                ExtractOneConfiguration drawing, model, defaultConf, abort
            ElseIf user.Export3D = exportLiked Then
                ExtractLikedConfigurations drawing, model, defaultConf, abort
            Else
                MsgBox "Ошибка при экспорте модели", vbCritical
            End If
        End If
    End If
End Sub

Sub TryConvertDoc(aDrawing As ModelDoc2, ByRef abort As Boolean)
    If Not IsDrawing(aDrawing) Or aDrawing.GetPathName = "" Then
        Exit Sub
    End If
    Link aDrawing
    
    ActivateDoc drawing  'only the active drawing converting to dxf/dwg
    ChangeNumberOfChanging drawing, user.incChanging, user.breakChanging
    
    'if need to translate views
    If user.useEngNames Then
        RenameDrawingViewsAndSheets drawing
    End If
    
    'if need then translate
    If user.Translate Then
        TryConvertDoc2 drawing, user.fileExtensions, user.mode, user.openAfter, user.singly, _
                       abort, user.attachStep, True, user.preview
    End If
    
    'russian always
    If Not abort Then
        TryConvertDoc2 drawing, user.fileExtensions, user.mode, user.openAfter, user.singly, _
                       abort, user.attachStep, False, user.preview
    End If
    
    'model copying
    If Not abort And user.Export3D <> exportNone Then
        ExportModel abort
    End If
    
    'xls-specification
    If Not abort And isXLSNeeded Then
        SaveBOMtoXLS GetBOM(drawing), gFSO.BuildPath(path, baseDsg + " " + defaultName)
    End If
    
    SaveThisDoc drawing
    If Not abort And user.closeAfter Then
        swApp.QuitDoc drawing.GetPathName
    End If
End Sub
