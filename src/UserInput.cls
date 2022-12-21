VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public fileExtensions As Variant
Public mode As String
Public closeAfter As Boolean
Public openAfter As Boolean
Public singly As Boolean
Public incChanging As Boolean
Public breakChanging As Boolean
Public attachStep As Boolean
Public xlsNeed As XlsNeedMode
Public Translate As Boolean
Public preview As Boolean
Public Export3D As ExportMode
Public useEngNames As Boolean
Public isColoredPdf As Boolean

Private Sub Class_Initialize()
    fileExtensions = GetFileExtensions
    mode = GetChangeMode
    closeAfter = GetCloseAfter
    openAfter = GetOpenAfter
    singly = GetSingly
    incChanging = GetIncChanging
    breakChanging = GetBreakChanging
    attachStep = GetAttachStep
    xlsNeed = GetXlsNeed
    Translate = GetNeedTranslate
    preview = GetNeedPreview
    Export3D = GetExportModel
    useEngNames = GetEngViews(changesProperties(mode))
    isColoredPdf = GetIsColoredPdf
End Sub
