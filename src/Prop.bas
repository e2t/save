Attribute VB_Name = "Prop"
Option Explicit

Private drawingProperties(5) As String
Private modelProperties(0) As String
Public changesProperties As Dictionary

Public Const UseEnglishNames = "@UseEnglishNames"
Public Const pDevel = "Разработал"
Public Const pDraft = "Начертил"
Public Const pCheck = "Проверил"
Public Const pTech = "Техконтроль"
Public Const pNorm = "Нормоконтроль"
Public Const pAppr = "Утвердил"
Public Const pFirm = "Организация"

Function InitializeProperties()  'mask for button
    'properties of model
    modelProperties(0) = pDevel
    
    'properties of drawing
    drawingProperties(0) = pDraft
    drawingProperties(1) = pCheck
    drawingProperties(2) = pTech
    drawingProperties(3) = pNorm
    drawingProperties(4) = pAppr
    drawingProperties(5) = pFirm
    
    'all variants of the changes
    Set changesProperties = New Dictionary
    changesProperties.Add "Без изменений", New Dictionary
    GetRowsFromFile
End Function

Sub ReserveAllProperties(ByRef propMgrs As Dictionary, ByRef values As Dictionary, _
                         ByRef drawing As DrawingDoc)
    Dim propName_ As Variant
    Dim propName As String
    Dim model As ModelDoc2
    Dim conf As String
    Dim mgr As CustomPropertyManager
    Dim propValue As String

    Set propMgrs = New Dictionary
    Set values = New Dictionary
    For Each propName_ In drawingProperties
        propName = propName_
        Set mgr = GetDrawingProperty(propValue, drawing, propName)
        If Not mgr Is Nothing Then
            propMgrs.Add propName, mgr
            values.Add propName, propValue
        End If
    Next
    Set model = GetReferencedDocument(drawing, conf)
    If Not model Is Nothing Then
        For Each propName_ In modelProperties
            propName = propName_
            Set mgr = GetModelProperty(propValue, model, conf, propName)
            If Not mgr Is Nothing Then
                propMgrs.Add propName, mgr
                values.Add propName, propValue
            End If
        Next
    End If
End Sub

Sub ChangeAllProperties(ByRef mgrs As Dictionary, mode As String)
    ChangePropertiesBy mgrs, changesProperties(mode)
End Sub

Sub RestoreAllProperties(ByRef mgrs As Dictionary, ByRef values As Dictionary)
    ChangePropertiesBy mgrs, values
End Sub

Sub ChangePropertiesBy(ByRef mgrs As Dictionary, ByRef changes As Dictionary)
    Dim propName_ As Variant
    Dim propName As String

    For Each propName_ In changes.Keys
        If propName_ = UseEnglishNames Then GoTo NextFor
        propName = propName_
        If mgrs.Exists(propName) Then  'TODO: property skips if it dont exists. Need to create and to remove after
            SetProperty mgrs(propName), propName, changes(propName)
        End If
NextFor:
    Next
End Sub
