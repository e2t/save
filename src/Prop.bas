Attribute VB_Name = "Prop"
'Written in 2015 by Eduard E. Tikhenko <aquaried@gmail.com>
'
'To the extent possible under law, the author(s) have dedicated all copyright
'and related and neighboring rights to this software to the public domain
'worldwide. This software is distributed without any warranty.
'You should have received a copy of the CC0 Public Domain Dedication along
'with this software.
'If not, see <http://creativecommons.org/publicdomain/zero/1.0/>

Option Explicit

Private drawingProperties(5) As String
Private modelProperties(0) As String
Private changesProperties(modeAsIs - 1) As Object

Function InitializeProperties()  'mask for button

    Dim defaultProperties As Dictionary
    Dim drokinProperties As Dictionary
    Dim ekotonProperties As Dictionary
    Dim polandProperties As Dictionary
    Const pDevel As String = "Разработал"
    Const pDraft As String = "Начертил"
    Const pCheck As String = "Проверил"
    Const pTech As String = "Техконтроль"
    Const pNorm As String = "Нормоконтроль"
    Const pAppr As String = "Утвердил"
    Const pFirm As String = "Организация"
    
    'properties of model
    modelProperties(0) = pDevel
    'properties of drawing
    drawingProperties(0) = pDraft
    drawingProperties(1) = pCheck
    drawingProperties(2) = pTech
    drawingProperties(3) = pNorm
    drawingProperties(4) = pAppr
    drawingProperties(5) = pFirm
    'changes by default
    Set defaultProperties = New Dictionary
    defaultProperties.Add pCheck, ""
    defaultProperties.Add pTech, ""
    defaultProperties.Add pNorm, ""
    defaultProperties.Add pAppr, ""
    defaultProperties.Add pFirm, ""
    'changes for Drokin
    Set drokinProperties = New Dictionary
    drokinProperties.Add pDevel, "Горбенко"
    drokinProperties.Add pDraft, "Горбенко"
    drokinProperties.Add pCheck, "Григоров"
    drokinProperties.Add pTech, ""
    drokinProperties.Add pNorm, ""
    drokinProperties.Add pAppr, "Дрокин"
    drokinProperties.Add pFirm, "ИП ""Дрокин"""
    'changes for Ekoton
    Set ekotonProperties = New Dictionary
    ekotonProperties.Add pDevel, "Вершинина"
    ekotonProperties.Add pDraft, "Холодов"
    ekotonProperties.Add pCheck, ""
    ekotonProperties.Add pTech, ""
    ekotonProperties.Add pNorm, ""
    ekotonProperties.Add pAppr, "Ватта"
    ekotonProperties.Add pFirm, "ЗАО НПФ ""Экотон"""
    'changes for Poland
    Set polandProperties = New Dictionary
    polandProperties.Add pFirm, "ООО ""Эко-Инвест"""
    'all variants of the changes
    Set changesProperties(modeDefault) = defaultProperties
    Set changesProperties(modeDrokin) = drokinProperties
    Set changesProperties(modeEkoton) = ekotonProperties
    Set changesProperties(modePoland) = polandProperties

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

Sub ChangeAllProperties(ByRef mgrs As Dictionary, mode As ChangeMode)
    
    If mode <> modeAsIs Then
        ChangePropertiesBy mgrs, changesProperties(mode)
    End If

End Sub

Sub RestoreAllProperties(ByRef mgrs As Dictionary, ByRef values As Dictionary)

    ChangePropertiesBy mgrs, values
    
End Sub

Sub ChangePropertiesBy(ByRef mgrs As Dictionary, ByRef changes As Dictionary)

    Dim propName_ As Variant
    Dim propName As String

    For Each propName_ In changes.Keys
        propName = propName_
        If mgrs.Exists(propName) Then  'TODO: property skips if it dont exists. Need to create and to remove after
            SetProperty mgrs(propName), propName, changes(propName)
        End If
    Next

End Sub
