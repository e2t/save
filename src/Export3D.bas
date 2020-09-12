Attribute VB_Name = "Export3D"
'Written in 2016 by Eduard E. Tikhenko <aquaried@gmail.com>
'
'To the extent possible under law, the author(s) have dedicated all copyright
'and related and neighboring rights to this software to the public domain
'worldwide. This software is distributed without any warranty.
'You should have received a copy of the CC0 Public Domain Dedication along
'with this software.
'If not, see <http://creativecommons.org/publicdomain/zero/1.0/>

Option Explicit

Sub ExportModel(drawing As ModelDoc2, ByRef abort As Boolean, Export3D As ExportMode)
    Dim model As ModelDoc2
    Dim newname As String
    Dim remark As String
    Dim curconf As String
    Dim defView As View
        
    Set defView = FindDefaultView(drawing)
    Set model = defView.ReferencedDocument
    curconf = defView.ReferencedConfiguration
    If model.GetType = swDocPART Then
        GetDrawingProperty remark, drawing, "Пометка"
        If remark <> "СБ" And remark <> "ВО" And remark <> "МЧ" Then
            If Export3D = exportCurrent Then
                ExtractOneConfiguration drawing, model, curconf, abort
            ElseIf Export3D = exportLiked Then
                ExtractLikedConfigurations drawing, model, curconf, abort
            Else
                MsgBox "Ошибка при экспорте модели", vbCritical
            End If
        End If
    End If
End Sub

Sub ExtractLikedConfigurations(drawing As ModelDoc2, model As ModelDoc2, _
                               curconf As String, ByRef abort As Boolean)
    Dim propSign As String
    Dim propName As String
    Dim basePropSign As String
    Dim x As Variant
    Dim conf As String
    Dim newname As String
        
    basePropSign = ExtractBaseSign(GetPropertySign(model, curconf))
    For Each x In model.GetConfigurationNames
        conf = x
        If Not conf Like "*SM-FLAT-PATTERN" Then
            propSign = GetPropertySign(model, conf)
            If ExtractBaseSign(propSign) = basePropSign Then
                propName = GetPropertyName(model, conf)
                newname = ExportedFilename(drawing, model.GetPathName, propSign, propName)
                TrySaveDocAs model, newname, Nothing, abort
                If Not abort Then
                    PurgeConfigurations newname, conf, propSign, propName
                End If
            End If
        End If
    Next
End Sub

Sub ExtractOneConfiguration(drawing As ModelDoc2, model As ModelDoc2, _
                            conf As String, ByRef abort As Boolean)
    Dim propSign As String
    Dim propName As String
    Dim newname As String
    
    propSign = GetPropertySign(model, conf)
    propName = GetPropertyName(model, conf)
    newname = ExportedFilename(drawing, model.GetPathName, propSign, propName)
    TrySaveDocAs model, newname, Nothing, abort
    If Not abort Then
        PurgeConfigurations newname, conf, propSign, propName
    End If
End Sub

Sub PurgeConfigurations(newModelName As String, savingConf As String, _
                        propSign As String, propName As String)
    Dim newModel As ModelDoc2
    Dim x As Variant
    Dim confname As String
    Dim conf As Configuration
    
    Set newModel = OpenThisDoc(newModelName) ' may be NULL
    newModel.ShowConfiguration2 savingConf
    Set conf = newModel.AddConfiguration3(RandomString(10), "", "", 0)
    For Each x In newModel.GetConfigurationNames
        confname = x
        newModel.DeleteConfiguration2 confname
    Next
    conf.name = propSign & " " & propName
    SaveThisDoc newModel
    swApp.CloseDoc newModelName
End Sub

Function ExtractBaseSign(propSign As String) As String
    Dim hyphen As Integer
    Dim i As Integer
    
    For i = Len(propSign) To 1 Step -1
        If Mid(propSign, i, 1) = "." Then
            Exit For
        End If
        If Mid(propSign, i, 1) = "-" Then
            hyphen = i
            Exit For
        End If
    Next
    If hyphen > 0 Then
        ExtractBaseSign = Left(propSign, hyphen - 1)
    Else
        ExtractBaseSign = propSign
    End If
End Function

Function ExportedFilename(drawing As ModelDoc2, modelname As String, _
                          propSign As String, propName As String) As String
    ExportedFilename = _
        ExportedDirectory(drawing) & "\" & _
        propSign & " " & propName & _
        " - Copy." & gFSO.GetExtensionName(modelname)
End Function

Function ExportedDirectory(drawing As ModelDoc2) As String
    Const suffix As String = "Копии моделей"
    Dim newFolder As String
    
    newFolder = gFSO.GetParentFolderName(drawing.GetPathName) & "\" & suffix
    If Not gFSO.FolderExists(newFolder) Then
        gFSO.CreateFolder newFolder
    End If
    ExportedDirectory = newFolder
End Function

Private Function GetPropertyName(model As ModelDoc2, conf As String) As String
    Dim value As String
    
    GetModelProperty value, model, conf, "Наименование"
    GetPropertyName = value
End Function

Private Function GetPropertySign(model As ModelDoc2, conf As String) As String
    Dim value As String
    
    GetModelProperty value, model, conf, "Обозначение"
    GetPropertySign = value
End Function


