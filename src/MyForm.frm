VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MyForm 
   Caption         =   "Сохранить чертеж как"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6720
   OleObjectBlob   =   "MyForm.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "MyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Written in 2015-2016 by Eduard E. Tikhenko <aquaried@gmail.com>
'
'To the extent possible under law, the author(s) have dedicated all copyright
'and related and neighboring rights to this software to the public domain
'worldwide. This software is distributed without any warranty.
'You should have received a copy of the CC0 Public Domain Dedication along
'with this software.
'If not, see <http://creativecommons.org/publicdomain/zero/1.0/>

Option Explicit

Public Function IsForAll() As Boolean
    IsForAll = allBox.value Or allInFolderBox.value
End Function

Private Sub allBox_Click()
    If allBox.value Then
        allInFolderBox.value = False
        allInFolderBox.Enabled = False
    Else
        allInFolderBox.Enabled = True
    End If
    ChangeCaptions
End Sub

Private Sub allInFolderBox_Click()
    If allInFolderBox.value Then
        allBox.value = False
        allBox.Enabled = False
    Else
        allBox.Enabled = True
    End If
    ChangeCaptions
End Sub

Private Sub attachBox_Click()
    SaveBoolSetting "attach", attachBox.value
End Sub

Private Sub breakBox_Click()
    If breakBox.value Then
        incBox.value = False
        incBox.Enabled = False
    Else
        incBox.Enabled = True
    End If
    ChangeCaptions  'strong after changing the checkbox values
End Sub

Private Sub cancelBut_Click()
    Unload Me
End Sub

Private Sub changeBox_Change()
    SaveIntSetting "changemode", changeBox.ListIndex
    useEngNames = GetEngViews
    ChangeCaptions  'strong after changing the checkbox values
End Sub

Private Sub chkPreview_Click()
    SaveBoolSetting "preview", chkPreview.value
End Sub

Private Sub chkTranslate_Click()
    SaveBoolSetting "translate", chkTranslate.value
End Sub

Private Sub closeBox_Click()
    SaveBoolSetting "close", closeBox.value
End Sub

Private Sub incBox_Click()
    If incBox.value Then
        breakBox.value = False
        breakBox.Enabled = False
    Else
        breakBox.Enabled = True
    End If
    ChangeCaptions  'strong after changing the checkbox values
End Sub

Private Sub openBox_Click()
    SaveBoolSetting "open", openBox.value
End Sub

Private Sub radExportCurrent_Click()
    SaveIntSetting "exportmodel", exportCurrent
End Sub

Private Sub radExportLiked_Click()
    SaveIntSetting "exportmodel", exportLiked
End Sub

Private Sub radExportNone_Click()
    SaveIntSetting "exportmodel", exportNone
End Sub

Private Sub saveBut_Click()
    Dim aForAllMode As ForAllMode
    
    If allBox.value Then
        aForAllMode = forAllOpened
    ElseIf allInFolderBox.value Then
        aForAllMode = forAllInFolder
    Else
        aForAllMode = forActive
    End If
    ConvertDocs aForAllMode
End Sub

Private Sub pdfBox_Click()
    SaveBoolSetting "pdf", pdfBox.value
End Sub

Private Sub dwgBox_Click()
    SaveBoolSetting "dwg", dwgBox.value
End Sub

Private Sub dxfBox_Click()
    SaveBoolSetting "dxf", dxfBox.value
End Sub

Private Sub scaleBox_Click()
    swApp.SetUserPreferenceIntegerValue swDxfOutputNoScale, scaleBox.value
End Sub

Private Sub singlyBox_Click()
    SaveBoolSetting "singly", singlyBox.value
End Sub

Private Sub tifBox_Click()
    SaveBoolSetting "tif", tifBox.value
End Sub

Private Sub psdBox_Click()
    SaveBoolSetting "psd", psdBox.value
End Sub

Private Sub UserForm_Initialize()
    InitSettings
End Sub

Private Sub xlsNeedBox_Change()
    SaveIntSetting "xlsneed", xlsNeedBox.ListIndex
End Sub
