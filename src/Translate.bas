Attribute VB_Name = "Translate"
'Written in 2016 by Eduard E. Tikhenko <aquaried@gmail.com>
'
'To the extent possible under law, the author(s) have dedicated all copyright
'and related and neighboring rights to this software to the public domain
'worldwide. This software is distributed without any warranty.
'You should have received a copy of the CC0 Public Domain Dedication along
'with this software.
'If not, see <http://creativecommons.org/publicdomain/zero/1.0/>

Option Explicit

Sub ChangeTR(doc As ModelDoc2, ByRef tr As Note, ByRef oldTR As String, preview As Boolean, ByRef abort As Boolean)
    Dim polTr As Note
    Dim translatedText As String
    Dim editedText As String
    
    Set tr = FindExistingTR(doc, TRname(langRussian))
    If tr Is Nothing Then
        If MsgBox("Технические требования не найдены!" & vbNewLine & doc.GetTitle _
                  & vbNewLine & vbNewLine & "Продолжить сохранение без перевода?", vbOKCancel) = vbCancel Then
            abort = True
        End If
        Exit Sub
    End If
    
    Set polTr = FindExistingTR(doc, TRname(langPoland))
    If Not polTr Is Nothing Then
        translatedText = polTr.PropertyLinkedText
    Else
        translatedText = TranslateTR(tr, abort)
    End If
    
    If Not abort Then
        If preview Then
            PreviewForm.txtMemo.text = translatedText
            PreviewForm.Show
            If PreviewForm.abort Then
                abort = True
                Exit Sub
            End If
            editedText = PreviewForm.txtMemo.text
        Else
            editedText = translatedText
        End If
        oldTR = tr.PropertyLinkedText
        If Not polTr Is Nothing Then
            polTr.SetText editedText
        End If
        tr.SetText editedText
    End If
End Sub

Function TranslateTR(ByRef tr As Note, ByRef abort As Boolean) As String
    Dim translatedLines As Variant

    translatedLines = TranslateText(Split(tr.GetText, vbNewLine), "ru-pl", abort)
    If Not abort Then
        TranslateTR = Join(translatedLines, vbNewLine)
    End If
End Function

Sub RestoreOldTR(tr As Note, oldText As String)
    tr.PropertyLinkedText = oldText
End Sub

Function FindExistingTR(drawing As DrawingDoc, notename As String) As Note
    Dim aView As View
    Dim anote As Note
    
    Set aView = drawing.GetFirstView
    Do
        Set anote = aView.GetFirstNote
        While Not anote Is Nothing
            If anote.GetName = notename Then
                Set FindExistingTR = anote
                Exit Function
            End If
            Set anote = anote.GetNext
        Wend
        Set aView = aView.GetNextView
    Loop While Not aView Is Nothing
End Function

Function TranslateText(lines As Variant, languagePair As String, ByRef abort As Boolean) As Variant
    Dim i As Integer
    Dim line As String
    Dim translated As String
    
    For i = LBound(lines) To UBound(lines)
        line = lines(i)
        translated = TranslateLine(line, languagePair, abort)
        If abort Then
            Exit Function
        End If
        If translated <> "" Then
            lines(i) = translated
        End If
    Next
    TranslateText = lines
End Function

''' Powered by Yandex.Translate
''' http://translate.yandex.com/.
Function TranslateLine(inputText As String, languagePair As String, ByRef abort As Boolean) As String
    Const key As String = "trnsl.1.1.20160511T054505Z.59486c62b8d05327.8a43b8707ad4b92f27dde97a2769620b793f4964"
    Dim url As String
    Dim webText As String
    
    url = "https://translate.yandex.net/api/v1.5/tr/translate" & _
          "?key=" & key & "&text=" & inputText & "&lang=" & languagePair
    webText = GetFromWebpage(url, abort)
    If abort Then
        Exit Function
    End If
    TranslateLine = GetTranslatedText(webText, abort)
End Function

Function GetTranslatedText(webText As String, ByRef abort As Boolean) As String
    Dim objXML As MSXML2.DOMDocument
    Dim point As IXMLDOMNode
    
    Set objXML = New MSXML2.DOMDocument
    objXML.loadXML webText
    Set point = objXML.FirstChild
    
    On Error GoTo ErrorTranslate
    GetTranslatedText = point.selectSingleNode("/Translation/text").text
    Exit Function
    
ErrorTranslate:
    MsgBox "Сервер сообщил об ошибке: " & vbNewLine & vbNewLine & webText, vbCritical
    abort = True
End Function

Function GetFromWebpage(url As String, ByRef abort As Boolean) As String
    On Error GoTo Err_GetFromWebpage
    
    Dim objWeb As Object
    Dim strXML As String
     
    Set objWeb = CreateObject("MSXML2.ServerXMLHTTP")
    objWeb.Open "GET", url, False
    objWeb.send
    strXML = objWeb.responsetext
    GetFromWebpage = strXML

End_GetFromWebpage:
    Set objWeb = Nothing
    Exit Function

Err_GetFromWebpage:
    MsgBox err.Description & " (" & err.number & ")"
    abort = True
    Resume End_GetFromWebpage
End Function


