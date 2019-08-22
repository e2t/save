Attribute VB_Name = "Exceling"
'Written in 2015-2017 by Eduard E. Tikhenko <aquaried@gmail.com>
'
'To the extent possible under law, the author(s) have dedicated all copyright
'and related and neighboring rights to this software to the public domain
'worldwide. This software is distributed without any warranty.
'You should have received a copy of the CC0 Public Domain Dedication along
'with this software.
'If not, see <http://creativecommons.org/publicdomain/zero/1.0/>

Option Explicit

Private Const xlsColumnDesignation As Integer = 1
Private Const xlsColumnNaming As Integer = 5

Function GetBOM(ByRef doc As ModelDoc2) As TableAnnotation

    Dim swFeat As Feature
    Dim swBomFeat As BomFeature
    
    Set swFeat = doc.FirstFeature
    Set GetBOM = Nothing
    Do While Not swFeat Is Nothing
        If "BomFeat" = swFeat.GetTypeName Then
            Set swBomFeat = swFeat.GetSpecificFeature2
            Set GetBOM = swBomFeat.GetTableAnnotations(0) 'get first BOM in order
            Exit Do
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop
    
End Function

Function GetColumnOf(colnames() As String, ByRef table As TableAnnotation) As Integer
    Dim i As Integer
    
    GetColumnOf = -1
    For i = 0 To table.ColumnCount - 1
        Dim j
        For Each j In colnames
            Dim colname As String
            colname = j
            If table.DisplayedText(0, i) Like ("*" & AnyCase(colname) & "*") Then
                GetColumnOf = i
                Exit Function
            End If
        Next
    Next
    MsgBox "Не найден столбец """ & colnames(0) & """"
End Function

Function GetCellText(row As Integer, col As Integer, ByRef table As TableAnnotation) As String
    Dim text As String
    
    text = table.DisplayedText(row, col)
    If text Like "<*>*" Then
        Dim ary() As String
        ary = Split(text, ">")
        text = ary(UBound(ary))
    End If
    text = Replace(text, vbCrLf, " ")
    GetCellText = text
End Function

Sub RewriteColumn(colExcel As Integer, colBOM As Integer, ByRef table As TableAnnotation, sheet As Excel.Worksheet)
    Dim row As Integer
    Dim text
    
    For row = 0 To table.RowCount - 1
        'sheet.Cells(row + 1, colExcel).value = GetCellText(row, colBOM, table)
        Dim x() As Byte
        x = StrConv(GetCellText(row, colBOM, table), vbFromUnicode)
        Dim i As Variant
        For Each i In x
            If i < 256 Then
                sheet.Cells(row + 1, colExcel).value = sheet.Cells(row + 1, colExcel).value + Chr(i)
            Else
                sheet.Cells(row + 1, colExcel).value = sheet.Cells(row + 1, colExcel).value + ChrW(i)
            End If
        Next
    Next
End Sub

Sub ImportBOMtoXLS(ByRef table As TableAnnotation, sheet As Excel.Worksheet)
    Dim colsign As Integer
    Dim colname As Integer
    Dim i As Integer
    
    sheet.Cells.NumberFormat = "@"
    
    Dim aDesignation(2) As String
    aDesignation(0) = "Обозначение"
    aDesignation(1) = "Designation"
    aDesignation(2) = "Item"
    colsign = GetColumnOf(aDesignation, table)
    If colsign < 0 Then Exit Sub
    
    RewriteColumn xlsColumnDesignation, colsign, table, sheet
    
    ' Configurations are placed after Name always
    Dim aName(1) As String
    aName(0) = "Наименование"
    aName(1) = "Name"
    colname = GetColumnOf(aName, table)
    If colname < 0 Then Exit Sub
    For i = colname To table.ColumnCount - 1
        RewriteColumn xlsColumnNaming - colname + i, i, table, sheet
    Next

    FormatXLS sheet
End Sub

Sub FormatXLS(sheet As Excel.Worksheet)
    Dim i As Integer
    
    For i = 6 To sheet.UsedRange.Columns.Count - 1
        If Not IsError(sheet.Cells(1, i)) Then
            If Len(sheet.Cells(1, i)) = 1 Then
                If sheet.Cells(1, i) = " " Then
                    sheet.Cells(1, i).value = "00"
                Else
                    sheet.Cells(1, i).value = "0" + sheet.Cells(1, i).text
                End If
            End If
        End If
    Next
    sheet.name = "List-0"  'constant, agreed with Ivanyna
    
    Dim titles(13) As String
    titles(0) = "*" + AnyCase("документация")
    titles(1) = "*" + AnyCase("комплек") + "[сСтТ]" + AnyCase("ы")
    titles(2) = "*" + AnyCase("сборочные единицы")
    titles(3) = "*" + AnyCase("детали")
    titles(4) = "*" + AnyCase("стандартные изделия")
    titles(5) = "*" + AnyCase("проч") + "[иИеЕ]" + AnyCase("е") + "*"
    titles(6) = "*" + AnyCase("материалы")
    titles(7) = "*" + AnyCase("покупные") + "*"
    titles(8) = "*" + AnyCase("assembly units")
    titles(9) = "*" + AnyCase("details")
    titles(10) = "*" + AnyCase("standard products")
    titles(11) = "*" + AnyCase("third party products")
    titles(12) = "*" + AnyCase("materials")
    titles(13) = "*" + AnyCase("other")
    
    sheet.Rows(1).Font.Bold = True
    
    Dim designation As Excel.Range
    For i = 1 To sheet.UsedRange.Columns.Count
        Set designation = sheet.Cells(1, i)
        designation.value = Capitalize(designation.text)
    Next
    For i = 1 To sheet.UsedRange.Rows.Count
        Set designation = sheet.Cells(i, xlsColumnNaming)
        Dim word As Variant
        For Each word In titles
            If designation Like word Then
                designation.Font.Bold = True
                designation.Font.size = 16
                designation.value = Capitalize(designation.text)
                Exit For
            End If
        Next
    Next
    
    sheet.Columns(xlsColumnDesignation).AutoFit
    sheet.Columns(xlsColumnNaming).AutoFit
End Sub

Sub SaveBOMtoXLS(ByRef swTable As TableAnnotation, fullFileNameNoExt As String)
    Const warning As String = "Спецификация не будет создана."
    Dim xlsfile As String
    Dim xlApp As Excel.Application
    Dim ExcelBOM As Excel.Workbook
    Dim countFilenameChars As Integer

    If Not swTable Is Nothing Then
        xlsfile = fullFileNameNoExt + ".xls"
        countFilenameChars = Len(xlsfile)
        If countFilenameChars > maxPathLength Then
            MsgBox "Слишком длинное имя файла (" & str(countFilenameChars) & " > " & str(maxPathLength) & "):" & vbNewLine & _
                   xlsfile & vbNewLine & _
                   warning, vbCritical
        End If
        If Not RemoveOldFile(xlsfile) Then
            MsgBox "Не удается удалить старый файл:" & vbNewLine & _
                   xlsfile & vbNewLine & _
                   warning, vbCritical
            Exit Sub
        End If
        
        'add columns
        Dim countNewColumns As Integer
        countNewColumns = AddColumnToBOM(swTable, "Заготовка") + _
                          AddColumnToBOM(swTable, "Материал") + _
                          AddColumnToBOM(swTable, "Типоразмер") + _
                          AddColumnToBOM(swTable, "Длина") + _
                          AddColumnToBOM(swTable, "Ширина")
                          
        'Open Excel
        Set xlApp = New Excel.Application
        Set ExcelBOM = xlApp.Workbooks.Add
        ImportBOMtoXLS swTable, ExcelBOM.Worksheets(1)
        
        'remove columns
        Dim i As Integer
        If countNewColumns > 0 Then
            For i = 1 To countNewColumns
                swTable.DeleteColumn swTable.ColumnCount - 1
            Next
        End If
        
        'Close Excel
        ExcelBOM.SaveAs xlsfile, 39, , , , , , 2
        ExcelBOM.Close
        xlApp.Quit
        Set ExcelBOM = Nothing
        Set xlApp = Nothing
    End If
End Sub

Private Function AnyCase(text As String) As String

    Dim i As Integer, length As Integer, char As String
    AnyCase = ""
    length = Len(text)
    If length > 0 Then
        For i = 1 To length
            char = Mid(text, i, 1)
            AnyCase = AnyCase + "[" + LCase(char) + UCase(char) + "]"
        Next
    End If
    
End Function

Private Function Capitalize(text As String) As String

    Dim length As Integer
    Capitalize = ""
    length = Len(text)
    If length > 0 Then
        Capitalize = UCase(Left(text, 1)) + LCase(Mid(text, 2, Len(text) - 1))
    End If
    
End Function

Private Function AddColumnToBOM(ByRef swTable As TableAnnotation, Prop As String) As Integer

    AddColumnToBOM = 0
    Dim swBom As BomTableAnnotation
    Set swBom = swTable
    If swTable.InsertColumn2(swTableItemInsertPosition_Last, 0, Prop, swInsertColumn_DefaultWidth) Then
        swBom.SetColumnCustomProperty swTable.ColumnCount - 1, Prop
        AddColumnToBOM = 1
    End If
    
End Function
