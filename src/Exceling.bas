Attribute VB_Name = "Exceling"
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

Function GetColumnOf(property As String, table As TableAnnotation) As Integer
   
   Dim i As Integer
   Dim bomTable As BomTableAnnotation
   
   GetColumnOf = -1
   Set bomTable = table
   For i = 0 To table.ColumnCount - 1
      If StrComp(bomTable.GetColumnCustomProperty(i), property, vbTextCompare) = 0 Then
         GetColumnOf = i
         Exit Function
      End If
   Next

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

Sub RewriteCell(text As String, cell As Range)

   Dim x() As Byte
   Dim i As Variant
   Dim symbol As String
   
   x = StrConv(text, vbFromUnicode)
   For Each i In x
      If i < 256 Then
         symbol = Chr(i)
      Else
         symbol = ChrW(i) 'для польских символов
      End If
      cell.value = cell.value + symbol
   Next
   
End Sub

Sub RewriteColumn(colExcel As Integer, colBom As Integer, ByRef table As TableAnnotation, _
                  sheet As Excel.Worksheet, header As String)
                  
   Dim row As Integer
   
   RewriteCell header, sheet.Cells(1, colExcel)
   For row = 1 To table.RowCount - 1
      RewriteCell GetCellText(row, colBom, table), sheet.Cells(row + 1, colExcel)
   Next
   
End Sub

Sub WriteColumnOf(property As String, xlsCol As Integer, table As TableAnnotation, sheet As Excel.Worksheet)

   Dim needRemove As Boolean
   Dim col As Integer
   
   needRemove = False
   col = GetColumnOf(property, table)
   If col < 0 Then
      col = AddColumnToBOM(property, table)
      If col < 0 Then
         If MsgBox("Невозможно добавить столбец """ & property & """.", vbOKCancel) = vbCancel Then
            ExitApp
         Else
            Exit Sub
         End If
      End If
      needRemove = True
   End If
   RewriteColumn xlsCol, col, table, sheet, property
   If needRemove Then
      table.DeleteColumn col
   End If

End Sub

Sub ImportBOMtoXLS(ByRef table As TableAnnotation, sheet As Excel.Worksheet)

   Dim col As Integer
   Dim delta As Integer
   
   sheet.Cells.NumberFormat = "@"
   
   WriteColumnOf pDsg, xlsColumnDesignation, table, sheet
   WriteColumnOf pName, xlsColumnNaming, table, sheet
   
   delta = 1
   For col = 0 To table.ColumnCount - 1
      If table.GetColumnType(col) = swBomTableColumnType_Quantity Then
         RewriteColumn xlsColumnNaming + delta, col, table, sheet, GetCellText(0, col, table)
         delta = delta + 1
      End If
   Next
   
   WriteColumnOf "Примечание", xlsColumnNaming + delta, table, sheet
   delta = delta + 1
   
   WriteColumnOf "Заготовка", xlsColumnNaming + delta, table, sheet
   delta = delta + 1
   
   WriteColumnOf "Материал", xlsColumnNaming + delta, table, sheet
   delta = delta + 1
   
   WriteColumnOf "Типоразмер", xlsColumnNaming + delta, table, sheet
   delta = delta + 1
   
   WriteColumnOf "Длина", xlsColumnNaming + delta, table, sheet
   delta = delta + 1
   
   WriteColumnOf "Ширина", xlsColumnNaming + delta, table, sheet
   delta = delta + 1
   
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
    
    Dim Designation As Excel.Range
    For i = 1 To sheet.UsedRange.Columns.Count
        Set Designation = sheet.Cells(1, i)
        Designation.value = Capitalize(Designation.text)
    Next
    For i = 1 To sheet.UsedRange.Rows.Count
        Set Designation = sheet.Cells(i, xlsColumnNaming)
        Dim word As Variant
        For Each word In titles
            If Designation Like word Then
                Designation.Font.Bold = True
                Designation.Font.size = 16
                Designation.value = Capitalize(Designation.text)
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
        
        'Open Excel
        Set xlApp = New Excel.Application
        Set ExcelBOM = xlApp.Workbooks.Add
        ImportBOMtoXLS swTable, ExcelBOM.Worksheets(1)
        
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

Private Function AddColumnToBOM(Prop As String, swTable As TableAnnotation) As Integer

    Dim swBom As BomTableAnnotation
    Set swBom = swTable
    
    AddColumnToBOM = -1
    If swTable.InsertColumn2(swTableItemInsertPosition_Last, 0, Prop, swInsertColumn_DefaultWidth) Then
        swBom.SetColumnCustomProperty swTable.ColumnCount - 1, Prop
        AddColumnToBOM = swTable.ColumnCount - 1
    End If
    
End Function
