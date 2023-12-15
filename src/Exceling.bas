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
   Dim I As Integer
   Dim bomTable As BomTableAnnotation
   
   GetColumnOf = -1
   Set bomTable = table
   For I = 0 To table.ColumnCount - 1
      If StrComp(bomTable.GetColumnCustomProperty(I), property, vbTextCompare) = 0 Then
         GetColumnOf = I
         Exit Function
      End If
   Next
End Function

Function GetCellText(Row As Integer, Col As Integer, ByRef table As TableAnnotation) As String
    Dim text As String
    
    text = table.DisplayedText(Row, Col)
    If text Like "<*>*" Then
        Dim ary() As String
        ary = Split(text, ">")
        text = ary(UBound(ary))
    End If
    text = Replace(text, vbCrLf, " ")
    GetCellText = text
End Function

Sub RewriteCell(text As String, _
                Cell As Object) 'Range
   Dim x() As Byte
   Dim I As Variant
   Dim symbol As String
   
   x = StrConv(text, vbFromUnicode)
   For Each I In x
      If I < 256 Then
         symbol = Chr(I)
      Else
         symbol = ChrW(I) 'для польских символов
      End If
      Cell.value = Cell.value + symbol
   Next
End Sub

'@sheet is Excel.Worksheet
Sub RewriteColumn(colExcel As Integer, colBom As Integer, ByRef table As TableAnnotation, _
                  sheet As Object, header As String)
   Dim Row As Integer
   
   RewriteCell header, sheet.Cells(1, colExcel)
   For Row = 1 To table.RowCount - 1
      RewriteCell GetCellText(Row, colBom, table), sheet.Cells(Row + 1, colExcel)
   Next
End Sub

'@sheet is Excel.Worksheet
Sub WriteColumnOf(property As String, xlsCol As Integer, table As TableAnnotation, sheet As Object)
   Dim needRemove As Boolean
   Dim Col As Integer
   
   needRemove = False
   Col = GetColumnOf(property, table)
   If Col < 0 Then
      Col = AddColumnToBOM(property, table)
      If Col < 0 Then
         If MsgBox("Невозможно добавить столбец """ & property & """.", vbOKCancel) = vbCancel Then
            ExitApp
         Else
            Exit Sub
         End If
      End If
      needRemove = True
   End If
   RewriteColumn xlsCol, Col, table, sheet, property
   If needRemove Then
      table.DeleteColumn Col
   End If
End Sub

Function DefineNameProperty(table As TableAnnotation) As String
    Dim Prop As String
    Dim name As Variant
    Dim I As Integer
    Dim Bom As BomTableAnnotation
    
    DefineNameProperty = pName
    Set Bom = table
    For I = 0 To table.ColumnCount
        For Each name In pAllNames
            Prop = Bom.GetColumnCustomProperty(I)
            If StrComp(Prop, name, vbTextCompare) = 0 Then
                DefineNameProperty = Prop
                GoTo EndFunction
            End If
        Next
    Next
EndFunction:
End Function

'@sheet is Excel.Worksheet
Sub ImportBOMtoXLS(ByRef table As TableAnnotation, sheet As Object)
   Dim Col As Integer
   Dim delta As Integer
   
   sheet.Cells.NumberFormat = "@"
   
   WriteColumnOf pDsg, xlsColumnDesignation, table, sheet
   WriteColumnOf DefineNameProperty(table), xlsColumnNaming, table, sheet
   
   delta = 1
   For Col = 0 To table.ColumnCount - 1
      If table.GetColumnType(Col) = swBomTableColumnType_Quantity Then
         RewriteColumn xlsColumnNaming + delta, Col, table, sheet, GetCellText(0, Col, table)
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

'@sheet is Excel.Worksheet
Sub FormatXLS(sheet As Object)
    Dim I As Integer
    Dim NameCell As Object 'Excel.Range
    
    For I = 6 To sheet.UsedRange.Columns.Count - 1
        If Not IsError(sheet.Cells(1, I)) Then
            If Len(sheet.Cells(1, I)) = 1 Then
                If sheet.Cells(1, I) = " " Then
                    sheet.Cells(1, I).value = "00"
                Else
                    sheet.Cells(1, I).value = "0" + sheet.Cells(1, I).text
                End If
            End If
        End If
    Next
    sheet.name = "List-0" 'constant, agreed with Ivanyna
   
    sheet.Rows(1).Font.Bold = True 'headers
        
    For I = 2 To sheet.UsedRange.Rows.Count
        Set NameCell = sheet.Cells(I, xlsColumnNaming)
        If IsGroupRow(sheet, I) Then
            NameCell.Font.Bold = True
            NameCell.Font.size = 16
        End If
    Next
    
    sheet.Columns(xlsColumnDesignation).AutoFit
    sheet.Columns(xlsColumnNaming).AutoFit
End Sub

'@sheet is Excel.Worksheet
Function IsGroupRow(sheet As Object, Row As Integer) As Boolean
    Dim I As Integer

    IsGroupRow = True
    For I = xlsColumnNaming + 1 To sheet.UsedRange.Columns.Count
        If Not IsCellEmpty(sheet, Row, I) Then
            IsGroupRow = False
            GoTo EndFunction
        End If
    Next
    For I = 1 To xlsColumnNaming - 1
        If Not IsCellEmpty(sheet, Row, I) Then
            IsGroupRow = False
            GoTo EndFunction
        End If
    Next
EndFunction:
End Function

'@sheet is Excel.Worksheet
Function IsCellEmpty(sheet As Object, Row As Integer, Col As Integer)
    Dim Cell As Object 'Excel.Range
    
    Set Cell = sheet.Cells(Row, Col)
    IsCellEmpty = (Trim(Cell) = "")
End Function

Sub SaveBOMtoXLS(ByRef swTable As TableAnnotation, fullFileNameNoExt As String)
    Const warning As String = "Спецификация не будет создана."
    Dim xlsfile As String
    Dim xlApp As Object 'Excel.Application
    Dim ExcelBOM As Object 'Excel.Workbook
    Dim countFilenameChars As Integer

    If Not swTable Is Nothing Then
        xlsfile = fullFileNameNoExt + ".xls"
        countFilenameChars = Len(xlsfile)
        If countFilenameChars > maxPathLength Then
            MsgBox "Слишком длинное имя файла (" & str(countFilenameChars) & " > " & str(maxPathLength) & "):" _
                & vbNewLine & xlsfile & vbNewLine & warning, vbCritical
        End If
        If Not RemoveOldFile(xlsfile) Then
            MsgBox "Не удается удалить старый файл:" & vbNewLine & _
                xlsfile & vbNewLine & warning, vbCritical
            Exit Sub
        End If
        
        'Open Excel
        Set xlApp = CreateObject("Excel.Application")
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
    Dim I As Integer, Length As Integer, char As String
    
    AnyCase = ""
    Length = Len(text)
    If Length > 0 Then
        For I = 1 To Length
            char = Mid(text, I, 1)
            AnyCase = AnyCase + "[" + LCase(char) + UCase(char) + "]"
        Next
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
