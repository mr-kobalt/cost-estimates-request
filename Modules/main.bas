Attribute VB_Name = "main"
Option Explicit
Option Base 1

Sub createSalesOffer()
    Dim desiredColumns As Collection, purchColumns As Collection
    Dim salesRange As Range, salesHeader As Range, salesFooter As Range
    Dim timeStamp As String

    'Dim StartTime As Double
    'StartTime = Timer
    
    On Error GoTo ErrorHandler
    Application.StatusBar = "Begin making the sales offer..."
    changeUpdatingState False

    ' ���������� � ������ ����, �� ������� ����� ����������� ��
    ' ��������� ����������� ������ (ROW_OFFSET - 1) �����
    On Error GoTo ErrorHandler2
'    If isExistSheet(SALES_SHEET_NAME) Then
'        Sheets(SALES_SHEET_NAME).Activate
'        Application.StatusBar = "������� ����� """ & SALES_SHEET_NAME & """"
'        On Error GoTo ErrorHandler
'        'initializeShapes
'        Application.ActiveSheet.[A1].Select ' deselecting all shapes which were selected before creating Sales Offer
'
'        With Application.ActiveSheet
'            .UsedRange.EntireRow.Delete
'            .UsedRange.ClearFormats
'            .UsedRange.ClearOutline
'            .UsedRange.columns.Hidden = False
'            .UsedRange.Rows.Hidden = False
'            .columns(COLUMN_OFFSET + 1).NumberFormat = "@"
'        End With
'    Else
'        Application.StatusBar = "�������� �����""" & SALES_SHEET_NAME & """"
'        With ThisWorkbook.Sheets.Add(after:=ThisWorkbook.Sheets(SPEC_SHEET_NAME))
'            .name = SALES_SHEET_NAME
'            .Activate
'            .columns(COLUMN_OFFSET + 1).NumberFormat = "@"
'        End With
'    End If
    initializeShapes
    timeStamp = Replace(CStr(Now), ":", ".")
    Application.StatusBar = "�������� �����""" & SALES_SHEET_NAME & " " & timeStamp & """"
    With ThisWorkbook.Sheets.Add(after:=ThisWorkbook.Sheets(SPEC_SHEET_NAME))
        .name = SALES_SHEET_NAME & " " & timeStamp
        .Activate
        .columns(COLUMN_OFFSET + 1).NumberFormat = "@"
    End With
    sheetsListUpdate

    Application.StatusBar = "��������� ��������� ��� ����������� � �� �������"
    Set desiredColumns = getDesiredColumns() ' 0.002

    Application.StatusBar = "�������� ��������� ������� �� ������� � �� � ������ �����"
    Set salesRange = makeSalesTable(desiredColumns, SALES_SHEET_NAME & " " & timeStamp) ' 0.4

    Application.StatusBar = "������� �� �� ��� ������, ������ ������ ������� �� �������� ������"
    Set salesRange = delEmptyRows(salesRange) ' 0.13

    If Not (salesRange Is Nothing) Then
        Application.StatusBar = "��������� �����"
        Set salesHeader = insHeader(salesRange, desiredColumns) ' 0.05

        Application.StatusBar = "������ ������� � ���������, ������������"
        correctIndexColumn salesRange.columns(SalesColumns.INDEX_NUMBER) ' 0.07

        Application.StatusBar = "��������� �� ������� � ���������"
        salesRange.Sort key1:=salesRange(1, SalesColumns.INDEX_NUMBER), Order1:=xlAscending ' 0.004

        Application.StatusBar = "������ ������� ��, ������ ���������, ������, �����������"
        Set salesRange = parseSalesRange(salesRange, desiredColumns) ' 0.3

        Application.StatusBar = "��������� ������ ������"
        Set salesFooter = insFooter(salesRange, desiredColumns) ' 0.015

        Application.Calculation = xlCalculationAutomatic
        adjustingSalesRange salesRange, desiredColumns ' 0.6

        Application.StatusBar = "������ �������, ���������� ���������� � ������ " & CONTROL_GROUP_NAME
        hideSalesColumns salesSheetName:=SALES_SHEET_NAME & " " & timeStamp ' 0.16
    Else
        MsgBox "�������� ����� � ���������� ������� �� ����� ""������ �������"""
    End If

    'salesRange.Select

    'salesRange.EntireRow.AutoFit

    'salesRange.Select
    'formatRangeAsType salesRange, "basic"
    'salesRange.Rows.Ungroup
CleanExit:
    Application.StatusBar = "������� �� �����"
    Set desiredColumns = Nothing
    Set purchColumns = Nothing
    Set salesRange = Nothing
    Set salesHeader = Nothing
    Set salesFooter = Nothing

    changeUpdatingState True
    Application.StatusBar = False
    'MsgBox Timer - StartTime
    Exit Sub

ErrorHandler:
    MsgBox "Error " & err.number & ": " & err.Description
    Resume CleanExit

ErrorHandler2:
    MsgBox "Error " & err.number & ": " & err.Description & vbCrLf & vbCrLf _
            & "�� ������ ���� � ��������, ������� ������� ��� ������� � ������� �������. �������� �� ���������."
    Resume CleanExit
End Sub


Public Function getDesiredColumns() As Collection
    Dim checkboxes As Collection
    Dim shape As shape
    Dim topCheckbox As shape
    Dim minTop As Single
    Dim i As Variant

    Set getDesiredColumns = New Collection
    Set checkboxes = New Collection
    minTop = MAXSINGLE

    For Each shape In ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(CONTROL_GROUP_NAME).GroupItems
        If shape.Type = msoFormControl Then
            If shape.FormControlType = xlCheckBox Then
                checkboxes.Add shape, shape.AlternativeText
            End If
        End If
    Next shape

    If checkboxes.Count > 0 Then
        Do
            minTop = MAXSINGLE
            For Each shape In checkboxes
                If shape.Top < minTop Then
                    Set topCheckbox = shape
                    minTop = topCheckbox.Top
                End If
            Next shape
            checkboxes.Remove topCheckbox.AlternativeText

            getDesiredColumns.Add CLng(topCheckbox.AlternativeText), topCheckbox.AlternativeText
        Loop While checkboxes.Count > 0
    End If


    For i = SalesColumns.[_MIDDLE] + 1 To SalesColumns.[_LAST] - 1
        getDesiredColumns.Add i, CStr(i)
    Next i
End Function

Public Function findColNumber(columns As Collection, query As String) As Variant
' -------------------------------------------------------------------------------- '
' ���� � ��������� �������, ������� ����� ������������, ������ query
' � ������ ���������� ���������� � ������ � ���������
' � ��������� ������ ���������� Nothing
' -------------------------------------------------------------------------------- '
    Dim i As Long

    For i = 1 To columns.Count
        If CStr(columns.Item(i)) = query Then
            findColNumber = i
            Exit Function
        End If
    Next i

    findColNumber = Null
End Function

Private Function findCheckboxName(query As String) As Variant
' -------------------------------------------------------------------------------- '
' ���� ��� �������� �� ��������������� ������ � ������ CONTROL_GROUP_NAME
' �� ����� � ��������
' -------------------------------------------------------------------------------- '
    Dim shape As shape
    findCheckboxName = Null

    For Each shape In ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(CONTROL_GROUP_NAME).GroupItems
        If shape.Type = msoFormControl Then
            If shape.FormControlType = xlCheckBox And shape.AlternativeText = query Then
                findCheckboxName = shape.OLEFormat.Object.Text
                Exit Function
            End If
        End If
    Next shape
End Function

Private Function insHeader(ByVal salesRange As Range, desiredColumns As Collection) As Range
    Dim i As Long
    Dim name As Variant, temp As Variant

    salesRange.Cells(1).EntireRow.Insert
    Set insHeader = salesRange.Resize(1).offset(-1)

    For i = 1 To insHeader.Cells.Count
        name = findCheckboxName(CStr(desiredColumns.Item(i)))
        If Not IsNull(name) Then
            insHeader.Cells(i).Value2 = name
        End If

        Select Case True
            Case i = SalesColumns.Price Or i = SalesColumns.total
                temp = insHeader.Cells(i).Value2
                If ThisWorkbook.Sheets(SPEC_SHEET_NAME).Range(INCLUDE_VAT_CELL_NAME).Value2 <> _
                            ThisWorkbook.Sheets(SERVICE_SHEET_NAME).Range(VAT_ARRAY_NAME).Cells(3).Value2 Then
                    temp = temp & " " & ThisWorkbook.Sheets(SPEC_SHEET_NAME).Range(INCLUDE_VAT_CELL_NAME).Value2
                End If
                insHeader.Cells(i).Value2 = temp & ", " & ThisWorkbook.Sheets(SERVICE_SHEET_NAME).Range(CURRENCIES_HEADER_ARRAY_NAME).Cells( _
                                Application.Match(ThisWorkbook.Sheets(SPEC_SHEET_NAME).Range(SALES_CURRENCY_CELL_NAME).Value2, _
                                ThisWorkbook.Sheets(SERVICE_SHEET_NAME).Range(CURRENCIES_ARRAY_NAME), 0)).Value2

            Case i = SalesColumns.vat:
                insHeader.Cells(i).Value2 = insHeader.Cells(i).Value2 & ", " & ThisWorkbook.Sheets(SERVICE_SHEET_NAME).Range(CURRENCIES_HEADER_ARRAY_NAME).Cells( _
                                Application.Match(ThisWorkbook.Sheets(SPEC_SHEET_NAME).Range(SALES_CURRENCY_CELL_NAME).Value2, _
                                ThisWorkbook.Sheets(SERVICE_SHEET_NAME).Range(CURRENCIES_ARRAY_NAME), 0)).Value2
        End Select
    Next

    Set insHeader = salesRange.Resize(1).offset(-1)
    formatRangeAsType insHeader, "header"
End Function

Private Function insFooter(ByVal salesRange As Range, desiredColumns As Collection) As Range
    Dim i As Long
    Dim columnTotal As Long
    Dim columnVAT As Long
    Dim tempValue As Variant

    Set insFooter = salesRange.Resize(1).offset(salesRange.Rows.Count)

    columnTotal = findColNumber(desiredColumns, SalesColumns.total)
    columnVAT = findColNumber(desiredColumns, SalesColumns.vat)

    insFooter.columns(columnTotal).FormulaR1C1 = "=subtotal(9," & salesRange(1, columnTotal) _
                                                .Address(False, False, xlR1C1, , insFooter.columns(columnTotal).Cells(1)) & _
                                                ":R[-1]C)"
    insFooter.columns(columnVAT).FormulaR1C1 = insFooter.columns(columnTotal).FormulaR1C1

    formatRangeAsType insFooter, "footer"

    With salesRange.Parent.Range(insFooter.Cells(1), insFooter.Cells(columnTotal - 1))
        .ClearContents
        .Merge
        tempValue = TEXTS_TOTAL & " "
        If ThisWorkbook.Sheets(SPEC_SHEET_NAME).Range(INCLUDE_VAT_CELL_NAME).Value2 <> _
                    ThisWorkbook.Sheets(SERVICE_SHEET_NAME).Range(VAT_ARRAY_NAME).Cells(3).Value2 Then
            tempValue = tempValue & ThisWorkbook.Sheets(SPEC_SHEET_NAME).Range(INCLUDE_VAT_CELL_NAME).Value2
        Else
            tempValue = tempValue & "(" & TEXTS_NOT_SUBJECT_VAT & ")"
        End If

        .Value2 = tempValue & ", " & ThisWorkbook.Sheets(SERVICE_SHEET_NAME).Range(CURRENCIES_HEADER_ARRAY_NAME).Cells( _
                                Application.Match(ThisWorkbook.Sheets(SPEC_SHEET_NAME).Range(SALES_CURRENCY_CELL_NAME).Value2, _
                                ThisWorkbook.Sheets(SERVICE_SHEET_NAME).Range(CURRENCIES_ARRAY_NAME), 0)).Value2 & ":"
    End With
End Function

Private Function makeSalesTable(desiredColumns As Collection, salesSheetName As String)
' -------------------------------------------------------------------------------- '
' �������� �� ��������� "������" ����������� ��� �� �������, � �����
' ��������� �����, ����� ��� ����, ��� � ������.
'
' ���� ������� �� desiredColumns �� �������, �� ��������� ������ ����� ������� #N/A
' -------------------------------------------------------------------------------- '
    Dim i As Long, j As Long
    Dim newColumn As Range
    Dim columnValue As Variant
    Dim shape As shape
    Dim tempAddress1 As String, tempFormula As String, total As String

    With ThisWorkbook.Sheets(SPEC_SHEET_NAME).Range(PURCHASE_TABLE_NAME)
        For i = 1 To desiredColumns.Count
            Set newColumn = ThisWorkbook.Sheets(salesSheetName).Range(Cells(ROW_OFFSET + 1, COLUMN_OFFSET + i), _
                                  Cells(ROW_OFFSET + .Rows.Count, COLUMN_OFFSET + i))
            columnValue = desiredColumns.Item(i)

            Select Case True
                Case columnValue = SalesColumns.INDEX_NUMBER
                    newColumn.Value2 = .columns(PurchaseColumns.INDEX_NUMBER).Value2
                    formatRangeAsType newColumn

                Case columnValue = SalesColumns.PN
                    newColumn.Value2 = .columns(PurchaseColumns.PN).Value2
                    formatRangeAsType newColumn, "wo-zeros"

                Case columnValue = SalesColumns.NAME_AND_DESCRIPTION
                    newColumn.Value2 = .columns(PurchaseColumns.NAME_AND_DESCRIPTION).Value2
                    formatRangeAsType newColumn, "wo-zeros"

                Case columnValue = SalesColumns.qty
                    newColumn.Value2 = .columns(PurchaseColumns.qty).Value2
                    formatRangeAsType newColumn, "center"

                Case columnValue = SalesColumns.Unit
                    newColumn.Value2 = .columns(PurchaseColumns.Unit).Value2
                    formatRangeAsType newColumn, "center"

                Case columnValue = SalesColumns.Price
                    newColumn.Value2 = .columns(PurchaseColumns.PRICE_SALES).Value2
                    formatRangeAsType newColumn, "price"

                Case columnValue = SalesColumns.total
                    newColumn.FormulaR1C1 = "=RC[" & (findColNumber(desiredColumns, SalesColumns.qty) - findColNumber(desiredColumns, SalesColumns.total)) & "]" & _
                                            "*RC[" & (findColNumber(desiredColumns, SalesColumns.Price) - findColNumber(desiredColumns, SalesColumns.total)) & "]"
                    formatRangeAsType newColumn, "price"

                Case columnValue = SalesColumns.vat
                    For j = 1 To .Rows.Count
                        tempFormula = "=ROUND(RC[" & CStr(findColNumber(desiredColumns, SalesColumns.total) - findColNumber(desiredColumns, SalesColumns.vat)) & "]"
                        If .columns(PurchaseColumns.VAT_SALES).Cells(j).Value2 = ThisWorkbook.Sheets(SERVICE_SHEET_NAME).Range(VAT_ARRAY_NAME).Cells(1).Value2 Then
                            tempFormula = tempFormula & "*1/1.18*0.18"
                        ElseIf .columns(PurchaseColumns.VAT_SALES).Cells(j).Value2 = ThisWorkbook.Sheets(SERVICE_SHEET_NAME).Range(VAT_ARRAY_NAME).Cells(2).Value2 Then
                            tempFormula = tempFormula & "*0.18"
                        Else
                            tempFormula = tempFormula & "*0"
                        End If
                        newColumn.Cells(j).FormulaR1C1 = tempFormula & "," & PRICE_ROUNDING_UP_TO_QTY & ")"
                    Next j
                    formatRangeAsType newColumn, "price"

                Case columnValue = SalesColumns.DELIVERY_TIME
                    newColumn.Value2 = .columns(PurchaseColumns.DELIVERY_TIME).Value2
                    formatRangeAsType newColumn, "wo-zeros"

                Case columnValue > PurchaseColumns.[_FIRST] And columnValue < PurchaseColumns.[_LAST]
                    newColumn.Value2 = .columns(columnValue).Value2
                    formatRangeAsType newColumn, "wo-zeros"

                Case Else
                    newColumn.Value2 = CVErr(xlErrNA)
            End Select
        Next

        Set makeSalesTable = ThisWorkbook.Sheets(salesSheetName) _
                                        .Range(Cells(ROW_OFFSET + 1, COLUMN_OFFSET + 1), _
                                               Cells(ROW_OFFSET + .Rows.Count, COLUMN_OFFSET + desiredColumns.Count))
    End With
End Function

Private Function delEmptyRows(ByVal toClearRange As Range)
' -------------------------------------------------------------------------------- '
' ������� ��� ������ �� ���������, ������� ����� ������ ������� �� ��������
' ��������, � ���������� ����� ��������.
'
' ���������� Nothing, ���� ���� ������� ��� ������ � ���������
' -------------------------------------------------------------------------------- '
    Dim delRange As Range
    Dim isError As Boolean
    Dim rc As Long

    On Error Resume Next
    Set delRange = toClearRange.columns(1).SpecialCells(xlCellTypeBlanks)
    isError = (err.number <> 0)
    On Error GoTo 0

    If isError Then
        Set delEmptyRows = toClearRange
    Else
        rc = toClearRange.Rows.Count - delRange.Cells.Count
        delRange.EntireRow.Delete

        If rc <= 0 Then
            Set delEmptyRows = Nothing
        Else
            Set delEmptyRows = toClearRange.Resize(rc)
        End If
    End If
End Function

Public Sub correctIndexColumn(ByVal column As Range)
' -------------------------------------------------------------------------------- '
' ������������ ������� � ��������� ����� �� - ������� �� ������ �����, ����������
' � ������������; ������������ ������������������ �������� (������� � �������
' � ����������� 1)
' -------------------------------------------------------------------------------- '
    Dim i As Long, j As Long
    Dim indexArray() As Variant
    Dim temp As Variant
    Dim regexp As Object

    ReDim indexArray(column.Cells.Count, INDEX_RANK_QTY)
    Set regexp = CreateObject("vbscript.regexp")

    Call correctIndexRange(column)
    For i = 1 To column.Cells.Count
        With regexp
            .Global = True
            ' ��������� ��������� ������ ��������, ������ ������ �������� ������ ������ ������ � �� -
            ' ���� ��� ��������� ����� (�� ���������� ���������� INDEX_RANK_QTY), ����������
            ' �� �������� �������. ���� ���������� �������� � ������� ������ INDEX_RANK_QTY,
            ' �� ������������� �������� ����������� ������ ������� - "" (vbNullString)
            .Pattern = "\d+"
            For j = LBound(indexArray, 2) To UBound(indexArray, 2)
                If j <= .Execute(column.Cells(i).Value2).Count Then
                    indexArray(i, j) = .Execute(column.Cells(i).Value2)(j - 1)
                Else
                    indexArray(i, j) = vbNullString
                End If
            Next j
        End With
    Next i

    ' ���������� �������
    Call shrinkColumnsIndices(indexArray)

    ' ���������� ������������ ������� ������� � �������
    For i = 1 To column.Cells.Count
        temp = vbNullString
        For j = 1 To UBound(indexArray, 2)
            If indexArray(i, j) = vbNullString Then: Exit For ' TODO: ��������� ������� � �������������� ������� �� �����
            temp = temp & indexArray(i, j) & "."
        Next j
        column.Cells(i).Value2 = Mid(temp, 1, Len(temp) - 1) ' TODO: ��������� ������� � ���������� ��������� �����
    Next i

    Erase indexArray
End Sub

Public Sub correctIndexRange(ByVal indexRange As Range)
    Dim regexp As Object
    Dim temp As Variant
    Dim indexArray() As Variant
    Dim i As Long

    If isArrayEmpty(indexRange.Value2) Then
        ReDim indexArray(1, 1)
        indexArray(1, 1) = indexRange.Value2
    Else
        indexArray = indexRange.Value2
    End If

    Set regexp = CreateObject("vbscript.regexp")
    For i = LBound(indexArray, 1) To UBound(indexArray, 1)
        temp = CStr(indexArray(i, 1))
        With regexp
            .Global = True
            ' ������� ������� � ���������� ���� � ������ � ����� ������
            .Pattern = "^([^\d]*0*)(?=\d)|[^\d]+$": temp = .Replace(temp, vbNullString)
            ' �������� ������� � ���������� ���� �� �����
            .Pattern = "[^\d]+0*(?=\d)":            temp = .Replace(temp, ".")
            ' ��������� � ������ ������ INDEX_RANK_QTY ������� ����� ���������� �������
            ' ��������, � ������ "13.0.87.1.12" ���� ��������; ��� INDEX_RANK_QTY = 3, �
            ' ����� ��������������, ������ ������ ��� - "13.0.87"
            .Pattern = "(\d+\.){0," & INDEX_RANK_QTY - 1 & "}\d+"
            If .test(temp) Then temp = .Execute(temp)(0) Else temp = vbNullString
        End With

        indexArray(i, 1) = temp
    Next i

    indexRange.Cells.Value2 = indexArray
End Sub

Public Sub updateIndexDesc()
    Dim indexRange As Range
    Dim indexValue As Variant
    Dim indexArray() As Variant
    Dim indexLen As Long
    Dim i As Long, j As Long

    With ThisWorkbook.Sheets(SPEC_SHEET_NAME).Range(PURCHASE_TABLE_NAME)
        Set indexRange = .columns(PurchaseColumns.INDEX_NUMBER)

        If isArrayEmpty(indexRange.Value2) Then
            ReDim indexArray(1, 1)
            indexArray(1, 1) = indexRange.Value2
        Else
            indexArray = indexRange.Value2
        End If

        For i = 1 To .Rows.Count
            indexValue = indexArray(i, 1)
            If Not IsEmpty(indexValue) Then
                If indexValue <> vbNullString And InStr(1, indexValue, ".") = 0 Then
                    indexLen = Len(indexValue) + 1
                    For j = 1 To .Rows.Count
                        If (indexValue & ".") = Left(indexArray(j, 1), indexLen) Then
                            .columns(PurchaseColumns.INDEX_DESC).Cells(i).Value2 = TEXTS_SUBTITLE
                            Exit For
                        End If
                    Next j

                    If j >= .Rows.Count Then
                        .columns(PurchaseColumns.INDEX_DESC).Cells(i).Value2 = vbNullString
                    End If
                ElseIf Right(indexValue, 2) = ".0" Then
                    .columns(PurchaseColumns.INDEX_DESC).Cells(i).Value2 = TEXTS_ASSEMBLY
                ElseIf Not IsEmpty(.columns(PurchaseColumns.INDEX_DESC).Cells(i).Value2) Then
                    .columns(PurchaseColumns.INDEX_DESC).Cells(i).Value2 = vbNullString
                End If
            Else
                .columns(PurchaseColumns.INDEX_DESC).Cells(i).Value2 = vbNullString
            End If
        Next i
    End With
End Sub

Private Function shrinkColumnsIndices(arr() As Variant, Optional column As Long = 1, Optional rowIndices As Collection) As Variant
' -------------------------------------------------------------------------------- '
' arr() - ��������� ������, ������ ������ �������� ������ ������ ������ � �� -
' ���� ��� ��������� ����� (�� ���������� ���������� INDEX_RANK_QTY), ����������
' �� �������� �������. ���� ���������� ����� � ������� ������ INDEX_RANK_QTY,
' �� ������������� �������� ������ ���� ��������� ������ ������� - "" (vbNullString).
' column - ������� �������, � ������� ������������ ������ �������� � ������� ��������
' rowIndices - ��������� ��������� �������� �������, � ������� ����� ������������
' ������ �������� � ������� ��������
' -------------------------------------------------------------------------------- '
' ������� ���������� ���������� ������� � ���������� ��������� �������, ��� ����:
' - ���� ����������� �������� minValue ������� � ������� ������� column �����
' �������� ��������� �������� rowIndices
' - ������� ��� ��������� ������� ��������� ������� ����� �������� rowIndices,
' ������� ����� ���������� ������������ ��������
' - �������� �������� ���� ��������� ��������� �� newIndex
' - ������� ��������� ������� �� ��������� rowIndices
' - ���������� �������� ����, ��������� � �������� ��������� ������, �����
' ��������� ������� � ��������� � ���� 2 �������
' - ����������� �������� ������� newIndex �� 1 � ��������� ����, ���� rowIndices
' �� ����
' -------------------------------------------------------------------------------- '
' ������ ������ �������:
'   10  3   5        3   1   1
'   8   1            1   2
'   8   0            1   1
'   9           -->  2
'   10  10  10       3   2   2
'   10  10  8        3   2   1
' -------------------------------------------------------------------------------- '
    Dim minValue As Variant         ' ������ ����������� �������� � ������� column ����� � ��������� rowIndices
    Dim maxValue As Variant         ' ������ ������������ �������� � ������� column ����� � ��������� rowIndices
    Dim uniqueIntValues As Variant  ' ������ ���-�� ���������� �������� � ������� column ����� � ��������� rowIndices
    Dim newIndex As Variant         ' ����� ������ ��� ������� �������; �������� �������� ���� ��������� ������
                                    ' minValue � ��������� rowIndices

    Dim temp As String
    Dim i As Variant
    Dim slicedIndices As Collection ' ������������ ��� �������� �������� ����� � ���������� ������� minValue

    ' ���� �������� �� ���� �������� � �������, �� ��������� ���� ��������� ��������,
    ' ���������� ��� ������� ������� arr()
    If rowIndices Is Nothing Then
        Set rowIndices = New Collection

        For i = LBound(arr) To UBound(arr)
            ' ������� �������� ��������� ����������� ���� � ���� �������� �������� ���������������� � ������,
            ' ��� ��������� ������� �������� ��������� � �����, �� ����� �������� ��������
            rowIndices.Add i, CStr(i)
        Next i
    End If

    If column <= INDEX_RANK_QTY Then

        newIndex = 1 ' �� ��������� ��������� ������� ������ ������� ������� � �������
        maxValue = maxValueInColumn(arr, column, rowIndices)
        uniqueIntValues = uniqueIntValuesInColumn(arr, column, rowIndices)

        Do While rowIndices.Count > 0 ' ���� ���� �������������� �������� �������
            minValue = minValueInColumn(arr, column, rowIndices)
            ' ���� ��������� ������������ �������� �� �������, �� �������� ��������� ��������
            ' � ��������� ������ �������
            If IsNull(minValue) Then
                Set rowIndices = New Collection
            Else
                Set slicedIndices = New Collection
                ' ���� ����������� �������� ������� ����� ���� � ��� �� ������ ������ �������,
                ' �� �������� ��������� ������ ��������� ������� � ����
                If minValue = 0 And column > 1 Then: newIndex = 0
                ' ������� ��� ��������� ������� ��������� ������� ����� �������� rowIndices,
                ' ������� ����� ���������� ������������ ��������
                For Each i In rowIndices
                    If arr(i, column) = CStr(minValue) Then
                        'temp = String(Len(CStr(maxValue)) - Len(CStr(newIndex)), "0")
                        temp = String(Len(CStr(uniqueIntValues)) - Len(CStr(newIndex)), "0")

                        arr(i, column) = temp & CStr(newIndex) ' �������� �������� ���� ��������� ��������� �� newIndex
                        slicedIndices.Add i, CStr(i) ' ��������� ��������� �������
                        rowIndices.Remove CStr(i) ' ������� ��������� ������� �� ��������� rowIndices
                    End If
                Next i

                ' ���������� �������� ����, ��������� � �������� ��������� ������, �����
                ' ��������� ������� � ��������� �������

                If slicedIndices.Count > 0 Then
                    Call shrinkColumnsIndices(arr, column + 1, slicedIndices)
                End If

                newIndex = newIndex + 1
            End If
        Loop
    End If
End Function

Public Function minValueInColumn(arr() As Variant, column As Long, Optional rowIndices As Collection)
' -------------------------------------------------------------------------------- '
' ������ ������� extremumValueInColumn ��� ������ ������������ �������� � �������
' �������
' -------------------------------------------------------------------------------- '
    minValueInColumn = extremumValueInColumn(arr, column, "<", rowIndices)
End Function

Public Function maxValueInColumn(arr() As Variant, column As Long, Optional rowIndices As Collection)
' -------------------------------------------------------------------------------- '
' ������ ������� extremumValueInColumn ��� ������ ������������� �������� � �������
' �������
' -------------------------------------------------------------------------------- '
    maxValueInColumn = extremumValueInColumn(arr, column, ">", rowIndices)
End Function

Private Function extremumValueInColumn(arr() As Variant, column As Long, op As String, Optional rowIndices As Collection)
' -------------------------------------------------------------------------------- '
' ���� ����������� ��� ������������ �������� (� ����������� �� ���������� op)
' � �������� ������� ������� (������ ���������). ����������� �������� ��������
' ��������� �������� �����, � ������� ����� ������������� �����.
' ���������� Nothing, ���� ������ ���� ��� �������� .
' -------------------------------------------------------------------------------- '
    Dim i As Variant

    If isArrayEmpty(arr) Then
        extremumValueInColumn = Null
    Else
        ' � ������ ����� ����������� ��� ������������ �������� ����������������
        ' ����������� MAXLONG ��� MINLONG
        'extremumValueInColumn = CLng(arr(LBound(arr, 1), column))
        If op = "<" Then
            extremumValueInColumn = MAXLONG
        Else
            extremumValueInColumn = MINLONG
        End If

        ' �������� ���� �� ���� ������� �������, ���� �� ������� � ���������,
        ' ����������� �� ��������� rowIndices
        If rowIndices Is Nothing Then
            For i = LBound(arr) To UBound(arr)
                ' ���� ����� �������� ������� ��� �����, �� ����� ��������� ���������
                ' �����.
                ' https://msdn.microsoft.com/ru-ru/library/215yacb6.aspx
                If IsNumeric(extremumValueInColumn) And IsNumeric(arr(i, column)) Then
                    If Application.Evaluate(CStr(arr(i, column)) & op & CStr(extremumValueInColumn)) Then
                        extremumValueInColumn = CLng(arr(i, column))
                    End If
                End If
            Next i
        Else
            If rowIndices.Count <> 0 Then
                For Each i In rowIndices
                    ' ���� ����� �������� ������� ��� �����, �� ����� ��������� ���������
                    ' �����.
                    ' https://msdn.microsoft.com/ru-ru/library/215yacb6.aspx
                    If IsNumeric(extremumValueInColumn) And IsNumeric(arr(i, column)) Then
                        If Application.Evaluate(CStr(arr(i, column)) & op & CStr(extremumValueInColumn)) Then
                            extremumValueInColumn = CLng(arr(i, column))
                        End If
                    End If
                Next i
            End If
        End If

        If extremumValueInColumn = MINLONG Or extremumValueInColumn = MAXLONG Then
            extremumValueInColumn = Null
        End If
    End If
End Function

Private Function uniqueIntValuesInColumn(arr() As Variant, column As Long, Optional rowIndices As Collection)
' -------------------------------------------------------------------------------- '
' ��������� ���-�� ���������� �������� � ������� column ������� arr
' -------------------------------------------------------------------------------- '
    Dim i As Variant
    Dim dict As New Collection ' �� �������, �� ���-�� �������

    If isArrayEmpty(arr) Then
        uniqueIntValuesInColumn = Null
    Else
        ' �������� ���� �� ���� ������� �������, ���� �� ������� � ���������,
        ' ����������� �� ��������� rowIndices
        If rowIndices Is Nothing Then
            For i = LBound(arr) To UBound(arr)
                On Error Resume Next
                If IsNumeric(arr(i, column)) Then
                    dict.Add arr(i, column), arr(i, column)
                End If
                On Error GoTo 0
            Next i
        Else
            If rowIndices.Count <> 0 Then
                For Each i In rowIndices
                    On Error Resume Next
                    If IsNumeric(arr(i, column)) Then
                        dict.Add arr(i, column), arr(i, column)
                    End If
                    On Error GoTo 0
                Next i
            End If
        End If

        uniqueIntValuesInColumn = dict.Count
    End If
End Function

Private Function parseSalesRange(ByVal salesRange As Range, desiredColumns As Collection) As Range
    Dim indexColumn() As Variant
    If isArrayEmpty(salesRange.columns(SalesColumns.INDEX_NUMBER).Value2) Then
        ReDim indexColumn(1, 1)
        indexColumn(1, 1) = salesRange.columns(SalesColumns.INDEX_NUMBER).Value2
    Else
        indexColumn = salesRange.columns(SalesColumns.INDEX_NUMBER).Value2
    End If

    Dim groupCount As Collection: Set groupCount = getIndexGroupCount(indexColumn)
    Dim regexp As Object: Set regexp = CreateObject("vbscript.regexp")
    Dim tempRegExp As Object
    Dim tempValue As Variant
    Dim currentGroup As String
    Dim kitRange As Range
    Dim i As Long, j As Long


    With regexp
        .Global = True
        .Pattern = "\d+"

        i = 1
        Do While i <= salesRange.Rows.Count
            Set tempRegExp = .Execute(salesRange(i, 1).Value2)

            '
            If tempRegExp.Count = 1 And i <= salesRange.Rows.Count Then
                If currentGroup <> tempRegExp(0) Then
                    currentGroup = tempRegExp(0)

                    If currentGroup <> salesRange(i + 1, 1).Value2 Then
                        For j = i + 1 To salesRange.Rows.Count
                            Set tempRegExp = .Execute(salesRange(j, 1).Value2)
                            If tempRegExp.Count > 0 Then
                                If currentGroup <> tempRegExp(0) Then: Exit For
                            End If
                        Next j

                        If (j - 1) - i > 0 Then: Set salesRange = makeSubGroup(salesRange, i, j, desiredColumns)
                    End If
                End If
            ElseIf tempRegExp.Count > 1 Then
                ' ����������� ��������� �������� ������ � �����. ���� ��� ����� ����, ��
                ' ������ ������. ������:
                ' "123.12.000" -> "000" -> 0 = 0
                ' "123.12.010" -> "010" -> 10 != 0
                If CLng(tempRegExp(tempRegExp.Count - 1)) = 0 Then
                    ' �������� ���� � ����� ������. ������:
                    ' "123.12.000" -> "123.12."
                    .Pattern = "(\d+\.){1,}"
                    tempValue = .Execute(salesRange(i, 1).Value2)(0)
                    ' �������� ��������� �����. ������:
                    ' "123.12." -> "123.12"
                    salesRange(i, 1).Value2 = Mid(tempValue, 1, Len(tempValue) - 1)

                    For j = i + 1 To salesRange.Rows.Count
                        Set tempRegExp = .Execute(salesRange(j, 1).Value2)
                        If tempRegExp.Count > 0 Then
                            If InStr(tempRegExp(0), tempValue) <> 1 Then: Exit For
                        Else
                            Exit For
                        End If
                    Next j

                    Set kitRange = salesRange.Parent.Range(salesRange(i, 1), salesRange(j - 1, salesRange.columns.Count))
                    makeKit kitRange, desiredColumns

                    ' ���������� ��� ������, ���������� � ������
                    i = j - 1

                    .Pattern = "\d+"
                End If
            End If

            i = i + 1
        Loop
    End With

    Set parseSalesRange = salesRange
End Function

Private Function makeSubGroup(ByRef salesRange As Range, firstRow As Long, lastRow As Long, desiredColumns As Collection) As Range
    Dim regexp As Object: Set regexp = CreateObject("vbscript.regexp")
    Dim tempRegExp As Object
    Dim tempValue As Variant
    Dim column As Long

    With regexp
        .Global = True
        .Pattern = "\d+"
        Set tempRegExp = .Execute(salesRange(firstRow, SalesColumns.INDEX_NUMBER).Value2)

        tempValue = salesRange(firstRow, findColNumber(desiredColumns, SalesColumns.NAME_AND_DESCRIPTION)).Value2

        With salesRange.Parent.Range(salesRange(firstRow, SalesColumns.INDEX_NUMBER + 1), _
                                                            salesRange(firstRow, salesRange.columns.Count))
            .ClearContents
            .Merge
            .Value2 = tempValue
        End With

        salesRange.Rows(lastRow).EntireRow.Insert xlShiftDown

        If lastRow > salesRange.Rows.Count Then
            Set salesRange = salesRange.Resize(salesRange.Rows.Count + 1, salesRange.columns.Count)
        End If

        column = findColNumber(desiredColumns, SalesColumns.total)
        formatRangeAsType Range(salesRange(firstRow, 1), salesRange(lastRow, salesRange.columns.Count)), "subgroup"

        With salesRange.Parent.Range(salesRange(lastRow, SalesColumns.INDEX_NUMBER), salesRange(lastRow, column - 1))
            .ClearContents
            .Merge
            tempValue = TEXTS_SUBTOTAL
            If ThisWorkbook.Sheets(SPEC_SHEET_NAME).Range(INCLUDE_VAT_CELL_NAME).Value2 <> _
                        ThisWorkbook.Sheets(SERVICE_SHEET_NAME).Range(VAT_ARRAY_NAME).Cells(3).Value2 Then
                tempValue = tempValue & " " & ThisWorkbook.Sheets(SPEC_SHEET_NAME).Range(INCLUDE_VAT_CELL_NAME).Value2
            End If
            .Value2 = tempValue & ", " & ThisWorkbook.Sheets(SERVICE_SHEET_NAME).Range(CURRENCIES_HEADER_ARRAY_NAME).Cells( _
                                    Application.Match(ThisWorkbook.Sheets(SPEC_SHEET_NAME).Range(SALES_CURRENCY_CELL_NAME).Value2, _
                                    ThisWorkbook.Sheets(SERVICE_SHEET_NAME).Range(CURRENCIES_ARRAY_NAME), 0)).Value2 & ":"
        End With
        With salesRange(lastRow, column)
            .FormulaR1C1 = "=SUBTOTAL(9," & _
                            salesRange(firstRow, column).Address(False, False, xlR1C1, , .Cells(1)) & _
                           ":R[-1]C)"
            salesRange(lastRow, findColNumber(desiredColumns, SalesColumns.vat)).FormulaR1C1 = .FormulaR1C1
        End With
    End With

    Set makeSubGroup = salesRange
End Function

Private Sub makeKit(ByVal kitRange As Range, desiredColumns As Collection)
    Dim i As Long, j As Long
    Dim column As Range
    Dim columnQty() As Variant
    Dim columnName As Variant, columnNumberQty As Variant, columnNumberUnit As Variant, columnNumberTotal As Variant
    Dim tempFormula As String
    Dim total As String
    Dim qtyGCD As Long

    On Error GoTo ErrorHandler
    If kitRange.Rows.Count > 1 Then
        columnNumberQty = findColNumber(desiredColumns, SalesColumns.qty)
        columnNumberUnit = findColNumber(desiredColumns, SalesColumns.Unit)
        columnNumberTotal = findColNumber(desiredColumns, SalesColumns.total)
        columnQty = kitRange.columns(columnNumberQty).Value2
        If columnQty(1, 1) = 0 Then
            qtyGCD = Application.WorksheetFunction.Gcd(kitRange.columns(columnNumberQty))
        Else
            qtyGCD = columnQty(1, 1)
        End If

        For i = 1 To kitRange.columns.Count
            With kitRange.columns(i)
                columnName = desiredColumns.Item(i)
                Select Case True
                    Case columnName = SalesColumns.NAME_AND_DESCRIPTION
                        Set column = kitRange.columns(i).Resize(kitRange.Rows.Count - 1).offset(1)

                        For j = 1 To column.Rows.Count
                            column.Cells(j).Value2 = column.Cells(j).Value2 & " - " & _
                                                        CStr(columnQty(j + 1, 1) / qtyGCD) & " " & _
                                                        kitRange(j + 1, columnNumberUnit).Value2
                        Next j

                    Case columnName = SalesColumns.qty
                        .Cells(1).Value2 = qtyGCD
                        .Resize(kitRange.Rows.Count - 1).offset(1).ClearContents
                        .Merge

                    Case columnName = SalesColumns.Unit
                        .ClearContents
                        .Merge
                        .Cells(1).Value2 = ThisWorkbook.Sheets(SERVICE_SHEET_NAME).Range(UNITS_ARRAY_NAME).Cells(2).Value2

                    Case columnName = SalesColumns.Price
                        For j = 1 To kitRange.Rows.Count
                            .Cells(j).Value2 = Round(kitRange.columns(columnNumberTotal).Cells(j).Value2 / qtyGCD, PRICE_ROUNDING_UP_TO_QTY)
                        Next j
                        .Cells(1).Value2 = Application.WorksheetFunction.Sum(kitRange.columns(i))
                        .Resize(kitRange.Rows.Count - 1).offset(1).ClearContents
                        .Merge

                    Case columnName = SalesColumns.total
                        .Resize(kitRange.Rows.Count - 1).offset(1).ClearContents
                        .Merge

                    Case columnName = SalesColumns.vat
                        .Resize(kitRange.Rows.Count - 1).offset(1).ClearContents
                        .Merge
                End Select
            End With
        Next i

        tempFormula = kitRange(1, findColNumber(desiredColumns, SalesColumns.NAME_AND_DESCRIPTION)).FormulaR1C1
        With kitRange.Parent.Range(kitRange(1, SalesColumns.INDEX_NUMBER + 1), kitRange(1, columnNumberQty - 1))
                .ClearContents
                .Merge
                .FormulaR1C1 = tempFormula
        End With

        formatRangeAsType kitRange, "kit"
    End If

CleanExit:
    Set column = Nothing
    Set kitRange = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error " & err.number & ": " & err.Description
    Resume CleanExit
End Sub

Public Function isArrayEmpty(anArray As Variant) As Boolean
' -------------------------------------------------------------------------------- '
' ��������� ������� ������� �������� ������ ������� ������� ��������� � ������
' ��������� ������
'
' http://stackoverflow.com/questions/206324/how-to-check-for-empty-array-in-vba-macro
' -------------------------------------------------------------------------------- '
    Dim i As Integer

    On Error Resume Next
    i = UBound(anArray, 1) ' Just try it. If it fails, Err.Number will be nonzero.
    isArrayEmpty = (err.number <> 0)
    err.Clear
End Function

Private Function isExist(col As Collection, key As Variant) As Boolean
' -------------------------------------------------------------------------------- '
' ��������� ������� �������� � ������ key � ���������
' -------------------------------------------------------------------------------- '
    On Error Resume Next
    col (key) ' Just try it. If it fails, Err.Number will be nonzero.
    isExist = (err.number = 0)
    err.Clear
End Function

Private Function isExistNamedRange(rangeName As String) As Boolean
' -------------------------------------------------------------------------------- '
' ��������� ������� ������������ ��������� � ������ rangeName
' -------------------------------------------------------------------------------- '
    On Error Resume Next
    Application.Names.Item (rangeName)  ' Just try it. If it fails, Err.Number will be nonzero.
    isExistNamedRange = (err.number = 0)
    err.Clear
End Function

Public Function isExistShape(shapeName As String, sheetName As String) As Boolean
' -------------------------------------------------------------------------------- '
' ��������� ������������� ����� shapeName �� ����� sheetName
' -------------------------------------------------------------------------------- '
    Dim shape As shape
    On Error Resume Next
    Set shape = ThisWorkbook.Sheets(sheetName).shapes(shapeName)
    isExistShape = (err.number = 0)

    err.Clear
    Set shape = Nothing
End Function

Public Function isExistSheet(sheetName As String) As Boolean
' -------------------------------------------------------------------------------- '
' ��������� ������� ������������ ��������� � ������ rangeName
' -------------------------------------------------------------------------------- '
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = ThisWorkbook.Sheets(sheetName)
    isExistSheet = (err.number = 0)
    err.Clear
    Set sheet = Nothing
End Function

Public Function customGCD(arr() As Variant) As Variant ' �� ����������� :(
    Dim i As Variant
    Dim j As Variant

    For i = customMin(arr) To 2 Step -1
        For j = LBound(arr, 1) To UBound(arr, 1)
            If arr(j, 1) Mod i <> 0 Then: Exit For
        Next j

        If j > UBound(arr) Then: Exit For
    Next i

    customGGC = i
End Function

Public Function customMin(arr() As Variant) As Variant ' �� ����������� :(
    Dim i As Variant
    customMin = MAXLONG

    For Each i In arr
        If IsNumeric(i) Then
            If customMin > CLng(i) And Not i = Empty Then: customMin = CLng(i)
        Else
            Set customMin = Nothing
            Exit For
        End If
    Next i
End Function

Private Function getIndexGroupCount(indexColumn() As Variant) As Collection ' �� ����������� :(
    Dim i As Long, _
        j As Long
    Dim minRankCount As Long: minRankCount = INDEX_RANK_QTY
    Dim maxRankCount As Long: maxRankCount = 0
    Dim temp As Object
    Dim regexp As Object: Set regexp = CreateObject("vbscript.regexp")

    Set getIndexGroupCount = New Collection

    With regexp
        .Global = True
        .Pattern = "\d+"
        For i = LBound(indexColumn, 1) To UBound(indexColumn, 1)
            If .test(indexColumn(i, 1)) Then
                Set temp = .Execute(indexColumn(i, 1))
                If temp.Count < minRankCount Then: minRankCount = temp.Count
                If temp.Count > maxRankCount Then: maxRankCount = temp.Count
            End If
        Next i
    End With

    getIndexGroupCount.Add minRankCount, key:="MIN"
    getIndexGroupCount.Add maxRankCount, key:="MAX"

End Function

Private Function vStack2D(arr1() As Variant, arr2() As Variant) ' �� ����������� :(
' -------------------------------------------------------------------------------- '
' ������ numpy.vstack ��� ���� ��������� �������� VBA.
' �������� ��� ������� ����������� (���������). ���������� ������� � ��������
' ������ ���������. ���� ���������� ������� �� ���������, � ����� ���� ����� ���
' ���������� �������, �� ���������� ������ ������ Variant()
' -------------------------------------------------------------------------------- '
    Dim temp() As Variant
    Dim i As Long, _
        j As Long, _
        k As Long
    Dim cols1 As Long, cols2 As Long, _
        rows1 As Long, rows2 As Long

    Select Case True
        Case isArrayEmpty(arr1)
            temp = arr2
        Case isArrayEmpty(arr2)
            temp = arr1
        Case Else
            cols1 = UBound(arr1, 2) - LBound(arr1, 2) + 1
            cols2 = UBound(arr2, 2) - LBound(arr2, 2) + 1
            rows1 = UBound(arr1, 1) - LBound(arr1, 1) + 1
            rows2 = UBound(arr2, 1) - LBound(arr2, 1) + 1

            If cols1 = cols2 Then
                ReDim temp(rows1 + rows2, cols1)
                k = 1

                For i = LBound(arr1, 1) To UBound(arr1, 1)
                    For j = LBound(arr1, 2) To UBound(arr1, 2)
                        temp(k, j) = arr1(i, j)
                    Next j

                    k = k + 1
                Next i

                For i = LBound(arr2, 1) To UBound(arr2, 1)
                    For j = LBound(arr2, 2) To UBound(arr2, 2)
                        temp(k, j) = arr2(i, j)
                    Next j

                    k = k + 1
                Next i
            End If
    End Select

    vStack2D = temp
End Function
