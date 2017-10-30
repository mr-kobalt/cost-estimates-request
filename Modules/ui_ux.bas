Attribute VB_Name = "ui_ux"
Option Explicit

Private Sub buttonPN_Click()
' заменяем кириллические символы в слобце P/N на латинские эквиваленты
    Dim i As Long

    changeUpdatingState False

    On Error GoTo ErrorHandler
    i = replaceCyrillicWithLatin(ThisWorkbook.Sheets(SPEC_SHEET_NAME).Range(PURCHASE_TABLE_NAME).columns(PurchaseColumns.PN))

    Select Case True
        Case i = 0
            MsgBox "Кириллических символов не найдено"
        Case i Mod 10 Like "[056789]" Or i Mod 100 Like "*1#"
            MsgBox "Произведено " & i & " замен"
        Case i Mod 10 Like "[234]"
            MsgBox "Произведено " & i & " замены"
        Case i Mod 10 = 1
            MsgBox "Произведена " & i & " замена"
        Case Else ' Если i целое число, то условия выше предусматривают всё множество возможных значений
            MsgBox "Something went wrong here..."
    End Select

CleanExit:
    changeUpdatingState True
    Exit Sub

ErrorHandler:
    MsgBox "Error " & err.number & ": " & err.Desctription
    Resume CleanExit
End Sub

Private Sub buttonDescription_Click()
' очищаем столбец с наименованием от непечатных символов и лишних пробельных символов
    changeUpdatingState False

    On Error GoTo ErrorHandler
    trimAndClearRange ThisWorkbook.Sheets(SPEC_SHEET_NAME).Range(PURCHASE_TABLE_NAME).columns(PurchaseColumns.NAME_AND_DESCRIPTION)

CleanExit:
    changeUpdatingState True
    Exit Sub

ErrorHandler:
    MsgBox "Error " & err.number & ": " & err.Description
    Resume CleanExit
End Sub

Private Sub addFiveRowsLast_Click()
' добавляем пять строк в конец таблицы PURCHASE_TABLE_NAME
    addRows 5, "last"
    ActiveWindow.ScrollRow = ActiveWindow.ScrollRow + 5
End Sub

Private Sub addTwentyRowsLast_Click()
' добавляем двадцать строк в конец таблицы PURCHASE_TABLE_NAME
    addRows 20, "last"
    ActiveWindow.ScrollRow = ActiveWindow.ScrollRow + 20
End Sub
Private Sub addFiveRowsFirst_Click()
' добавляем двадцать строк в конец таблицы PURCHASE_TABLE_NAME
    addRows 5, "first"
End Sub
Private Sub addTwentyRowsFirst_Click()
' добавляем двадцать строк в конец таблицы PURCHASE_TABLE_NAME
    addRows 20, "first"
End Sub
Public Sub addRows(Optional rowCount As Long = 1, Optional pos As String = "last")
    Dim leftBorderRange As Range
    Dim rightBorderRange As Range
    Dim offset As Long

    changeUpdatingState False
    On Error GoTo ErrorHandler

    With ThisWorkbook.Sheets(SPEC_SHEET_NAME)
        offset = .ListObjects(PURCHASE_TABLE_NAME).Range.Rows.Count + rowCount
        .ListObjects(DELIVERY_TABLE_NAME).Range.Cut Destination:=.Cells(.ListObjects(DELIVERY_TABLE_NAME).Range.Row + offset, 1)

        If pos = "last" Then
            .ListObjects(PURCHASE_TABLE_NAME).TotalsRowRange.EntireRow.Resize(rowCount).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Else
            .ListObjects(PURCHASE_TABLE_NAME).DataBodyRange.Rows(1).EntireRow.Resize(rowCount).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
        End If

        .ListObjects(DELIVERY_TABLE_NAME).Range.Cut Destination:=.Cells(.ListObjects(DELIVERY_TABLE_NAME).Range.Row - offset, 1)

        With .Range(PURCHASE_TABLE_NAME)
            Set leftBorderRange = Application.Union(.columns(PurchaseColumns.PRICE_SALES), _
                                                    .columns(PurchaseColumns.MARGIN_AMOUNT), _
                                                    .columns(PurchaseColumns.PRICE_GPL), _
                                                    .columns(PurchaseColumns.PRICE_PURCHASE))
            Set rightBorderRange = Application.Union(.columns(PurchaseColumns.Unit), _
                                                     .columns(PurchaseColumns.VAT_AMOUNT), _
                                                     .columns(PurchaseColumns.MARGIN_AMOUNT), _
                                                     .columns(PurchaseColumns.TOTAL_GPL), _
                                                     .columns(PurchaseColumns.TOTAL_PURCHASE_RECALCULATED))
        End With
    End With

    With leftBorderRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rightBorderRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With

CleanExit:
    Set leftBorderRange = Nothing
    Set rightBorderRange = Nothing
    changeUpdatingState True
    Exit Sub

ErrorHandler:
    MsgBox "Error " & err.number & ": " & err.Description
    Resume CleanExit
End Sub

Private Sub SalesColumns_Click()
' скрываем столбцы КП при переключении соответствующих чекбоксов
    Dim shape As shape
    Dim column As Long
    Dim cell As Range

    Application.EnableEvents = False

    On Error GoTo ErrorHandler
    Set shape = ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(Application.Caller)

    sheetsListUpdate
    column = findColNumber(getDesiredColumns(), shape.AlternativeText) + COLUMN_OFFSET

    With ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(DROPDOWN_SHAPE_NAME).OLEFormat.Object
        For Each cell In Application.Range(.ListFillRange)
            If isExistSheet(cell.Value2) Then
                With ThisWorkbook.Sheets(cell.Value2).Cells.columns(column)
                    If shape.OLEFormat.Object.Value = xlOff Then
                        .Hidden = True
                    ElseIf shape.OLEFormat.Object.Value = xlOn Then
                        .Hidden = False
                    End If
                End With
            End If
        Next cell
    End With

CleanExit:
    Application.EnableEvents = True
    Set shape = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error " & err.number & ": " & err.Description
    Resume CleanExit
End Sub

Private Sub CalcButton_Click()
    Application.EnableEvents = False

    On Error GoTo ErrorHandler
    createSalesOffer

CleanExit:
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "Error " & err.number & ": " & err.Description
    Resume CleanExit
End Sub

Private Sub Dropdown_Change()
    Dim ole As OLEFormat

    On Error GoTo ErrorHandler
    sheetsListUpdate

CleanExit:
    Exit Sub

ErrorHandler:
    MsgBox "Error " & err.number & ": " & err.Description
    Resume CleanExit
End Sub

Private Sub excelExportButton_Click()
    Dim wb As Workbook
    Dim sheetName As String

    On Error GoTo ErrorHandler
    With ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(DROPDOWN_SHAPE_NAME).OLEFormat.Object
        sheetName = Application.Range(.ListFillRange).Cells(.Value).Value2
    End With
    If isExistSheet(sheetName) Then
        Set wb = Application.Workbooks.Add
        ThisWorkbook.Sheets(sheetName).Cells(ROW_OFFSET + 1, COLUMN_OFFSET + 1).CurrentRegion.Copy

        wb.ActiveSheet.Cells(1).PasteSpecial Paste:=xlPasteFormulas, Operation:=xlPasteSpecialOperationNone
        wb.ActiveSheet.Cells(1).PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlPasteSpecialOperationNone
        wb.ActiveSheet.Cells(1).PasteSpecial Paste:=xlPasteFormats, Operation:=xlPasteSpecialOperationNone
    Else
        MsgBox "Лист '" & sheetName & "' не найден"
    End If

    sheetsListUpdate

CleanExit:
    Set wb = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error " & err.number & ": " & err.Description
    Resume CleanExit
End Sub

Private Sub wordExportButton_Click()
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim wdRng As Object
    Dim wdTbl As Object
    Dim wdFld As Object

    Dim sheetName As String
    Dim i As Long
    Dim temp As String, currText As String
    Dim arr() As Variant
    Dim revenue As Range, VATamount As Range
    Dim cell As Range, salesRange As Range
    Dim calcCurrency As String
    Dim VATtype As Long

    On Error GoTo ErrorHandler
    With ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(DROPDOWN_SHAPE_NAME).OLEFormat.Object
        sheetName = ThisWorkbook.Sheets(SERVICE_SHEET_NAME).Range(Mid(.ListFillRange, Len(SERVICE_SHEET_NAME) + 2)).Cells(.Value).Value2
    End With
    If isExistSheet(sheetName) Then
        ' основные параметры КП
        Set salesRange = ThisWorkbook.Sheets(sheetName).UsedRange
        Set revenue = salesRange.Cells(salesRange.Rows.Count, findColNumber(getDesiredColumns, SalesColumns.total))
        Set VATamount = salesRange.Cells(salesRange.Rows.Count, findColNumber(getDesiredColumns, SalesColumns.vat))

        Select Case True
            Case InStr(salesRange.Resize(1, 1).offset(salesRange.Rows.Count - 1).Value2, Application.Range(CURRENCIES_HEADER_ARRAY_NAME).Cells(1).Value2) > 0:
                currText = Application.Range(CURRENCIES_HEADER_ARRAY_NAME).Cells(1).Value2
                calcCurrency = Application.Range(CURRENCIES_ARRAY_NAME).Cells(1).Value2
            Case InStr(salesRange.Resize(1, 1).offset(salesRange.Rows.Count - 1).Value2, Application.Range(CURRENCIES_HEADER_ARRAY_NAME).Cells(2).Value2) > 0:
                currText = Application.Range(CURRENCIES_HEADER_ARRAY_NAME).Cells(2).Value2
                calcCurrency = Application.Range(CURRENCIES_ARRAY_NAME).Cells(2).Value2
            Case InStr(salesRange.Resize(1, 1).offset(salesRange.Rows.Count - 1).Value2, Application.Range(CURRENCIES_HEADER_ARRAY_NAME).Cells(3).Value2) > 0:
                currText = Application.Range(CURRENCIES_HEADER_ARRAY_NAME).Cells(3).Value2
                calcCurrency = Application.Range(CURRENCIES_ARRAY_NAME).Cells(3).Value2
        End Select

        Select Case True
            Case InStr(salesRange.Resize(1, 1).offset(salesRange.Rows.Count - 1).Value2, Application.Range(VAT_ARRAY_NAME).Cells(1).Value2) > 0:
                VATtype = 1
            Case InStr(salesRange.Resize(1, 1).offset(salesRange.Rows.Count - 1).Value2, Application.Range(VAT_ARRAY_NAME).Cells(2).Value2) > 0:
                VATtype = 2
            Case InStr(salesRange.Resize(1, 1).offset(salesRange.Rows.Count - 1).Value2, TEXTS_NOT_SUBJECT_VAT) > 0:
                VATtype = 3
        End Select

        Set wdApp = CreateObject("Word.Application")
        Set wdDoc = wdApp.Documents.Add

        ' активируем окно Word и пытаемся вывести его на передний план
        wdApp.Visible = True
        wdApp.Activate
        wdApp.WindowState = wdWindowStateMaximize

        ' параметры полей документа
        With wdDoc.PageSetup
            .LeftMargin = Application.CentimetersToPoints(2)
            .RightMargin = Application.CentimetersToPoints(1)
            .TopMargin = Application.CentimetersToPoints(2)
            .BottomMargin = Application.CentimetersToPoints(1.75)
            .HeaderDistance = Application.CentimetersToPoints(1)
        End With

        ' умолчания для текста в основном теле документа
        With wdDoc.Content
            .Font.name = DEFAULT_FONT
            .Font.Size = 11
            .ParagraphFormat.SpaceAfter = 0
        End With

        ' заполняем шапку
        Set wdRng = wdDoc.sections(1).headers(1).Range
        Set wdTbl = wdDoc.Tables.Add(Range:=wdRng, NumRows:=1, NumColumns:=3)
        With wdTbl
            .RightPadding = 0
            .LeftPadding = 0
            .Range.Cells.VerticalAlignment = wdAlignVerticalCenter
            With .Rows(1)
                .Height = Application.CentimetersToPoints(1.5)
                ' разделительная линия
                With .Borders(wdBorderBottom)
                    .LineStyle = wdLineStyleSingle
                    .LineWidth = 8
                    .Color = COMPANY_COLOR
                End With
            End With

            With .Range.Font
                .name = DEFAULT_FONT
                .ColorIndex = 15
                .Size = 10
            End With

            ' логотип из файла расчёта
            With .cell(1, 1)
                ThisWorkbook.Worksheets(SERVICE_SHEET_NAME).shapes("logo").CopyPicture Appearance:=xlScreen, Format:=xlPicture
                .Range.Paste
                .Range.InlineShapes(1).ScaleHeight = 25
                .Range.InlineShapes(1).ScaleWidth = 25
                .SetWidth ColumnWidth:=Application.CentimetersToPoints(1.5), RulerStyle:=wdAdjustFirstColumn
            End With

            ' фирменное мотто
            With .cell(1, 2).Range
                .Text = TEXTS_MOTTO
                .ParagraphFormat.Alignment = wdAlignParagraphLeft
            End With

            ' адрес и контакты
            With .cell(1, 3).Range
                .Text = TEXTS_ADDRESS
                .ParagraphFormat.Alignment = wdAlignParagraphRight
                .Font.Size = 7
            End With
        End With

        ' от/куда
        Set wdRng = wdDoc.Content
        Set wdTbl = wdDoc.Tables.Add(Range:=wdRng, NumRows:=1, NumColumns:=2)
        With wdTbl
            With .Range.Font
                .name = DEFAULT_FONT
                .Size = 10
            End With

            ' от
            With .cell(1, 1).Range
                .Text = TEXTS_FROM & TEXTS_4X4_SHORT & vbCrLf & vbCrLf & TEXTS_REFERENCE
                .ParagraphFormat.Alignment = wdAlignParagraphLeft
                Set wdRng = wdDoc.Range(Len(TEXTS_FROM), .Paragraphs(1).Range.End)
                wdRng.Font.Bold = True

                ' поле даты
                Set wdRng = wdDoc.Range(.End, .End)
                wdRng.Collapse Direction:=wdCollapseEnd
                wdRng.MoveEnd Unit:=wdCharacter, Count:=-1
                wdDoc.Fields.Add Range:=wdRng, Type:=wdFieldCreateDate, Text:=DATE_FIELD_FORMAT
            End With

            ' кому
            With .cell(1, 2).Range
                .Text = TEXTS_WHOM & Application.Range(CUSTOMER_CELL_NAME).Value2
                .ParagraphFormat.Alignment = wdAlignParagraphRight
                Set wdRng = wdDoc.Range(.Paragraphs(1).Range.Start + Len(TEXTS_WHOM), .Paragraphs(1).Range.End)
                wdRng.Font.Bold = True
            End With
        End With

        ' заголовок документа
        Set wdRng = wdDoc.Content
        With wdRng
            .Collapse Direction:=wdCollapseEnd
            .Text = TEXTS_SALES_OFFER
            .Font.Bold = True
            .Font.name = ALTERNATIVE_FONT
            .Font.Size = 16
            .ParagraphFormat.SpaceAfter = 10
            .ParagraphFormat.SpaceBefore = 10
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
        End With

        ' питч
        wdDoc.Paragraphs.Add
        Set wdRng = wdDoc.Content
        With wdRng
            .Collapse Direction:=wdCollapseEnd
            .Text = TEXTS_4X4_LONG & " " & TEXTS_PITCH & vbCrLf
            .ParagraphFormat.Alignment = wdAlignParagraphJustify
        End With

        ' копируем спецификацию с листа расчёта продажи
        Set wdRng = wdDoc.Content
        With wdRng
            ThisWorkbook.Sheets(sheetName).Cells(ROW_OFFSET + 1, COLUMN_OFFSET + 1).CurrentRegion.Copy
            .Collapse Direction:=wdCollapseEnd
            wdRng.PasteAndFormat wdFormatOriginalFormatting
            wdRng.MoveEnd Unit:=wdCharacter, Count:=wdDoc.Content.Characters.Count - wdRng.Start
            wdRng.Tables(1).AutoFitBehavior Behavior:=wdAutoFitWindow
        End With

        ' Итоги прописью
        Set wdRng = wdDoc.Content
        With wdRng
            .Collapse Direction:=wdCollapseEnd
            .ParagraphFormat.Alignment = wdAlignParagraphJustify

            Select Case VATtype
                Case 1: .Text = TEXTS_TOTAL & ": " & revenue.Text & " " & currText & _
                    " (" & convertPriceToText(CDbl(revenue.Value2), calcCurrency) & _
                    "), " & TEXTS_SUBJECT_VAT & " " & VATamount.Text & " " & currText & _
                     " (" & convertPriceToText(CDbl(VATamount.Value2), calcCurrency) & _
                     ")." & vbCrLf
                Case 2: .Text = TEXTS_TOTAL & ": " & Format(revenue.Value2 + VATamount.Value2, "#,##0.00") & " " & currText & _
                    " (" & convertPriceToText(CDbl(revenue.Value2 + VATamount.Value2), calcCurrency) & _
                    "), " & TEXTS_SUBJECT_VAT & " " & VATamount.Text & " " & currText & _
                     " (" & convertPriceToText(CDbl(VATamount.Value2), calcCurrency) & _
                     ")." & vbCrLf
                Case 3: .Text = TEXTS_TOTAL & ": " & revenue.Text & " " & currText & _
                    " (" & convertPriceToText(CDbl(revenue.Value2), calcCurrency) & _
                    "), " & TEXTS_NOT_SUBJECT_VAT & "." & vbCrLf
            End Select

            .ParagraphFormat.SpaceAfter = 10
            .ParagraphFormat.SpaceBefore = 10
            .Font.Bold = True
        End With

        ' условия оплаты
        Set wdRng = wdDoc.Content
        With wdRng
            .Collapse Direction:=0
            .Text = TEXTS_TERMS_OF_PAYMENT & vbCrLf
            .ListFormat.ApplyNumberDefault

            arr = Application.Range(TERMS_OF_PAYMENT_ARRAY_NAME).Value2
            For i = 1 To Application.Range(TERMS_OF_PAYMENT_ARRAY_NAME).Rows.Count
                .Collapse Direction:=0
                temp = vbNullString

                If IsEmpty(arr(i, TermsOfPaymentColumns.typen)) Then
                    temp = temp & "___"
                Else
                    temp = temp & CStr(arr(i, TermsOfPaymentColumns.typen))
                End If
                temp = temp & " в размере "
                If IsEmpty(arr(i, TermsOfPaymentColumns.PART)) Then
                    temp = temp & "___"
                Else
                    temp = temp & CStr(arr(i, TermsOfPaymentColumns.PART) * 100) & "% от итоговой суммы, а именно "

                    Select Case VATtype
                        Case 1: temp = temp & Format(revenue.Value2 * arr(i, TermsOfPaymentColumns.PART), "#,##0.00") & _
                                        " " & currText & " (" & convertPriceToText(CDbl(revenue.Value2 * arr(i, TermsOfPaymentColumns.PART)), calcCurrency) & _
                                        "), " & TEXTS_SUBJECT_VAT & " " & Format(VATamount.Value2 * arr(i, TermsOfPaymentColumns.PART), "#,##0.00") & " " & currText & _
                                        " (" & convertPriceToText(CDbl(VATamount.Value2 * arr(i, TermsOfPaymentColumns.PART)), calcCurrency) & ")"
                        Case 2: temp = temp & Format((revenue.Value2 + VATamount.Value2) * arr(i, TermsOfPaymentColumns.PART), "#,##0.00") & _
                                        " " & currText & " (" & convertPriceToText(CDbl((revenue.Value2 + VATamount.Value2) * arr(i, TermsOfPaymentColumns.PART)), _
                                        calcCurrency) & "), " & TEXTS_SUBJECT_VAT & " " & Format(VATamount.Value2 * arr(i, TermsOfPaymentColumns.PART), "#,##0.00") & _
                                        " " & currText & " (" & convertPriceToText(CDbl(VATamount.Value2 * arr(i, TermsOfPaymentColumns.PART)), calcCurrency) & ")"
                        Case 3: temp = temp & Format(revenue.Value2 * arr(i, TermsOfPaymentColumns.PART), "#,##0.00") & " " & currText & _
                                        " (" & convertPriceToText(CDbl(revenue.Value2 * arr(i, TermsOfPaymentColumns.PART)), calcCurrency) & _
                                        "), " & TEXTS_NOT_SUBJECT_VAT
                    End Select
                End If

                temp = temp & ", в течение "
                If IsEmpty(arr(i, TermsOfPaymentColumns.TIMEAMOUNT)) Then
                    temp = temp & "___ "
                Else
                    temp = temp & CStr(arr(i, TermsOfPaymentColumns.TIMEAMOUNT)) & " "
                End If
                If IsEmpty(arr(i, TermsOfPaymentColumns.TIMETYPE)) Then
                    temp = temp & "___ "
                Else
                    temp = temp & CStr(arr(i, TermsOfPaymentColumns.TIMETYPE)) & " "
                End If
                If IsEmpty(arr(i, TermsOfPaymentColumns.TIMEDIMENSION)) Then
                    temp = temp & "___"
                Else
                    temp = temp & CStr(arr(i, TermsOfPaymentColumns.TIMEDIMENSION))
                End If
                temp = temp & " с момента "
                If IsEmpty(arr(i, TermsOfPaymentColumns.FROM)) Then
                    temp = temp & "___;"
                Else
                    temp = temp & CStr(arr(i, TermsOfPaymentColumns.FROM)) & ";"
                End If

                .Text = temp & vbCrLf
                .ListFormat.ApplyBulletDefault
                .ListFormat.ListIndent
            Next i


        End With

        ' условия поставки
        Set wdRng = wdDoc.Content
        With wdRng
            .Collapse Direction:=0
            .Text = TEXTS_TERMS_OF_SERVICE & vbCrLf
            .ListFormat.ApplyNumberDefault

            arr = Application.Range(TERMS_OF_SERVICE_ARRAY_NAME).Value2
            For i = 1 To Application.Range(TERMS_OF_SERVICE_ARRAY_NAME).Rows.Count
                .Collapse Direction:=0

                temp = vbNullString
                If IsEmpty(arr(i, TermsOfServiceColumns.PART)) Then
                    temp = temp & "___"
                Else
                    temp = temp & CStr(arr(i, TermsOfServiceColumns.PART) * 100) & "% от объёма спецификации"
                End If
                temp = temp & " в течение "
                If IsEmpty(arr(i, TermsOfServiceColumns.TIMEAMOUNT)) Then
                    temp = temp & "___ "
                Else
                    temp = temp & CStr(arr(i, TermsOfServiceColumns.TIMEAMOUNT)) & " "
                End If
                If IsEmpty(arr(i, TermsOfServiceColumns.TIMETYPE)) Then
                    temp = temp & "___ "
                Else
                    temp = temp & CStr(arr(i, TermsOfServiceColumns.TIMETYPE)) & " "
                End If
                If IsEmpty(arr(i, TermsOfServiceColumns.TIMEDIMENSION)) Then
                    temp = temp & "___"
                Else
                    temp = temp & CStr(arr(i, TermsOfServiceColumns.TIMEDIMENSION))
                End If
                temp = temp & " в "
                If IsEmpty(arr(i, TermsOfServiceColumns.CITY)) Then
                    temp = temp & "___"
                Else
                    temp = temp & "г. " & CStr(arr(i, TermsOfServiceColumns.CITY))
                End If
                temp = temp & " с момента "
                If IsEmpty(arr(i, TermsOfServiceColumns.FROM)) Then
                    temp = temp & "___"
                Else
                    temp = temp & CStr(arr(i, TermsOfServiceColumns.FROM)) & ";"
                End If

                .Text = temp & vbCrLf
                .ListFormat.ApplyBulletDefault
                .ListFormat.ListIndent
            Next i
        End With


        ' доставка включена в стоимость оборудования
        Set wdRng = wdDoc.Content
        With wdRng
            .Collapse Direction:=0

            If Application.Range(INCLUDE_DELIVERY_CELL_NAME).Value2 = "да" Then
                .Text = TEXTS_DELIVERY_INCLUDED & ";" & vbCrLf
            Else
                .Text = TEXTS_DELIVERY_NOT_INCLUDED & "." & vbCrLf
            End If

            .ListFormat.ApplyNumberDefault
        End With

        ' подпись и печать
        wdDoc.Paragraphs.Add
        Set wdRng = wdDoc.Content
        With wdRng
            .Collapse Direction:=0
            If Application.Range(PM_CELL_NAME).Value2 = vbNullString Then
                .Text = TEXTS_WORK_TITLE & vbCrLf & TEXTS_4X4_SHORT & vbCrLf & vbCrLf & TEXTS_SIGN & " / " & TEXTS_SIGN & _
                    " /" & vbCrLf & TEXTS_LOCUS_SIGILI
            ElseIf Application.WorksheetFunction.CountIf(Application.Range(MANAGERS_NAMES_ARRAY_NAME), Application.Range(PM_CELL_NAME).Value2) > 0 Then
                .Text = Application.Range(MANAGERS_TITLES_ARRAY_NAME).Cells(Application.Match(Application.Range(PM_CELL_NAME).Value2, _
                       Application.Range(MANAGERS_NAMES_ARRAY_NAME), 0)).Value2 & _
                       vbCrLf & TEXTS_4X4_SHORT & vbCrLf & vbCrLf & TEXTS_SIGN & " / " & _
                       Application.Range(PM_CELL_NAME).Value2 & " /" & vbCrLf & TEXTS_LOCUS_SIGILI
            Else
                .Text = TEXTS_WORK_TITLE & vbCrLf & TEXTS_4X4_SHORT & vbCrLf & vbCrLf & TEXTS_SIGN & " / " & _
                       Application.Range(PM_CELL_NAME).Value2 & " /" & vbCrLf & TEXTS_LOCUS_SIGILI
            End If

            .ParagraphFormat.Alignment = wdAlignParagraphLeft
            .ParagraphFormat.SpaceAfter = 0
            .ParagraphFormat.SpaceBefore = 0
        End With
    Else
        MsgBox "Лист '" & sheetName & "' не найден"
    End If

    sheetsListUpdate

CleanExit:
    Set wdApp = Nothing
    Set wdDoc = Nothing
    Set wdRng = Nothing
    Set wdTbl = Nothing
    Set wdFld = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error " & err.number & ": " & err.Description
    Resume CleanExit
End Sub

Public Sub initializeShapes()
    Dim shape As shape
    Dim mainLeft As Double, mainTop As Double, mainHeight As Double, mainWidth As Double, margin As Double, buttonHeight As Double, dropdownHeight As Double
    Dim topLeftCellRange As Range, bottomRightCellRange As Range

    Set topLeftCellRange = ThisWorkbook.Sheets(SPEC_SHEET_NAME).Cells(3, PurchaseColumns.VAT_SALES)
    Set bottomRightCellRange = ThisWorkbook.Sheets(SPEC_SHEET_NAME).Cells(12, PurchaseColumns.PROFIT_PERCENT)

    mainLeft = topLeftCellRange.Left
    mainTop = topLeftCellRange.Top
    mainHeight = bottomRightCellRange.Top + bottomRightCellRange.Height - mainTop
    mainWidth = 250
    margin = mainHeight / 37

    dropdownHeight = 14
    buttonHeight = (mainHeight - 7 * margin - dropdownHeight) / 4

    On Error Resume Next
    ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(CONTROL_GROUP_NAME).Placement = xlFreeFloating
    ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(CONTROL_GROUP_NAME).Visible = msoTrue
    ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(CONTROL_GROUP_NAME).Width = mainWidth

    With ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(BOARD_SHAPE_NAME)
        .Placement = xlFreeFloating
        .Left = mainLeft
        .Top = mainTop
        .Height = mainHeight
        .Width = mainWidth
        .Visible = msoTrue
        .Line.Weight = 1
        .Line.ForeColor.RGB = RGB(0, 0, 0)
        .Line.DashStyle = msoLineSolid
        .Line.Style = msoLineThinThick
        .Adjustments(1) = 0
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With

    With ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(CHECKBOXES_FRAME_NAME)
        .Placement = xlFreeFloating
        .Left = ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(BOARD_SHAPE_NAME).Left + 2 * margin
        .Top = ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(BOARD_SHAPE_NAME).Top + 1.5 * margin
        .Height = mainHeight - 3 * margin
        .Width = (mainWidth - 6 * margin) / 2
        .Visible = msoTrue
    End With

    With ThisWorkbook.Sheets(SPEC_SHEET_NAME)
        For Each shape In .shapes(CONTROL_GROUP_NAME).GroupItems
            If shape.Type = msoFormControl Then
                If shape.FormControlType = xlCheckBox Then
                    shape.Placement = xlFreeFloating
                    shape.Left = .shapes(CHECKBOXES_FRAME_NAME).Left + margin
                    shape.Top = .shapes(CHECKBOXES_FRAME_NAME).Top + 3 * CInt(shape.AlternativeText) * margin - 3 * margin / 2
                    shape.Height = margin
                    shape.Width = .shapes(CHECKBOXES_FRAME_NAME).Width - 2 * margin
                    shape.Visible = msoTrue
                End If
            End If
        Next shape
    End With

    With ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(CALC_BUTTON_SHAPE_NAME)
        .Placement = xlFreeFloating
        .Left = ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(CHECKBOXES_FRAME_NAME).Left + _
                ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(CHECKBOXES_FRAME_NAME).Width + _
                2 * margin
        .Top = ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(CHECKBOXES_FRAME_NAME).Top
        .Height = buttonHeight
        .Width = mainWidth - ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(CHECKBOXES_FRAME_NAME).Width - 6 * margin
        .Visible = msoTrue
    End With

    With ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(EXPORT_LABEL_SHAPE_NAME)
        .Placement = xlFreeFloating
        .Left = ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(CHECKBOXES_FRAME_NAME).Left + _
                ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(CHECKBOXES_FRAME_NAME).Width + _
                2 * margin
        .Top = ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(CALC_BUTTON_SHAPE_NAME).Top + _
               ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(CALC_BUTTON_SHAPE_NAME).Height + _
               margin
        .Height = buttonHeight
        .Width = mainWidth - ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(CHECKBOXES_FRAME_NAME).Width - 6 * margin
        .Visible = msoTrue
    End With

    With ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(DROPDOWN_SHAPE_NAME)
        .Placement = xlFreeFloating
        .Left = ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(CHECKBOXES_FRAME_NAME).Left + _
                ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(CHECKBOXES_FRAME_NAME).Width + _
                2 * margin
        .Top = ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(EXPORT_LABEL_SHAPE_NAME).Top + _
               ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(EXPORT_LABEL_SHAPE_NAME).Height
        .Height = dropdownHeight
        .Width = mainWidth - ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(CHECKBOXES_FRAME_NAME).Width - 6 * margin
        .Visible = msoTrue
    End With

    With ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(EXPORT_EXCEL_BUTTON_SHAPE_NAME)
        .Placement = xlFreeFloating
        .Left = ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(CHECKBOXES_FRAME_NAME).Left + _
                ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(CHECKBOXES_FRAME_NAME).Width + _
                2 * margin
        .Top = ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(DROPDOWN_SHAPE_NAME).Top + _
               ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(DROPDOWN_SHAPE_NAME).Height + _
               margin
        .Height = buttonHeight
        .Width = ((mainWidth - ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(CHECKBOXES_FRAME_NAME).Width - 6 * margin) - margin) / 2
        .Visible = msoTrue
    End With

    With ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(EXPORT_WORD_BUTTON_SHAPE_NAME)
        .Placement = xlFreeFloating
        .Left = ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(EXPORT_EXCEL_BUTTON_SHAPE_NAME).Left + _
                ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(EXPORT_EXCEL_BUTTON_SHAPE_NAME).Width + _
                margin
        .Top = ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(DROPDOWN_SHAPE_NAME).Top + _
               ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(DROPDOWN_SHAPE_NAME).Height + _
               margin
        .Height = buttonHeight
        .Width = ((mainWidth - ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(CHECKBOXES_FRAME_NAME).Width - 6 * margin) - margin / 2) / 2
        .Visible = msoTrue
    End With

        With ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(EXPORT_1C_BUTTON_SHAPE_NAME)
        .Placement = xlFreeFloating
        .Left = ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(CHECKBOXES_FRAME_NAME).Left + _
                ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(CHECKBOXES_FRAME_NAME).Width + _
                2 * margin
        .Top = ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(EXPORT_EXCEL_BUTTON_SHAPE_NAME).Top + _
               ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(EXPORT_EXCEL_BUTTON_SHAPE_NAME).Height + _
               margin
        .Height = buttonHeight
        .Width = mainWidth - ThisWorkbook.Sheets(SPEC_SHEET_NAME).shapes(CHECKBOXES_FRAME_NAME).Width - 6 * margin
        .Visible = msoTrue
    End With

CleanExit:
    Exit Sub

ErrorHandler:
    MsgBox "Error " & err.number & ": " & err.Description
    Resume CleanExit
End Sub
