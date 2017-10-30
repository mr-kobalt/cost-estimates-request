Attribute VB_Name = "ui_ux"
Option Explicit

Private Sub buttonPN_Click()
' заменяем кириллические символы в слобце P/N на латинские эквиваленты
    Dim i As Long
    
    changeUpdatingState False
    
    On Error GoTo ErrorHandler
    i = replaceCyrillicWithLatin(Application.Worksheets(PURCHASE_SHEET_NAME).Range(PURCHASE_TABLE_NAME).columns(PurchaseColumns.PN))
    
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
    trimAndClearRange Application.Worksheets(PURCHASE_SHEET_NAME).Range(PURCHASE_TABLE_NAME).columns(PurchaseColumns.NAME_AND_DESCRIPTION)
    
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

    Application.EnableEvents = False
    On Error GoTo ErrorHandler
    
    With Application.Worksheets(PURCHASE_SHEET_NAME)
        .ListObjects(DELIVERY_TABLE_NAME).Range.Cut Destination:=.Cells(.ListObjects(DELIVERY_TABLE_NAME).Range.Row + rowCount, 1)
        
        If pos = "last" Then
            Application.ActiveSheet.ListObjects(PURCHASE_TABLE_NAME).TotalsRowRange.EntireRow.Resize(rowCount).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Else
            Application.ActiveSheet.ListObjects(PURCHASE_TABLE_NAME).DataBodyRange.Rows(1).EntireRow.Resize(rowCount).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
        End If
        
        .ListObjects(DELIVERY_TABLE_NAME).Range.Cut Destination:=.Cells(.ListObjects(DELIVERY_TABLE_NAME).Range.Row - rowCount, 1)
        
        With .Range(PURCHASE_TABLE_NAME)
            Set leftBorderRange = Application.Union(.columns(PurchaseColumns.PRICE_GPL_RECALCULATED), _
                                                    .columns(PurchaseColumns.PRICE_PURCHASE_RECALCULATED), _
                                                    .columns(PurchaseColumns.GPL_CURRENCY), _
                                                    .columns(PurchaseColumns.PURCHASE_CURRENCY))
            Set rightBorderRange = Application.Union(.columns(PurchaseColumns.TOTAL_GPL_RECALCULATED), _
                                                     .columns(PurchaseColumns.VAT_PURCHASE), _
                                                     .columns(PurchaseColumns.PRICE_GPL), _
                                                     .columns(PurchaseColumns.VAT))
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
        End With
    End With
    
CleanExit:
    Set leftBorderRange = Nothing
    Set rightBorderRange = Nothing
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & err.number & ": " & err.Description
    Resume CleanExit
End Sub


Private Sub CurrencyComboBox_Change()
' Если вводим данные в выпадающий список CurrencyComboBox на листе Расчёт продажи (SALES_SHEET_NAME),
' то копируем новое значение в ячейку с Валютой расчёта (CALC_CURRENCY_CELL_NAME)

    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    Application.Worksheets(PURCHASE_SHEET_NAME).Range(CALC_CURRENCY_CELL_NAME).Value2 = _
                        Application.Range(CURRENCIES_ARRAY_NAME).Cells(Application.Worksheets(SALES_SHEET_NAME) _
                       .shapes(CURRENCY_SHAPE_NAME).OLEFormat.Object.Value)

CleanExit:
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & err.number & ": " & err.Description
    Resume CleanExit
End Sub

Private Sub DeliveryCostsCheckbox_Change()
' Если щёлкаем на чекбокс DeliveryCosts на листе Расчёт продажи (SALES_SHEET_NAME),
' то копируем новое значение в ячейку с информацией о включении транспортных
' расходов (INCLUDE_DELIVERY_CELL_NAME) в КП

    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    If Application.Worksheets(SALES_SHEET_NAME).shapes(DELIVERY_SHAPE_NAME).OLEFormat.Object.Value = xlOn Then
        Application.Worksheets(PURCHASE_SHEET_NAME).Range(INCLUDE_DELIVERY_CELL_NAME).Value2 = YES
    Else
        Application.Worksheets(PURCHASE_SHEET_NAME).Range(INCLUDE_DELIVERY_CELL_NAME).Value2 = NO
    End If

CleanExit:
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & err.number & ": " & err.Description
    Resume CleanExit
End Sub

Private Sub VATComboBox_Change()
' Если вводим данные в выпадающий список VATComboBox на листе Расчёт продажи (SALES_SHEET_NAME),
' то копируем новое значение в ячейку с индикатором включения НДС (INCLUDE_VAT_CELL_NAME)

    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    With Application.Worksheets(PURCHASE_SHEET_NAME)
        .Range(INCLUDE_VAT_CELL_NAME).Value2 = _
                    Application.Range(VAT_ARRAY_NAME).Cells(Application.Worksheets(SALES_SHEET_NAME) _
                    .shapes(VAT_SHAPE_NAME).OLEFormat.Object.Value)
        .Range(PURCHASE_TABLE_NAME).columns(PurchaseColumns.VAT_PURCHASE).Value2 = _
                    .Range(INCLUDE_VAT_CELL_NAME).Value2
    End With
CleanExit:
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & err.number & ": " & err.Description
    Resume CleanExit
End Sub

Private Sub SalesColumns_Click()
' скрываем столбцы КП при переключении соответствующих чекбоксов
    Dim shape As shape
    Dim column As Long
    
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    Set shape = Application.Worksheets(SALES_SHEET_NAME).shapes(Application.Caller)
    
    column = findColNumber(getDesiredColumns(), shape.AlternativeText) + COLUMN_OFFSET
    
    With Application.Worksheets(SALES_SHEET_NAME).Cells.columns(column)
        If shape.OLEFormat.Object.Value = xlOff Then
            .Hidden = True
        ElseIf shape.OLEFormat.Object.Value = xlOn Then
            .Hidden = False
        End If
    End With
CleanExit:
    Application.EnableEvents = True
    Set shape = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & err.number & ": " & err.Description
    Resume CleanExit
End Sub

Private Sub ProfitButton_Click()
' открываем форму ручного ввода процента прибыли
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    ProfitDialog.Show
    
CleanExit:
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & err.number & ": " & err.Description
    Resume CleanExit
End Sub

Private Sub excelExportButton_Click()
    Dim wb As Workbook
    
    On Error GoTo ErrorHandler
    Set wb = Application.Workbooks.Add
    ThisWorkbook.Worksheets(SALES_SHEET_NAME).Cells(ROW_OFFSET + 1, COLUMN_OFFSET + 1).CurrentRegion.Copy
    
    wb.ActiveSheet.Cells(1).PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone
    wb.ActiveSheet.Cells(1).PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlPasteSpecialOperationNone
    wb.ActiveSheet.Cells(1).PasteSpecial Paste:=xlPasteFormats, Operation:=xlPasteSpecialOperationNone
    
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
    
    Dim i As Long
    Dim temp As String, currText As String
    Dim arr() As Variant
    
    On Error GoTo ErrorHandler
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
        ThisWorkbook.Worksheets(SALES_SHEET_NAME).Cells(ROW_OFFSET + 1, COLUMN_OFFSET + 1).CurrentRegion.Copy
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
        currText = Application.Range(CURRENCIES_HEADER_ARRAY_NAME).Cells(Application.Match(Application.Range(CALC_CURRENCY_CELL_NAME).Value2, _
                                                                                           Application.Range(CURRENCIES_ARRAY_NAME), 0)).Value2
        Select Case Application.Match(Application.Range(INCLUDE_VAT_CELL_NAME), Application.Range(VAT_ARRAY_NAME), 0)
            Case 1: .Text = TEXTS_TOTAL & ": " & Application.Range(REVENUE_CELL_NAME).Text & " " & currText & _
                " (" & convertPriceToText(CDbl(Application.Range(REVENUE_CELL_NAME).Value2), Application.Range(CALC_CURRENCY_CELL_NAME).Value2) & _
                "), " & TEXTS_SUBJECT_VAT & " " & Application.Range(VAT_AMOUNT_CELL_NAME).Text & " " & currText & _
                 " (" & convertPriceToText(CDbl(Application.Range(VAT_AMOUNT_CELL_NAME).Value2), Application.Range(CALC_CURRENCY_CELL_NAME).Value2) & _
                 ")." & vbCrLf
            Case 2: .Text = TEXTS_TOTAL & ": " & Format(Application.Range(REVENUE_CELL_NAME).Value2 + Application.Range(VAT_AMOUNT_CELL_NAME).Value2, "# ##0.00") & " " & currText & _
                " (" & convertPriceToText(CDbl(Application.Range(REVENUE_CELL_NAME).Value2 + Application.Range(VAT_AMOUNT_CELL_NAME).Value2), Application.Range(CALC_CURRENCY_CELL_NAME).Value2) & _
                "), " & TEXTS_SUBJECT_VAT & " " & Application.Range(VAT_AMOUNT_CELL_NAME).Text & " " & currText & _
                 " (" & convertPriceToText(CDbl(Application.Range(VAT_AMOUNT_CELL_NAME).Value2), Application.Range(CALC_CURRENCY_CELL_NAME).Value2) & _
                 ")." & vbCrLf
            Case 3: .Text = TEXTS_TOTAL & ": " & Application.Range(REVENUE_CELL_NAME).Text & " " & currText & _
                " (" & convertPriceToText(CDbl(Application.Range(REVENUE_CELL_NAME).Value2), Application.Range(CALC_CURRENCY_CELL_NAME).Value2) & _
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
            'wdDoc.Paragraphs.Add
            .Collapse Direction:=0
            'wdRng.MoveEnd Unit:=wdCharacter, Count:=1
            
            temp = vbNullString
            If IsEmpty(arr(i, TermsOfPaymentColumns.typename)) Then
                temp = temp & "___"
            Else
                temp = temp & CStr(arr(i, TermsOfPaymentColumns.typename))
            End If
            temp = temp & " в размере "
            If IsEmpty(arr(i, TermsOfPaymentColumns.PART)) Then
                temp = temp & "___"
            Else
                temp = temp & CStr(arr(i, TermsOfPaymentColumns.PART) * 100) & "% от итоговой суммы, а именно "
                currText = Application.Range(CURRENCIES_HEADER_ARRAY_NAME).Cells(Application.Match(Application.Range(CALC_CURRENCY_CELL_NAME).Value2, _
                                                                                               Application.Range(CURRENCIES_ARRAY_NAME), 0)).Value2
                Select Case Application.Match(Application.Range(INCLUDE_VAT_CELL_NAME), Application.Range(VAT_ARRAY_NAME), 0)
                    Case 1: temp = temp & Format(Application.Range(REVENUE_CELL_NAME).Value2 * arr(i, TermsOfPaymentColumns.PART), "# ##0.00") & _
                                    " " & currText & " (" & convertPriceToText(CDbl(Application.Range(REVENUE_CELL_NAME).Value2 * arr(i, TermsOfPaymentColumns.PART)), Application.Range(CALC_CURRENCY_CELL_NAME).Value2) & _
                                    "), " & TEXTS_SUBJECT_VAT & " " & Format(Application.Range(VAT_AMOUNT_CELL_NAME).Value2 * arr(i, TermsOfPaymentColumns.PART), "# ##0.00") & " " & currText & _
                                    " (" & convertPriceToText(CDbl(Application.Range(VAT_AMOUNT_CELL_NAME).Value2 * arr(i, TermsOfPaymentColumns.PART)), Application.Range(CALC_CURRENCY_CELL_NAME).Value2) & ")"
                    Case 2: temp = temp & Format((Application.Range(REVENUE_CELL_NAME).Value2 + Application.Range(VAT_AMOUNT_CELL_NAME).Value2) * arr(i, TermsOfPaymentColumns.PART), "# ##0.00") & _
                                    " " & currText & " (" & convertPriceToText(CDbl((Application.Range(REVENUE_CELL_NAME).Value2 + Application.Range(VAT_AMOUNT_CELL_NAME).Value2) * arr(i, TermsOfPaymentColumns.PART)), _
                                    Application.Range(CALC_CURRENCY_CELL_NAME).Value2) & "), " & TEXTS_SUBJECT_VAT & " " & Format(Application.Range(VAT_AMOUNT_CELL_NAME).Value2 * arr(i, TermsOfPaymentColumns.PART), "# ##0.00") & _
                                    " " & currText & " (" & convertPriceToText(CDbl(Application.Range(VAT_AMOUNT_CELL_NAME).Value2 * arr(i, TermsOfPaymentColumns.PART)), Application.Range(CALC_CURRENCY_CELL_NAME).Value2) & ")"
                    Case 3: temp = temp & Format(Application.Range(REVENUE_CELL_NAME).Value2 * arr(i, TermsOfPaymentColumns.PART), "# ##0.00") & " " & currText & _
                                    " (" & convertPriceToText(CDbl(Application.Range(REVENUE_CELL_NAME).Value2 * arr(i, TermsOfPaymentColumns.PART)), Application.Range(CALC_CURRENCY_CELL_NAME).Value2) & _
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
        .Text = Application.Range(MANAGERS_TITLES_ARRAY_NAME).Cells(Application.Match(Application.Range(PM_CELL_NAME).Value2, _
                                                                                      Application.Range(MANAGERS_NAMES_ARRAY_NAME), 0)).Value2 & _
                vbCrLf & TEXTS_4X4_SHORT & vbCrLf & vbCrLf & TEXTS_SIGN & " / " & Application.Range(PM_CELL_NAME).Value2 & _
                " /" & vbCrLf & TEXTS_LOCUS_SIGILI
        .ParagraphFormat.Alignment = wdAlignParagraphLeft
        .ParagraphFormat.SpaceAfter = 0
        .ParagraphFormat.SpaceBefore = 0
    End With
    
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
    Dim mainLeft As Long, mainTop As Long, mainHeight As Long, mainWidth As Long, margin As Long
    
    mainLeft = 60
    mainTop = 20
    mainHeight = 234
    mainWidth = 310
    margin = 8
    
    On Error Resume Next
    With Application.Worksheets(SALES_SHEET_NAME).shapes(BOARD_SHAPE_NAME)
        .Placement = xlFreeFloating
        .Left = mainLeft
        .Top = mainTop
        .Height = mainHeight
        .Width = mainWidth
        .Visible = msoTrue
        .Line.Weight = 3
        .Line.ForeColor.RGB = RGB(0, 0, 0)
        .Line.DashStyle = msoLineSolid
        .Line.Style = msoLineThinThick
        .Adjustments(1) = 0.02
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With
    
    With Application.Worksheets(SALES_SHEET_NAME).shapes(CHECKBOXES_SUBGROUP_NAME)
        .Placement = xlFreeFloating
        .Left = mainLeft + margin
        .Top = mainTop + margin
        .Height = (mainHeight - 3 * margin) * 0.84
        .Width = 120
        .Visible = msoTrue
    End With
    
    For Each shape In Application.Worksheets(SALES_SHEET_NAME).shapes(CHECKBOXES_GROUP_NAME).GroupItems
        If shape.FormControlType = xlCheckBox Then
            With Application.Worksheets(SALES_SHEET_NAME).shapes(CHECKBOXES_SUBGROUP_NAME)
                shape.Placement = xlFreeFloating
                shape.Left = .Left + margin
                shape.Top = .Top + 2 * CInt(shape.AlternativeText) * margin - margin / 2
                shape.Height = margin
                shape.Width = .Width - 2 * margin
                shape.Visible = msoTrue
            End With
        End If
    Next shape
    
    With Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_PARAMS_SUBGROUP_NAME)
        .Placement = xlFreeFloating
        .Left = Application.Worksheets(SALES_SHEET_NAME).shapes(CHECKBOXES_SUBGROUP_NAME).Left + _
                Application.Worksheets(SALES_SHEET_NAME).shapes(CHECKBOXES_SUBGROUP_NAME).Width + _
                margin
        .Top = mainTop + margin
        .Height = ((mainHeight - 3 * margin) * 0.84 - margin) / 2
        .Width = mainWidth + mainLeft - .Left - margin
        .Visible = msoTrue
    End With
    
    With Application.Worksheets(SALES_SHEET_NAME).shapes(CURRENCY_LABEL_SHAPE_NAME)
        .Placement = xlFreeFloating
        .Left = Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_PARAMS_SUBGROUP_NAME).Left + _
                margin
        .Top = Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_PARAMS_SUBGROUP_NAME).Top + margin * 3 / 2
        .Height = margin * 3 / 2
        .Width = (Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_PARAMS_SUBGROUP_NAME).Width - 3 * margin) / 2
        .Visible = msoTrue
    End With
    
    With Application.Worksheets(SALES_SHEET_NAME).shapes(VAT_LABEL_SHAPE_NAME)
        .Placement = xlFreeFloating
        .Left = Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_PARAMS_SUBGROUP_NAME).Left + _
                margin
        .Top = Application.Worksheets(SALES_SHEET_NAME).shapes(CURRENCY_LABEL_SHAPE_NAME).Top + _
               Application.Worksheets(SALES_SHEET_NAME).shapes(CURRENCY_LABEL_SHAPE_NAME).Height + _
               margin
        .Height = margin * 3 / 2
        .Width = (Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_PARAMS_SUBGROUP_NAME).Width - 3 * margin) / 2
        .Visible = msoTrue
    End With
    
    With Application.Worksheets(SALES_SHEET_NAME).shapes(CURRENCY_SHAPE_NAME)
        .Placement = xlFreeFloating
        .Left = Application.Worksheets(SALES_SHEET_NAME).shapes(CURRENCY_LABEL_SHAPE_NAME).Left + _
                Application.Worksheets(SALES_SHEET_NAME).shapes(CURRENCY_LABEL_SHAPE_NAME).Width + _
                margin
        .Top = Application.Worksheets(SALES_SHEET_NAME).shapes(CURRENCY_LABEL_SHAPE_NAME).Top
        .Height = margin * 3 / 2
        .Width = (Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_PARAMS_SUBGROUP_NAME).Width - 3 * margin) / 2
        .Visible = msoTrue
    End With
    
    With Application.Worksheets(SALES_SHEET_NAME).shapes(VAT_SHAPE_NAME)
        .Placement = xlFreeFloating
        .Left = Application.Worksheets(SALES_SHEET_NAME).shapes(VAT_LABEL_SHAPE_NAME).Left + _
                Application.Worksheets(SALES_SHEET_NAME).shapes(VAT_LABEL_SHAPE_NAME).Width + _
                margin
        .Top = Application.Worksheets(SALES_SHEET_NAME).shapes(VAT_LABEL_SHAPE_NAME).Top
        .Height = margin * 3 / 2
        .Width = (Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_PARAMS_SUBGROUP_NAME).Width - 3 * margin) / 2
        .Visible = msoTrue
    End With
    
    With Application.Worksheets(SALES_SHEET_NAME).shapes(DELIVERY_SHAPE_NAME)
        .Placement = xlFreeFloating
        .Left = Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_PARAMS_SUBGROUP_NAME).Left + _
                margin
        .Top = Application.Worksheets(SALES_SHEET_NAME).shapes(VAT_LABEL_SHAPE_NAME).Top + _
               Application.Worksheets(SALES_SHEET_NAME).shapes(VAT_LABEL_SHAPE_NAME).Height + _
               margin
        .Height = margin * 7 / 2
        .Width = Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_PARAMS_SUBGROUP_NAME).Width - 2 * margin
        .Visible = msoTrue
    End With
    
    With Application.Worksheets(SALES_SHEET_NAME).shapes(PROFIT_SUBGROUP_NAME)
        .Placement = xlFreeFloating
        .Left = Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_PARAMS_SUBGROUP_NAME).Left
        .Top = Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_PARAMS_SUBGROUP_NAME).Top + _
               Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_PARAMS_SUBGROUP_NAME).Height + _
               margin
        .Height = ((mainHeight - 3 * margin) * 0.84 - margin) / 2
        .Width = mainWidth + mainLeft - .Left - margin
        .Visible = msoTrue
    End With
    
    With Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_TYPE_SUBGROUP_NAME)
        .Placement = xlFreeFloating
        .Left = Application.Worksheets(SALES_SHEET_NAME).shapes(PROFIT_SUBGROUP_NAME).Left + margin
        .Top = Application.Worksheets(SALES_SHEET_NAME).shapes(PROFIT_SUBGROUP_NAME).Top + margin * 3 / 2
        .Height = (Application.Worksheets(SALES_SHEET_NAME).shapes(PROFIT_SUBGROUP_NAME).Height - 3 * margin) * 2 / 3
        .Width = (Application.Worksheets(SALES_SHEET_NAME).shapes(PROFIT_SUBGROUP_NAME).Width - 3 * margin) / 2
        .Visible = msoTrue
    End With
    
    With Application.Worksheets(SALES_SHEET_NAME).shapes(MARKUP_SHAPE_NAME)
        .Placement = xlFreeFloating
        .Left = Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_TYPE_SUBGROUP_NAME).Left + _
                margin
        .Top = Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_TYPE_SUBGROUP_NAME).Top + margin
        .Height = margin
        .Width = Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_TYPE_SUBGROUP_NAME).Width - 2 * margin
        .Visible = msoTrue
    End With
    
    With Application.Worksheets(SALES_SHEET_NAME).shapes(MARGIN_SHAPE_NAME)
        .Placement = xlFreeFloating
        .Left = Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_TYPE_SUBGROUP_NAME).Left + _
                margin
        .Top = Application.Worksheets(SALES_SHEET_NAME).shapes(MARKUP_SHAPE_NAME).Top + _
               margin * 3 / 2
        .Height = margin
        .Width = Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_TYPE_SUBGROUP_NAME).Width - 2 * margin
        .Visible = msoTrue
    End With
    
    
    With Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_SOURCE_SUBGROUP_NAME)
        .Placement = xlFreeFloating
        .Left = Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_TYPE_SUBGROUP_NAME).Left + _
                Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_TYPE_SUBGROUP_NAME).Width + _
                margin
        .Top = Application.Worksheets(SALES_SHEET_NAME).shapes(PROFIT_SUBGROUP_NAME).Top + margin * 3 / 2
        .Height = (Application.Worksheets(SALES_SHEET_NAME).shapes(PROFIT_SUBGROUP_NAME).Height - 3 * margin) * 2 / 3
        .Width = (Application.Worksheets(SALES_SHEET_NAME).shapes(PROFIT_SUBGROUP_NAME).Width - 3 * margin) / 2
        .Visible = msoTrue
    End With
    
    With Application.Worksheets(SALES_SHEET_NAME).shapes(GPL_SHAPE_NAME)
        .Placement = xlFreeFloating
        .Left = Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_SOURCE_SUBGROUP_NAME).Left + _
                margin
        .Top = Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_SOURCE_SUBGROUP_NAME).Top + margin
        .Height = margin
        .Width = Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_SOURCE_SUBGROUP_NAME).Width - 2 * margin
        .Visible = msoTrue
    End With
    
    With Application.Worksheets(SALES_SHEET_NAME).shapes(NET_PRICE_SHAPE_NAME)
        .Placement = xlFreeFloating
        .Left = Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_SOURCE_SUBGROUP_NAME).Left + _
                margin
        .Top = Application.Worksheets(SALES_SHEET_NAME).shapes(GPL_SHAPE_NAME).Top + _
               margin * 3 / 2
        .Height = margin
        .Width = Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_SOURCE_SUBGROUP_NAME).Width - 2 * margin
        .Visible = msoTrue
    End With
    
    
    With Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_LABEL_SHAPE_NAME)
        .Placement = xlFreeFloating
        .Left = Application.Worksheets(SALES_SHEET_NAME).shapes(PROFIT_SUBGROUP_NAME).Left + margin / 2
        .Top = Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_TYPE_SUBGROUP_NAME).Top + _
               Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_TYPE_SUBGROUP_NAME).Height + _
               margin
        .Height = (Application.Worksheets(SALES_SHEET_NAME).shapes(PROFIT_SUBGROUP_NAME).Height - 3 * margin) / 3
        .Width = (Application.Worksheets(SALES_SHEET_NAME).shapes(PROFIT_SUBGROUP_NAME).Width - 3 * margin) / 2 + margin
        .Visible = msoTrue
    End With
    
    With Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_BUTTON_SHAPE_NAME)
        .Placement = xlFreeFloating
        .Left = Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_LABEL_SHAPE_NAME).Left + _
                Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_LABEL_SHAPE_NAME).Width + _
                margin / 2
        .Top = Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_TYPE_SUBGROUP_NAME).Top + _
               Application.Worksheets(SALES_SHEET_NAME).shapes(CALC_TYPE_SUBGROUP_NAME).Height + _
               margin
        .Height = (Application.Worksheets(SALES_SHEET_NAME).shapes(PROFIT_SUBGROUP_NAME).Height - 3 * margin) / 3
        .Width = (Application.Worksheets(SALES_SHEET_NAME).shapes(PROFIT_SUBGROUP_NAME).Width - 3 * margin) / 2
        .Visible = msoTrue
    End With
    

    With Application.Worksheets(SALES_SHEET_NAME).shapes(EXPORT_SUBGROUP_NAME)
        .Placement = xlFreeFloating
        .Left = Application.Worksheets(SALES_SHEET_NAME).shapes(CHECKBOXES_SUBGROUP_NAME).Left
        .Top = Application.Worksheets(SALES_SHEET_NAME).shapes(CHECKBOXES_SUBGROUP_NAME).Top + _
               Application.Worksheets(SALES_SHEET_NAME).shapes(CHECKBOXES_SUBGROUP_NAME).Height + _
               margin
        .Height = (mainHeight - 3 * margin) * 0.16
        .Width = mainWidth - 2 * margin
        .Visible = msoTrue
    End With
    
    With Application.Worksheets(SALES_SHEET_NAME).shapes(EXPORT_LABEL_SHAPE_NAME)
        .Placement = xlFreeFloating
        .Left = Application.Worksheets(SALES_SHEET_NAME).shapes(EXPORT_SUBGROUP_NAME).Left + margin / 2
        .Top = Application.Worksheets(SALES_SHEET_NAME).shapes(EXPORT_SUBGROUP_NAME).Top + margin * 3 / 2
        .Height = Application.Worksheets(SALES_SHEET_NAME).shapes(EXPORT_SUBGROUP_NAME).Height - 2 * margin
        .Width = Application.Worksheets(SALES_SHEET_NAME).shapes(CHECKBOXES_SUBGROUP_NAME).Width - margin / 2
        .Visible = msoTrue
    End With

    With Application.Worksheets(SALES_SHEET_NAME).shapes(EXPORT_WORD_BUTTON_SHAPE_NAME)
        .Placement = xlFreeFloating
        .Left = Application.Worksheets(SALES_SHEET_NAME).shapes(PROFIT_SUBGROUP_NAME).Left + margin
        .Top = Application.Worksheets(SALES_SHEET_NAME).shapes(EXPORT_LABEL_SHAPE_NAME).Top - margin / 2
        .Height = Application.Worksheets(SALES_SHEET_NAME).shapes(EXPORT_LABEL_SHAPE_NAME).Height + margin / 2
        .Width = (Application.Worksheets(SALES_SHEET_NAME).shapes(PROFIT_SUBGROUP_NAME).Width - 3 * margin) / 2
        .Visible = msoTrue
    End With

    With Application.Worksheets(SALES_SHEET_NAME).shapes(EXPORT_EXCEL_BUTTON_SHAPE_NAME)
        .Placement = xlFreeFloating
        .Left = Application.Worksheets(SALES_SHEET_NAME).shapes(EXPORT_WORD_BUTTON_SHAPE_NAME).Left + _
                Application.Worksheets(SALES_SHEET_NAME).shapes(EXPORT_WORD_BUTTON_SHAPE_NAME).Width + margin
        .Top = Application.Worksheets(SALES_SHEET_NAME).shapes(EXPORT_LABEL_SHAPE_NAME).Top - margin / 2
        .Height = Application.Worksheets(SALES_SHEET_NAME).shapes(EXPORT_LABEL_SHAPE_NAME).Height + margin / 2
        .Width = Application.Worksheets(SALES_SHEET_NAME).shapes(EXPORT_WORD_BUTTON_SHAPE_NAME).Width
        .Visible = msoTrue
    End With

CleanExit:
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & err.number & ": " & err.Description
    Resume CleanExit
End Sub
