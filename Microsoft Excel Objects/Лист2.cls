VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "����2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
' -------------------------------------------------------------------------------- '
' ��������� �������� ��������� � ������ ����� �������������
' 1. ��������� ����� ��������� (��������, �� �������� ������)
' 2. ����� �������� ����������� ��� ������ � �������, �������������� �� ����������
' 3. ����� �������� ����������� � ������� � ������ (����� � ����) ��������
'    �������� ������������� � ����� � ��������� ������ (��. convertCellsValueToDbl)
' 4. ����� ��������� ����� � ����� �� �� �� � ������ USD/EUR �� ��� ���� ���� �
'    ���������� �������� ���� ����� � ����� �� ��
' -------------------------------------------------------------------------------- '
    Dim isect As Range, priceRange As Range
    Dim undoControl As Object
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    'Application.EnableEvents = False
    changeUpdatingState False

    ' ������������� ������� � ��������� � ������ ������ �������� � �������, �� �� �������������� � ��.
    ' �������������� ������� ������������� �����:
    ' http://www.siddharthrout.com/2011/08/15/vba-excelallow-paste-special-only/
    
    ' ��������� ������� ���������� �� ������ ��������� ��������. 128 - ��� ������ �� ���������
    Set undoControl = Application.CommandBars("Standard").FindControl(ID:=128)
    
    ' ���� ���������� ����� Control �� ������� 128 ��� ��� �������� ��� ������, �� ��������� ���������
    If undoControl.Enabled = True Then
        If undoControl.ListCount <> 0 Then
            ' ���� ��� ���������� �������� �� ��� ��������� ��������. ������ �������� ��� ������� (�����)
            ' � ���������� (����� �� ������������, �� ���� �� ������ ���� ������ ��������) �����������.
            ' TODO: ����� ����� ������� ����� ����������� ���������� ��������
            If Left(undoControl.list(1), 5) = "Paste" Or Left(undoControl.list(1), 7) = "�������" Then
                ' �������� ��������� �������� ��� ������� ������ ������
                Application.Undo
                ' �������� �������, � ������� ���������� ���������� �������
                Target.Select
                ' ��������� ������ ��� ��������������
                On Error GoTo CutHandler
                Select Case Application.CutCopyMode
                    Case 0
                        ' ���� �������� ����������� �� ������ Excel, �� �������� �������� ��� HTML ��� ��������������,
                        ' � � ������ ������� - ������ ��������� ��� ����
                        On Error Resume Next
                        ActiveSheet.PasteSpecial Format:="HTML", Link:=False, DisplayAsIcon:=False, NoHTMLFormatting:=True
                        If Err.number <> 0 Then: Application.ActiveSheet.Paste
                        Err.Clear
                    Case Else
                        ' ���� Excel � ������ ����������� ��� ��������� �����, �� ��������� ������ ��������. ���� ������
                        ' ���� ���� ��������, �� PasteSpecial ���������� �������� ������ � ���������� ����� �������� �
                        ' ���������� CutHandler. �.�. ������ ������ �� ������� ���������� ����� � ����� ������ � ����.
                        Target.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone
                End Select
            ElseIf Left(undoControl.list(1), 4) = "Fill" Or Left(undoControl.list(1), 10) = "��������������" Then
                Application.Undo
                MsgBox "���������� ����� �� ���� ����� ���������. ����������� �����������, ������� � ��������"
                GoTo CleanExit
            End If
        End If
    End If

    With Application.Worksheets(PURCHASE_SHEET_NAME)
        On Error GoTo ErrorHandler2
        With .Range(PURCHASE_TABLE_NAME)
            Set priceRange = Union(.columns(PurchaseColumns.PRICE_GPL), _
                                   .columns(PurchaseColumns.PRICE_PURCHASE))
        End With
        On Error GoTo ErrorHandler
        
        ' ���� ��������� ��� ������ ������ � ������� � ������, �� ��������������
        ' �������� ������������� �� � �����
        Set isect = Application.Intersect(Target, priceRange)
        If Not (isect Is Nothing) Then: convertCellsValueToDbl isect
        
        
        ' ���� ��������� ��� ������ ������ � ������ � ������� ������� (CALC_CURRENCY_CELL_NAME),
        ' �� �������� ����� �������� � ���������� ������ �� ����� ������ ������� (SALES_SHEET_NAME)
        Set isect = Application.Intersect(Target, .Range(CALC_CURRENCY_CELL_NAME))
        If Not (isect Is Nothing) Then
            Application.Worksheets(SALES_SHEET_NAME).Shapes(CURRENCY_SHAPE_NAME).OLEFormat.Object.Value = _
                        Application.Match(.Range(CALC_CURRENCY_CELL_NAME).Value2, _
                        Application.Worksheets(SERVICE_SHEET_NAME).Range(CURRENCIES_ARRAY_NAME).Value2, 0)
        End If
        
        ' ���� ��������� ��� ������ ������ � ������-��������� ��������� ��� � ������ (INCLUDE_VAT_CELL_NAME),
        ' �� �������� ����� �������� � ���������� ������ �� ����� ������ ������� (SALES_SHEET_NAME)
        Set isect = Application.Intersect(Target, .Range(INCLUDE_VAT_CELL_NAME))
        If Not (isect Is Nothing) Then
            Application.Worksheets(SALES_SHEET_NAME).Shapes(VAT_SHAPE_NAME).OLEFormat.Object.Value = _
                        Application.Match(.Range(INCLUDE_VAT_CELL_NAME).Value2, _
                        Application.Worksheets(SERVICE_SHEET_NAME).Range(VAT_ARRAY_NAME).Value2, 0)
            createValidation Target:=.Range(PURCHASE_TABLE_NAME).columns(PurchaseColumns.VAT_PURCHASE), typename:="VAT"
            With .Range(PURCHASE_TABLE_NAME).columns(PurchaseColumns.VAT_PURCHASE)
                For i = 1 To .Cells.Count
                    If .Cells(i).Value2 <> ThisWorkbook.Sheets(SERVICE_SHEET_NAME).Range(VAT_ARRAY_NAME).Cells(3).Value2 And _
                                            .Cells(i).Value2 <> vbNullString Then
                        .Cells(i) = isect.Value2
                    End If
                Next i
            End With
            
        End If
        
        ' ���� ��������� ��� ������ ������ � ������-��������� ��������� ��������� �������� ������ (INCLUDE_DELIVERY_CELL_NAME),
        ' �� ������ ������� �� ����� ������ ������� (SALES_SHEET_NAME)
        Set isect = Application.Intersect(Target, .Range(INCLUDE_DELIVERY_CELL_NAME))
        If Not (isect Is Nothing) Then
            If isect.Value2 = YES Then
                Application.Worksheets(SALES_SHEET_NAME).Shapes(DELIVERY_SHAPE_NAME).OLEFormat.Object.Value = xlOn
            Else
                Application.Worksheets(SALES_SHEET_NAME).Shapes(DELIVERY_SHAPE_NAME).OLEFormat.Object.Value = xlOff
            End If
        End If
        
        ' ���� �������� ������ � ����� �����, �� ���������� �������� � ����� �� ��
        Set isect = Application.Intersect(Target, .Range(CURRENT_RATE_DATE_CELL_NAME))
        If Not (isect Is Nothing) Then: isect.Value2 = "���� �� �� �� " & parseXML(CBR_XML_URL, CURRENT_RATE_DATE_XPATH)
        
        ' ���� �������� ������ � ������� ������ USD, �� ���������� �������� � ����� �� ��
        Set isect = Application.Intersect(Target, .Range(USD_RATE_CELL_NAME))
        If Not (isect Is Nothing) Then: isect.Value2 = Replace(parseXML(CBR_XML_URL, USD_RATE_XPATH), ",", ".")
        
        ' ���� �������� ������ � ������� ������ EUR, �� ���������� �������� � ����� �� ��
        Set isect = Application.Intersect(Target, .Range(EUR_RATE_CELL_NAME))
        If Not (isect Is Nothing) Then: isect.Value2 = Replace(parseXML(CBR_XML_URL, EUR_RATE_XPATH), ",", ".")
        
        ' ���� �������� ������ � ��������, �� ������������ �� ����������
        Set isect = Application.Intersect(Target, .Range(PURCHASE_TABLE_NAME).columns(PurchaseColumns.INDEX_NUMBER))
        If Not (isect Is Nothing) Then
            correctIndexRange isect
            updateIndexDesc
        End If
        
        '���� �������� ������ � ������� �������, �� ��� ������� ������������ � ��������� �� ���������
        Set isect = Application.Intersect(Target, .Range(PURCHASE_TABLE_NAME))
'        If Not (isect Is Nothing) Then: resetFormulasInPurchaseTable
        If Not (isect Is Nothing) Then
            For i = isect.Rows(1).Cells(1).column To isect.Rows(1).Cells(isect.columns.Count).column
                resetFormulasInPurchaseTable i
            Next i
        End If
    End With
CleanExit:
    Set Target = Nothing
    Set isect = Nothing
    Set priceRange = Nothing
    changeUpdatingState True
    Exit Sub

CutHandler:
    MsgBox "��������� ����� �� ���� ����� ���������. ����������� �����������, ������� � ��������" & _
           vbCrLf & vbCrLf & "Error " & Err.number & ": " & Err.Description
    Resume CleanExit

ErrorHandler:
    MsgBox "Error " & Err.number & ": " & Err.Description
    Resume CleanExit
    
ErrorHandler2:
    MsgBox "Error " & Err.number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "�� ������ ���� � ��������, ������� ������� ��� ������� � ������� �������. �������� �� ���������."
    Resume CleanExit

End Sub


Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim priceRange As Range
    Dim isect As Range
    Dim cell As Range
    Dim tempPN As Variant
    Dim tempName As Variant
    Dim tempComment As String
    
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler2
    With Application.Worksheets(PURCHASE_SHEET_NAME).Range(PURCHASE_TABLE_NAME)
        Set priceRange = Union(.columns(PurchaseColumns.PRICE_GPL), _
                               .columns(PurchaseColumns.PRICE_PURCHASE))
        On Error GoTo ErrorHandler
        
        Set isect = Application.Intersect(Target, priceRange)
        If Not (isect Is Nothing) Then
            priceRange.ClearComments
            For Each cell In isect
                tempPN = .Cells(cell.Row - .Cells(1).Row + 1, PurchaseColumns.PN).Value2
                tempName = .Cells(cell.Row - .Cells(1).Row + 1, PurchaseColumns.NAME_AND_DESCRIPTION).Value2
                tempComment = vbNullString
                If Not IsEmpty(tempPN) Then: tempComment = CStr(tempPN) & vbCrLf
                If Not IsEmpty(tempName) Then: tempComment = tempComment & CStr(tempName)
                If tempComment <> vbNullString Then: cell.AddComment tempComment
            Next cell
        End If
    End With
CleanExit:
    Set Target = Nothing
    Set isect = Nothing
    Set cell = Nothing
    Set priceRange = Nothing
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.number & ": " & Err.Description
    Resume CleanExit
    
ErrorHandler2:
    MsgBox "Error " & Err.number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "�� ������ ���� � ��������, ������� ������� ��� ������� � ������� �������. �������� �� ���������."
    Resume CleanExit
End Sub
