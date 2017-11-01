VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProfitDialog 
   Caption         =   "Размер "
   ClientHeight    =   1190
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   3120
   OleObjectBlob   =   "ProfitDialog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProfitDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()
' нажатие на кнопку отмена закрывает диалоговое окно
    Unload Me
End Sub

Private Sub CancelButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
' если фокус на кнопке отмена, то нажаните Enter и Esс закрывает диалоговое окно
    Select Case KeyCode
    Case vbKeyReturn
        CancelButton_Click
    Case vbKeyEscape
        CancelButton_Click
    End Select
End Sub

Private Sub OkButton_Click()
' При клике на кнопку OK записываем введённые данные
    With Application.Worksheets(SALES_SHEET_NAME).Shapes(CALC_BUTTON_SHAPE_NAME).OLEFormat.Object
        If IsNumeric(TextBox1.Value) Then
            .Caption = CStr(CDbl(TextBox1.Value))
        Else
            .Caption = "0"
        End If
    End With
    
    createSalesOffer
    Unload Me
End Sub

Private Sub OkButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
' если фокус на кнопке ОК, то Enter записает введённые данные,
' а Esс закрывает диалоговое окно
    Select Case KeyCode
    Case vbKeyReturn
        OkButton_Click
    Case vbKeyEscape
        CancelButton_Click
    End Select
End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
' если фокус на текстовом поле, то Enter записает введённые данные,
' а Esс закрывает диалоговое окно
    Select Case KeyCode
    Case vbKeyReturn
        OkButton_Click
    Case vbKeyEscape
        CancelButton_Click
    End Select
    
    If Shift > 2 Then KeyCode = 0
End Sub

Private Sub TextBox1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
' если фокус на текстовом поле, то Enter записает введённые данные,
' а Esс закрывает диалоговое окно
' если зажата клавиша Shift, Ctrl, Alt или любая их комбинация, то игнорируем введённый символ
    Dim regexp As Object
    Dim minus As String
    Dim comma As String
    Dim temp As String
    Dim numerator As String
    Dim fraction As String
    
    Dim commaPos As Integer
    
    Set regexp = CreateObject("vbscript.regexp")
    minus = vbNullString
    comma = vbNullString
    
    temp = CStr(TextBox1.Value)

    ' Пропускаем ячеейки, данные в которых невозможно преобразовать
    'On Error Resume Next
    With regexp
        .Global = True
        ' шаблон соответствует всем нечисловым символам, кроме последней точки
        ' или запятой, которая считается разделителем целой и дробной части
        .Pattern = "[^\d\.\,]+|[^\d]+(?=.*[\.\,].*$)"

            
        If Len(temp) > 0 Then
            If Mid(temp, 1, 1) = "-" Then: minus = "-"
        End If
        
        temp = Replace(.Replace(temp, vbNullString), ".", ",")
        commaPos = InStr(1, temp, ",")
        
        Select Case True
        Case temp = vbNullString
            temp = minus & temp
        Case commaPos = 0
            temp = minus & CStr(CDbl(temp))
        Case commaPos = 1 And Len(temp) = 1
            temp = minus & "0,"
        Case commaPos = 1 And Len(temp) > 1
            fraction = Mid(temp, 2)
            temp = minus & "0," & fraction
        Case commaPos > 1 And commaPos = Len(temp)
            numerator = Mid(temp, 1, Len(temp) - 1)
            temp = minus & CStr(CDbl(numerator)) & ","
        Case Else
            numerator = Mid(temp, 1, InStr(1, temp, ",") - 1)
            fraction = Mid(temp, InStr(1, temp, ",") + 1)
            temp = minus & CStr(CDbl(numerator)) & "," & fraction
        End Select
    End With
    
    If IsNumeric(temp) Then
        If (CDbl(temp) > 100 And _
                            Application.Worksheets(SALES_SHEET_NAME).Shapes(MARGIN_SHAPE_NAME).OLEFormat.Object.Value = xlOn) Then
            temp = "100"
        ElseIf (CDbl(temp) < -100 And _
                            Application.Worksheets(SALES_SHEET_NAME).Shapes(MARKUP_SHAPE_NAME).OLEFormat.Object.Value = xlOn) Then
            temp = "-100"
        End If
    End If
    
    TextBox1.Value = temp
End Sub

Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
' фильтруем вводимые в текстовое поле символы - допускается ввод только чисел и десятичной запятой
    Select Case KeyAscii
        Case vbKey0 To vbKey9 ' [0-9]
            If (CDbl(TextBox1.Value & Chr(KeyAscii)) > 100 And _
                        Application.Worksheets(SALES_SHEET_NAME).Shapes(MARGIN_SHAPE_NAME).OLEFormat.Object.Value = xlOn) Or _
                        (CDbl(TextBox1.Value & Chr(KeyAscii)) < -100 And _
                        Application.Worksheets(SALES_SHEET_NAME).Shapes(MARKUP_SHAPE_NAME).OLEFormat.Object.Value = xlOn) Then
                KeyAscii = 0
            ElseIf TextBox1.Value = "0" Then
                TextBox1.Value = vbNullString
            End If
            
        Case 44 ' ","
            ' если в поле уже есть одна десятичная запятая, то игнорируем ввод других
            If "," Like "[" & TextBox1.Value & "]" Then: KeyAscii = 0
        Case 46 ' "."
            ' введённая точка интерпретируется как десятичный разделитель и преобразуется в запятую
            If "," Like "[" & TextBox1.Value & "]" Then
                KeyAscii = 0
            Else
                KeyAscii = 44
            End If
        Case 45 ' "-"
            KeyAscii = 0
            If Not ("-" Like "[" & TextBox1.Value & "]") Then: TextBox1.Value = "-" & TextBox1.Value
        Case 43 ' "+"
            KeyAscii = 0
            If "-" Like "[" & TextBox1.Value & "]" Then: TextBox1.Value = Mid(TextBox1.Value, 2)
        Case Else
            KeyAscii = 0
    End Select

End Sub

Private Sub UserForm_Initialize()
' Инициализируем форму ручного ввода процента маржи или прибыли
    Dim shape As shape
    Dim str As String
    
'    For Each shape In Application.Worksheets(SALES_SHEET_NAME).shapes(PROFIT_GROUP_NAME).GroupItems
'        If shape.FormControlType = xlOptionButton Then
'            If shape.OLEFormat.Object.Value = xlOn Then
'                str = Mid(shape.AlternativeText, 1, Len(shape.AlternativeText) - 1)
'
'                ProfitDialog.Caption = ProfitDialog.Caption & str & "и"
'                Label1.Caption = Label1.Caption & str & "и"
'                Exit For
'            End If
'        End If
'    Next shape
    
    If Application.Worksheets(SALES_SHEET_NAME).Shapes(MARGIN_SHAPE_NAME).OLEFormat.Object.Value = xlOn Then
        Set shape = Application.Worksheets(SALES_SHEET_NAME).Shapes(MARGIN_SHAPE_NAME)
        Label2.Caption = TEXTS_NOTICE_MARGIN
    ElseIf Application.Worksheets(SALES_SHEET_NAME).Shapes(MARKUP_SHAPE_NAME).OLEFormat.Object.Value = xlOn Then
        Set shape = Application.Worksheets(SALES_SHEET_NAME).Shapes(MARKUP_SHAPE_NAME)
        Label2.Caption = TEXTS_NOTICE_MARKUP
    End If
    str = Mid(shape.AlternativeText, 1, Len(shape.AlternativeText) - 1)
    
    ProfitDialog.Caption = ProfitDialog.Caption & str & "и"
    'Label1.Caption = Label1.Caption & str & "и"
    Label1.Caption = shape.OLEFormat.Object.Caption
    
    With Application.Worksheets(SALES_SHEET_NAME).Shapes(CALC_BUTTON_SHAPE_NAME).OLEFormat.Object
        If IsNumeric(.Caption) Then
            If CDbl(.Caption) <> 0 Then
                TextBox1.Value = CStr(.Caption)
            Else
                TextBox1.Value = vbNullString
            End If
        Else
            TextBox1.Value = vbNullString
        End If
    End With
    
    Set shape = Nothing
End Sub
