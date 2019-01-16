Attribute VB_Name = "test"
Option Explicit
Option Base 1

Sub test()
    With Union(Range("Расчёт").columns(21), Range("Расчёт").columns(23))
        For i = 1 To .FormatConditions.Count
            .FormatConditions(1).Delete
        Next i
        
        With .FormatConditions _
            .Add(xlExpression, , "=" & Range("Расчёт").columns(20).Cells(1).Address(False, False) & "=""RUR""")
            .NumberFormat = formatRUR
            .StopIfTrue = False
        End With
        
        With .FormatConditions _
            .Add(xlExpression, , "=" & Range("Расчёт").columns(20).Cells(1).Address(False, False) & "=""EUR""")
            .NumberFormat = formatEUR
            .StopIfTrue = False
        End With
        
        With .FormatConditions _
            .Add(xlExpression, , "=" & Range("Расчёт").columns(20).Cells(1).Address(False, False) & "=""USD""")
            .NumberFormat = formatUSD
            .StopIfTrue = False
        End With
    End With
End Sub

Sub test2()
    changeUpdatingState True
End Sub

Sub test6()
    changeUpdatingState False
End Sub

Private Sub test1()
    Dim i As Long, j As Long
    Dim arr1(4, 2) As Variant
    Dim arr2 As Collection
    Dim arr3 As Variant
    
    Set arr2 = New Collection
    
    arr1(1, 1) = 4:  arr1(1, 2) = 1
    arr1(2, 1) = 2:  arr1(2, 2) = 1
    arr1(3, 1) = 4:  arr1(3, 2) = 1
    arr1(4, 1) = -1:  arr1(4, 2) = 1
    
    MsgBox minValueInColumn(arr1, 1, arr2)
    MsgBox maxValueInColumn(arr1, 1, arr2)
End Sub
Sub test3()
    Dim oXML As Object
    Set oXML = CreateObject("MSXML2.DOMDocument")
    Dim xml As String

    oXML.async = False
    xml = "http://www.cbr.ru/scripts/XML_daily_eng.asp"
    oXML.Load xml
    
    Dim oSeqNodes As Object
    Dim oSeqNode As Object
    'Set iseqnodes = CreateObject("IXMLDOMNode")
    Set oSeqNode = oXML.SelectSingleNode("//ValCurs/Valute[@ID='R01235']/Value")
    If oXML.readystate > 2 Then: MsgBox typename(oSeqNode.Text): MsgBox oXML.readystate
'    If oSeqNodes.Length = 0 Then
'       'show some message
'    Else
'        For Each oSeqNode In oSeqNodes
'             MsgBox oSeqNode.SelectSingleNode("name").Text
'        Next
'    End If
End Sub

Sub test4()
    MsgBox Format(Application.Range(REVENUE_CELL_NAME).Value2, "# ##0.00")
    
End Sub

Sub test5()
    MsgBox convertPriceToText(-1.11, Application.Range(CALC_CURRENCY_CELL_NAME))
    
End Sub
