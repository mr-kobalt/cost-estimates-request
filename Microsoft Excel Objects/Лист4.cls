VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "����4"
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
    Dim isect As Range

    On Error GoTo ErrorHandler
    Application.EnableEvents = False

    With Application.Worksheets(AGREEMENT_SHEET_NAME)
        ' ���� ��������� ��� ������ ������ � ������� � ������, �� ��������������
        ' �������� ������������� �� � �����
        Set isect = Application.Intersect(Target, .Range(TENDER_CELL_NAME))
        If Not (isect Is Nothing) Then
            If isect.Value2 = "��" Then
                .Range(TENDER_ARRAY_NAME).Resize(.Range(TENDER_ARRAY_NAME).Rows.Count - 3).offset(2).EntireRow.Hidden = False
            Else
                .Range(TENDER_ARRAY_NAME).Resize(.Range(TENDER_ARRAY_NAME).Rows.Count - 3).offset(2).EntireRow.Hidden = True
            End If
        End If
        
        Set isect = Application.Intersect(Target, .Range(ASSURANCE_CELL_NAME))
        If Not (isect Is Nothing) Then
            If isect.Value2 = "��" Then
                .Range(ASSURANCE_ARRAY_NAME).Resize(.Range(ASSURANCE_ARRAY_NAME).Rows.Count - 3).offset(2).EntireRow.Hidden = False
            Else
                .Range(ASSURANCE_ARRAY_NAME).Resize(.Range(ASSURANCE_ARRAY_NAME).Rows.Count - 3).offset(2).EntireRow.Hidden = True
            End If
        End If
    End With
    
CleanExit:
    Set isect = Nothing
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.number & ": " & Err.Description
    Resume CleanExit
End Sub

