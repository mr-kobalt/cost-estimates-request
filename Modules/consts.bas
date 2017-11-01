Attribute VB_Name = "consts"
Public Const MAXLONG = (2 ^ 31) - 1
Public Const MINLONG = (2 ^ 31) * (-1)
Public Const MAXSINGLE = (2 ^ 15) - 1

' ����� ������, ������������ � ��������
Public Const PURCHASE_SHEET_NAME = "������ �������" ' �������� ����� � �������� ���������� ���
Public Const SALES_SHEET_NAME = "������ �������" ' �������� ����� �� ������� ����� ������������� ��
Public Const SERVICE_SHEET_NAME = "���������" ' �������� ����� �� ������� �������� ��������� ������
Public Const AGREEMENT_SHEET_NAME = "���� ������������" ' �������� ����� ������������ ������� ��������

' ����� "�����" ������
Public Const PURCHASE_TABLE_NAME = "������" ' �������� ������� � �������� ���������� ���
Public Const DELIVERY_TABLE_NAME = "��������" ' �������� ������� � �������� ��������

' ����������� ���������
Public Const CURRENCIES_ARRAY_NAME = "������" ' ������ �����
Public Const CURRENCIES_HEADER_ARRAY_NAME = "������_��" ' �������� ����� ��� ��������� � ����� ��
Public Const VAT_ARRAY_NAME = "���" ' ������ ���
Public Const PROFIT_TYPE_ARRAY_NAME = "�������" ' �������� ������� � ������ �������
Public Const CALC_SOURCE_ARRAY_NAME = "��������"
Public Const CALC_CURRENCIES_ARRAY_NAME = "������_�����" ' �������� ������� ����������� �����
Public Const CALC_VAT_ARRAY_NAME = "������_���" ' �������� ������� ����������� ���
Public Const TENDER_ARRAY_NAME = "������_�������"
Public Const ASSURANCE_ARRAY_NAME = "�����������_�������"
Public Const SERVICE_COLUMNS_ARRAY_NAME = "����_�������"
Public Const MANAGERS_NAMES_ARRAY_NAME = "���������"
Public Const MANAGERS_TITLES_ARRAY_NAME = "���������_���������"
Public Const TERMS_OF_PAYMENT_ARRAY_NAME = "�������_������"
Public Const TERMS_OF_SERVICE_ARRAY_NAME = "�������_����������"

' ����������� ������
Public Const CALC_CURRENCY_CELL_NAME = "������_�������" ' �������� ������ � ������� �������
Public Const INCLUDE_VAT_CELL_NAME = "��������_���" ' �������� �� ��� � ������
Public Const INCLUDE_DELIVERY_CELL_NAME = "��������_������������_�������" ' �������� �� � ������ ��������
Public Const CURRENT_RATE_DATE_CELL_NAME = "����_��������_�����" ' �������� ������ � �����, �� ������� �������������� ����
Public Const TOTAL_COST_CELL_NAME = "�������������" ' ������������� ������ ��� �������� � ������ CALC_CURRENCY_CELL_NAME � INCLUDE_VAT_CELL_NAME
Public Const TOTAL_GPL_CELL_NAME = "GPL" ' ����� GPL ������ ��� �������� � ������ CALC_CURRENCY_CELL_NAME � INCLUDE_VAT_CELL_NAME
Public Const DELIVERY_COST_CELL_NAME = "���������_��������" ' ������������ �������
Public Const REVENUE_CELL_NAME = "�������"
Public Const VAT_AMOUNT_CELL_NAME = "������_���"
Public Const TENDER_CELL_NAME = "������"
Public Const ASSURANCE_CELL_NAME = "�����������"
Public Const USD_RATE_CELL_NAME = "����_USD_��" ' ������� ���� USD �� ��
Public Const EUR_RATE_CELL_NAME = "����_EUR_��" ' ������� ���� EUR �� ��
Public Const CALC_USD_RATE_CELL_NAME = "����_USD_���������" ' ��������� ���� USD
Public Const CALC_EUR_RATE_CELL_NAME = "����_EUR_���������" ' ��������� ���� EUR
Public Const CUSTOMER_CELL_NAME = "����������"
Public Const PM_CELL_NAME = "���������_��������"
Public Const DELIVERY_COST_FRACTION_CELL_NAME = "����_��������"

' ����� ���� � �� �����
Public Const BOARD_SHAPE_NAME = "�����"

Public Const CHECKBOXES_GROUP_NAME = "������� ������� ��"
Public Const CHECKBOXES_SUBGROUP_NAME = "���� ������: ������� ������� ��"
Public Const SALESCOLUMNS_INDEX_NUMBER_SHAPE_NAME = "�"
Public Const SALESCOLUMNS_MANUFACTURER_SHAPE_NAME = "�������������"
Public Const SALESCOLUMNS_PN_SHAPE_NAME = "p/n"
Public Const SALESCOLUMNS_NAME_AND_DESCRIPTION_SHAPE_NAME = "������������"
Public Const SALESCOLUMNS_QTY_SHAPE_NAME = "���-��"
Public Const SALESCOLUMNS_UNIT_SHAPE_NAME = "��. ���."
Public Const SALESCOLUMNS_PRICE_SHAPE_NAME = "����"
Public Const SALESCOLUMNS_TOTAL_SHAPE_NAME = "�����"
Public Const SALESCOLUMNS_VAT_SHAPE_NAME = "���"
Public Const SALESCOLUMNS_DELIVERY_TIME_SHAPE_NAME = "���� ��������"

Public Const CALC_PARAMS_GROUP_NAME = "��������� �������"
Public Const CALC_PARAMS_SUBGROUP_NAME = "���� ������: ��������� �������"
Public Const CURRENCY_LABEL_SHAPE_NAME = "����� _������_"
Public Const VAT_LABEL_SHAPE_NAME = "����� _���_"
Public Const CURRENCY_SHAPE_NAME = "������ _������_"
Public Const VAT_SHAPE_NAME = "������ _���_"
Public Const DELIVERY_SHAPE_NAME = "�������� ��������"

Public Const PROFIT_GROUP_NAME = "������ �������"
Public Const PROFIT_SUBGROUP_NAME = "���� ������: ������ �������"
Public Const CALC_TYPE_SUBGROUP_NAME = "���� ������: ������ �������"
Public Const MARKUP_SHAPE_NAME = "�������"
Public Const MARGIN_SHAPE_NAME = "�����"
Public Const CALC_SOURCE_SUBGROUP_NAME = "���� ������: ��������"
Public Const GPL_SHAPE_NAME = "�� GPL"
Public Const NET_PRICE_SHAPE_NAME = "�� �����"
Public Const CALC_LABEL_SHAPE_NAME = "������� _����������_"
Public Const CALC_BUTTON_SHAPE_NAME = "������ _����������_"

Public Const EXPORT_GROUP_NAME = "�������"
Public Const EXPORT_SUBGROUP_NAME = "���� ������: ��������������"
Public Const EXPORT_LABEL_SHAPE_NAME = "������� _��������������_"
Public Const EXPORT_WORD_BUTTON_SHAPE_NAME = "������: � word"
Public Const EXPORT_EXCEL_BUTTON_SHAPE_NAME = "������: � excel"

' ������
Public Const PRICE_ROUNDING_UP_TO_QTY = 2 ' ������ ����� ������� ��� ���������� ���
Public Const INDEX_RANK_QTY = 3 ' ������������ ���������� �������� � ������� ����� �� (��. correctNumberColumn)
Public Const YES = "��"
Public Const NO = "���"

' ������ �� ������� ���������
Public Const CBR_XML_URL = "http://www.cbr.ru/scripts/XML_daily_eng.asp" ' XML � ����������� ������� � ����� �� ��

' XML ������� � ������� xPath
Public Const CURRENT_RATE_DATE_XPATH = "//ValCurs/@Date" ' �������� ���� �� CBR_XML_URL
Public Const USD_RATE_XPATH = "//ValCurs/Valute[@ID='R01235']/Value" ' ���� ������� �� CBR_XML_URL
Public Const EUR_RATE_XPATH = "//ValCurs/Valute[@ID='R01239']/Value" ' ���� ���� �� CBR_XML_URL

' ����� ������ ���� ��������������� ������������ ������� ��� � ���� ��������� ��������,
' �� VBA �� ����� ����������� �������� � �������� � IDE, � ������� ChrW(), ��� � ������
' ���������,������ ������������ ��� ������� �������� ��������, ������� ��� ������ ����������
' � �������� ���������� � ������ ���������/�������, ������� � ��� ���������. �������� ����
' ���������� ������������ ����:
'Public Const formatRUR = "# ##0,00\ [$" & ChrW(8381) & "-ru-RU]_-;[�������]-# ##0,00\ [$" & ChrW(8381) & "-ru-RU]_-;""-""??\ [$" & ChrW(8381) & "-ru-RU]_-"
'Public Const formatEUR = "# ##0,00\ [$�-x-euro1]_-;[�������]-# ##0,00\ [$�-x-euro1]_-;""-""??\ [$�-x-euro1]_-"
'Public Const formatUSD = "# ##0,00 $_-;[�������]-# ##0,00 $_-;""-""?? $_-"
Public Const DEFAULT_FONT = "Century Gothic"
Public Const ALTERNATIVE_FONT = "Century"
Public Const DATE_FIELD_FORMAT = "\@ ""d MMMM yyyy '�.'"""
Public Const COMPANY_COLOR = 11762456 ' RGB(24, 123, 179)

' ��������� ��� ����������� �������������� ������� ��:
Public Const ROW_OFFSET = 19 ' ����� �� ��������� ������������ ������ R1C1
Public Const COLUMN_OFFSET = 1 ' ����� �� ����������� ������������ ������ R1C1

' ��������� ������� � �������� ������������ �������
Public Enum PurchaseColumns
    [_FIRST] = 0
    INDEX_NUMBER = 1 ' ���������� �����/������
    MANUFACTURER = 2 ' �������������
    PN = 3 ' �������/����������� �����/���������
    NAME_AND_DESCRIPTION = 4 ' ������������/��������
    qty = 5 ' ����������
    Unit = 6 ' ������� ���������
    PRICE_GPL_RECALCULATED = 7 ' ���� �����-����� ����� ��������� � ������_�������
    TOTAL_GPL_RECALCULATED = 8 ' ����� �����-����� ����� ��������� � ������_�������
    DISCOUNT = 9 ' ������, ����������� �� ����� �����-����� � ����� ����� � ������ �������
    PRICE_PURCHASE_RECALCULATED = 10 ' ���� ������� ����� ��������� � ������_�������
    TOTAL_PURCHASE_RECALCULATED = 11 ' ����� ������� ����� ��������� � ������_�������
    VAT_PURCHASE = 12 ' ��� �������
    DELIVERY_TIME = 13 ' ���� ��������
    SUPPLIER = 14 ' ���������
    USER_COMMENTS = 15 ' �����������
    UNIT_WEIGHT = 16 ' ��� �����
    TOTAL_WEIGHT = 17 ' ��� ���������
    UNIT_VOLUME = 18 ' ����� �����
    TOTAL_VOLUME = 19 ' ����� ���������
    GPL_CURRENCY = 20 ' ������ �����-�����
    PRICE_GPL = 21 ' ���� �����-�����
    PURCHASE_CURRENCY = 22 ' ������ �����-�����
    PRICE_PURCHASE = 23 ' ���� �������
    vat = 24 ' Value Added Tax - ���
    TOTAL_GPL = 25 ' ����� �����-�����
    TOTAL_PURCHASE = 26 ' ����� �������
    INDEX_DESC = 27 ' �������� ���� �������
    VAT_PURCHASE_AMOUNT = 28 ' ������ ��� �������
    [_LAST]
End Enum

' ������� ������� ������� ��
Public Enum SalesColumns
    [_FIRST] = 0
    INDEX_NUMBER = 1 ' ���������� �����/������
    MANUFACTURER = 2 ' �������������
    PN = 3 ' �������/����������� �����/���������
    NAME_AND_DESCRIPTION = 4 ' ������������/��������
    qty = 5 ' ����������
    Unit = 6 ' ������� ���������
    Price = 7 ' ����
    total = 8 ' �����
    vat = 9 ' ���
    DELIVERY_TIME = 10 ' ���� ��������
    [_MIDDLE] = 99
    BLANK = 100 ' ������ �������
    Row = 101 ' ����� ��������������� ������ � ������� ������� ������� (���������)
    PROFIT_TYPE = 102
    CALC_SOURCE = 103
    PROFIT = 104 ' ����� � ���������
    [_LAST]
End Enum

' ������� ������� ������� ������
Public Enum TermsOfPaymentColumns
    [_FIRST] = 0
    typename = 1
    PART = 3
    TIMEAMOUNT = 5
    TIMETYPE = 6
    TIMEDIMENSION = 7
    FROM = 8
    [_LAST]
End Enum

' ������� ������� ������� ��������/���������� �����
Public Enum TermsOfServiceColumns
    [_FIRST] = 0
    CITY = 1
    PART = 4
    TIMEAMOUNT = 6
    TIMETYPE = 7
    TIMEDIMENSION = 8
    FROM = 9
    [_LAST]
End Enum

' ��������� �����
Public Const TEXTS_SUBTOTAL = "�������"
Public Const TEXTS_TOTAL = "�����"
Public Const TEXTS_NOT_SUBJECT_VAT = "��� �� ���������� � ������������ � ��.26 �.2 ��.149 �� ��"
Public Const TEXTS_SUBJECT_VAT = "� �.�. ��� 18%"
Public Const TEXTS_NOTICE_MARGIN = "����� �� ����� ���� ������ 100%"
Public Const TEXTS_NOTICE_MARKUP = "������� �� ����� ���� ������ -100%"
Public Const TEXTS_MOTTO = "IT-����������" & vbCrLf & "� ������ ��������"
Public Const TEXTS_ADDRESS = "117587, �. ������" & vbCrLf & "���������� �����, �. 125�, ����. 6" & vbCrLf & "sales@4by4.ru, +7 (499) 753-23-44"
Public Const TEXTS_4X4_SHORT = "��� ""4�4 ��"""
Public Const TEXTS_4X4_LONG = "�������� � ������������ ���������������� ""4�4 ����������� ��������"""
Public Const TEXTS_FROM = "��: "
Public Const TEXTS_REFERENCE = "���. �_________" & vbCrLf & "�� "
Public Const TEXTS_WHOM = "����: "
Public Const TEXTS_SALES_OFFER = "������������ �����������"
Public Const TEXTS_PITCH = "������ �������� ����� � (���) ��������� ������ �������� ������������:"
Public Const TEXTS_TERMS_OF_PAYMENT = "������� ������:"
Public Const TEXTS_TERMS_OF_SERVICE = "������� ��������/���������� �����:"
Public Const TEXTS_DELIVERY_INCLUDED = "��������� �������� �������� � ��������� ������������"
Public Const TEXTS_DELIVERY_NOT_INCLUDED = "��������� ������������ �� �������� ��������� �������� � ������������ �����"
Public Const TEXTS_WORK_TITLE = "_______________________"
Public Const TEXTS_SIGN = "____________________"
Public Const TEXTS_LOCUS_SIGILI = "�.�."
Public Const TEXTS_SUBTITLE = "subtitle"
Public Const TEXTS_ASSEMBLY = "assembly"

' MS Word constants
Public Const wdWindowStateMaximize = 1
Public Const wdAutoFitWindow = 2
Public Const wdFormatOriginalFormatting = 16
Public Const wdCollapseEnd = 0
Public Const wdAlignVerticalCenter = 1
Public Const wdAlignParagraphLeft = 0
Public Const wdAlignParagraphCenter = 1
Public Const wdAlignParagraphRight = 2
Public Const wdAlignParagraphJustify = 3
Public Const wdAlignParagraphDistribute = 4
Public Const wdCharacter = 1
Public Const wdFieldCreateDate = 21
Public Const wdAdjustFirstColumn = 2
Public Const wdLineStyleSingle = 1
Public Const wdBorderBottom = -3
