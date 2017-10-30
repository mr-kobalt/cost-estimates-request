Attribute VB_Name = "consts"
Public Const MAXLONG = (2 ^ 31) - 1
Public Const MINLONG = (2 ^ 31) * (-1)
Public Const MAXSINGLE = (2 ^ 15) - 1

' Имена листов, используемых в макросах
Public Const SPEC_SHEET_NAME = "Внутренняя спецификация" ' название листа с расчётом продажи
Public Const SALES_SHEET_NAME = "КП" ' название листа на котором будет формироваться КП
Public Const SERVICE_SHEET_NAME = "Служебный" ' название листа на котором хранятся служебные данные
Public Const AGREEMENT_SHEET_NAME = "Лист согласования" ' название листа согласования условий договора

' Имена "умных" таблиц
Public Const PURCHASE_TABLE_NAME = "Расчёт" ' название таблицы с расчётом закупочных цен
Public Const DELIVERY_TABLE_NAME = "Доставка" ' название таблицы с расчётом доставки
Public Const SHEETS_TABLE_NAME = "Листы"

' Именованные диапазоны
Public Const CURRENCIES_ARRAY_NAME = "валюта" ' список валют
Public Const CURRENCIES_HEADER_ARRAY_NAME = "валюта_кп" ' названия валют для включения в шапку КП
Public Const VAT_ARRAY_NAME = "НДС" ' список НДС
Public Const PROFIT_TYPE_ARRAY_NAME = "прибыль" ' название массива с типами прибыли
Public Const CALC_SOURCE_ARRAY_NAME = "источник"
Public Const UNITS_ARRAY_NAME = "ед_изм"
Public Const CALC_CURRENCIES_ARRAY_NAME = "расчёт_курса" ' название матрицы перерасчёта курса
Public Const CALC_VAT_ARRAY_NAME = "расчёт_НДС" ' название матрицы перерасчёта НДС
Public Const TENDER_ARRAY_NAME = "тендер_область"
Public Const ASSURANCE_ARRAY_NAME = "обеспечение_область"
Public Const MANAGERS_NAMES_ARRAY_NAME = "менеджеры"
Public Const MANAGERS_TITLES_ARRAY_NAME = "менеджеры_должность"
Public Const TERMS_OF_PAYMENT_ARRAY_NAME = "условия_оплаты"
Public Const TERMS_OF_SERVICE_ARRAY_NAME = "условия_выполнения"

' Именованные ячейки
Public Const SALES_CURRENCY_CELL_NAME = "валюта_продажи" ' название ячейки с валютой продажи
Public Const INCLUDE_VAT_CELL_NAME = "включить_НДС" ' включить ли НДС в расчёт
Public Const INCLUDE_DELIVERY_CELL_NAME = "включить_транспортные_расходы" ' добавить ли в расчёт доставку
Public Const CURRENT_RATE_DATE_CELL_NAME = "дата_текущего_курса" ' название ячейки с датой, на которую рассчитывается курс
Public Const TOTAL_COST_CELL_NAME = "себестоимость" ' себестоимость товара без доставка с учётом SALES_CURRENCY_CELL_NAME и INCLUDE_VAT_CELL_NAME
Public Const TOTAL_GPL_CELL_NAME = "GPL" ' сумма GPL товара без доставка с учётом SALES_CURRENCY_CELL_NAME и INCLUDE_VAT_CELL_NAME
Public Const DELIVERY_COST_CELL_NAME = "стоимость_доставки" ' транспортные расходы
Public Const REVENUE_CELL_NAME = "выручка"
Public Const VAT_AMOUNT_CELL_NAME = "размер_НДС"
Public Const VAT_AMOUNT_PURCHASE_CELL_NAME = "размер_НДС_закупки"
Public Const TENDER_CELL_NAME = "тендер"
Public Const ASSURANCE_CELL_NAME = "обеспечение"
Public Const USD_RATE_CELL_NAME = "курс_USD_ЦБ" ' текущий курс USD по ЦБ
Public Const EUR_RATE_CELL_NAME = "курс_EUR_ЦБ" ' текущий курс EUR по ЦБ
Public Const CALC_USD_RATE_CELL_NAME = "курс_USD_расчётный" ' расчётный курс USD
Public Const CALC_EUR_RATE_CELL_NAME = "курс_EUR_расчётный" ' расчётный курс EUR
Public Const CUSTOMER_CELL_NAME = "контрагент"
Public Const PM_CELL_NAME = "проектный_менеджер"

' Имена форм и их групп
Public Const CONTROL_GROUP_NAME = "Панель управления"
Public Const BOARD_SHAPE_NAME = "Доска"

Public Const CHECKBOXES_SUBGROUP_NAME = "Колонки таблицы КП"
Public Const CHECKBOXES_FRAME_NAME = "Окно группы: колонки таблицы КП"
Public Const SALESCOLUMNS_INDEX_NUMBER_SHAPE_NAME = "№"
Public Const SALESCOLUMNS_MANUFACTURER_SHAPE_NAME = "Производитель"
Public Const SALESCOLUMNS_PN_SHAPE_NAME = "p/n"
Public Const SALESCOLUMNS_NAME_AND_DESCRIPTION_SHAPE_NAME = "Наименование"
Public Const SALESCOLUMNS_QTY_SHAPE_NAME = "Кол-во"
Public Const SALESCOLUMNS_UNIT_SHAPE_NAME = "Ед. изм."
Public Const SALESCOLUMNS_PRICE_SHAPE_NAME = "Цена"
Public Const SALESCOLUMNS_TOTAL_SHAPE_NAME = "Сумма"
Public Const SALESCOLUMNS_VAT_SHAPE_NAME = "НДС"
Public Const SALESCOLUMNS_DELIVERY_TIME_SHAPE_NAME = "Срок доставки"

Public Const DROPDOWN_SHAPE_NAME = "Выбор расчёта"
Public Const CALC_BUTTON_SHAPE_NAME = "Кнопка _рассчитать_"
Public Const EXPORT_LABEL_SHAPE_NAME = "Надпись: экспортировать"
Public Const EXPORT_WORD_BUTTON_SHAPE_NAME = "Кнопка: в word"
Public Const EXPORT_EXCEL_BUTTON_SHAPE_NAME = "Кнопка: в excel"
Public Const EXPORT_1C_BUTTON_SHAPE_NAME = "Кнопка: в 1С"

' Прочее
Public Const PRICE_ROUNDING_UP_TO_QTY = 2 ' знаков после запятой при округлении цен
Public Const INDEX_RANK_QTY = 3 ' максимальное количество разрядов в номерах строк КП (см. correctNumberColumn)
Public Const YES = "да"
Public Const NO = "нет"

' Ссылки на внешние источники
Public Const CBR_XML_URL = "http://www.cbr.ru/scripts/XML_daily_eng.asp" ' XML с актуальными курсами с сайта ЦБ РФ

' XML запросы в нотации xPath
Public Const CURRENT_RATE_DATE_XPATH = "//ValCurs/@Date" ' значение даты из CBR_XML_URL
Public Const USD_RATE_XPATH = "//ValCurs/Valute[@ID='R01235']/Value" ' курс доллара из CBR_XML_URL
Public Const EUR_RATE_XPATH = "//ValCurs/Valute[@ID='R01239']/Value" ' курс евро из CBR_XML_URL

' Здесь должны были декларироваться используемые форматы цен в виде строковых констант,
' но VBA не имеет возможности работать с юникодом в IDE, а функцию ChrW(), как и другие
' процедуры,нельзя использовать при задании значения констант, поэтому эти строки определены
' в качестве переменных в каждой процедуре/функции, которая в них нуждается. Значения этих
' переменных представлены ниже:
'Public Const formatRUR = "# ##0,00\ [$" & ChrW(8381) & "-ru-RU]_-;[Красный]-# ##0,00\ [$" & ChrW(8381) & "-ru-RU]_-;""-""??\ [$" & ChrW(8381) & "-ru-RU]_-"
'Public Const formatEUR = "# ##0,00\ [$€-x-euro1]_-;[Красный]-# ##0,00\ [$€-x-euro1]_-;""-""??\ [$€-x-euro1]_-"
'Public Const formatUSD = "# ##0,00 $_-;[Красный]-# ##0,00 $_-;""-""?? $_-"
Public Const DEFAULT_FONT = "Century Gothic"
Public Const ALTERNATIVE_FONT = "Century"
Public Const DATE_FIELD_FORMAT = "\@ ""d MMMM yyyy 'г.'"""
Public Const COMPANY_COLOR = 11762456 ' RGB(24, 123, 179)

' Константы для определения местоположения таблицы КП:
Public Const ROW_OFFSET = 0 ' Сдвиг по вертикали относительно ячейки R1C1
Public Const COLUMN_OFFSET = 0 ' Сдвиг по горизонтали относительно ячейки R1C1

' Нумерация колонок с расчётом департамента закупок
Public Enum PurchaseColumns
    [_FIRST] = 0
    INDEX_NUMBER = 1 ' порядковый номер/индекс
    MANUFACTURER = 2 ' производитель
    PN = 3 ' артикул/продуктовый номер/партномер
    NAME_AND_DESCRIPTION = 4 ' наименование/описание
    qty = 5 ' количество
    Unit = 6 ' единица измерения
    PRICE_SALES = 7
    TOTAL_SALES = 8
    VAT_AMOUNT = 9
    VAT_SALES = 10
    PROFIT_TYPE = 11
    PROFIT_SOURCE = 12
    PROFIT_PERCENT = 13
    MARGIN_AMOUNT = 14
    GPL_CURRENCY = 15 ' валюта прайс-листа
    PRICE_GPL = 16 ' цена прайс-листа
    TOTAL_GPL = 17 ' сумма прайс-листа
    VAT_GPL = 18
    TOTAL_GPL_RECALCULATED = 19 ' сумма прайс-листа
    discount = 20 ' скидка, вычисляемая из суммы прайс-листа и суммы входа в валюте расчёта
    PURCHASE_CURRENCY = 21 ' валюта прайс-листа
    PRICE_PURCHASE = 22 ' цена закупки
    TOTAL_PURCHASE = 23 ' сумма закупки
    VAT_PURCHASE = 24 ' НДС закупки
    TOTAL_PURCHASE_RECALCULATED = 25 ' сумма закупки пересчитанная в валюту продажи
    DELIVERY_TIME = 26 ' срок доставки
    SUPPLIER = 27 ' поставщик
    USER_COMMENTS = 28 ' комментарии
    UNIT_WEIGHT = 29 ' вес штуки
    TOTAL_WEIGHT = 30 ' вес суммарный
    UNIT_VOLUME = 31 ' объём штуки
    TOTAL_VOLUME = 32 ' объём суммарный
    INDEX_DESC = 33
    VAT_PURCHASE_AMOUNT = 34
    [_LAST]
End Enum

' Индексы колонок расчёта КП
Public Enum SalesColumns
    [_FIRST] = 0
    INDEX_NUMBER = 1 ' порядковый номер/индекс
    MANUFACTURER = 2 ' производитель
    PN = 3 ' артикул/продуктовый номер/партномер
    NAME_AND_DESCRIPTION = 4 ' наименование/описание
    qty = 5 ' количество
    Unit = 6 ' единица измерения
    Price = 7 ' цена
    total = 8 ' сумма
    vat = 9 ' НДС
    DELIVERY_TIME = 10 ' срок доставки
    [_MIDDLE] = 99
'    BLANK = 100 ' пустой столбец
'    Row = 101 ' номер соответствующей строки в таблице расчёта закупки (служебный)
'    PROFIT_TYPE = 102
'    CALC_SOURCE = 103
'    PROFIT = 104 ' маржа в процентах
    [_LAST]
End Enum

' Индексы колонок условий оплаты
Public Enum TermsOfPaymentColumns
    [_FIRST] = 0
    typen = 1
    PART = 3
    TIMEAMOUNT = 5
    TIMETYPE = 6
    TIMEDIMENSION = 7
    FROM = 8
    [_LAST]
End Enum

' Индексы колонок условий поставки/выполнения работ
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

' Различный текст
Public Const TEXTS_SUBTOTAL = "ПОДЫТОГ"
Public Const TEXTS_TOTAL = "ИТОГО"
Public Const TEXTS_NOT_SUBJECT_VAT = "НДС не облагается в соответствии с пп.26 ч.2 ст.149 НК РФ"
Public Const TEXTS_SUBJECT_VAT = "в т.ч. НДС 18%"
Public Const TEXTS_NOTICE_MARGIN = "маржа не может быть больше 100%"
Public Const TEXTS_NOTICE_MARKUP = "наценка не может быть меньше -100%"
Public Const TEXTS_MOTTO = "IT-ИНТЕГРАТОР" & vbCrLf & "С ПОЛНЫМ ПРИВОДОМ"
Public Const TEXTS_ADDRESS = "117587, г. Москва" & vbCrLf & "Варшавское шоссе, д. 125Ж, корп. 6" & vbCrLf & "sales@4by4.ru, +7 (499) 753-23-44"
Public Const TEXTS_4X4_SHORT = "ООО ""4х4 УК"""
Public Const TEXTS_4X4_LONG = "Общество с ограниченной ответственностью ""4х4 управляющая компания"""
Public Const TEXTS_FROM = "От: "
Public Const TEXTS_REFERENCE = "Исх. №_________" & vbCrLf & "от "
Public Const TEXTS_WHOM = "Кому: "
Public Const TEXTS_SALES_OFFER = "КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ"
Public Const TEXTS_PITCH = "готово передать Товар и (или) выполнить Работы согласно спецификации:"
Public Const TEXTS_TERMS_OF_PAYMENT = "Условия оплаты:"
Public Const TEXTS_TERMS_OF_SERVICE = "Условия поставки/выполнения работ:"
Public Const TEXTS_DELIVERY_INCLUDED = "стоимость доставки включена в стоимость оборудования"
Public Const TEXTS_DELIVERY_NOT_INCLUDED = "стоимость оборудования не включает стоимость доставки и разгрузочных работ"
Public Const TEXTS_WORK_TITLE = "_______________________"
Public Const TEXTS_SIGN = "____________________"
Public Const TEXTS_LOCUS_SIGILI = "М.П."
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
