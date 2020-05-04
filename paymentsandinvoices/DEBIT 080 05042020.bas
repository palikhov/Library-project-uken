Sub Auto_Open()

'==============================================================================================================
'==============================================================================================================
' Module    : Statement_Processing
' Author    : Антон Палихов aka Palant | http://palikhov.wordpress.com | e-mail: a.v.palikhov@hotmail.com
' Client    : NPC Ukrenergo
' Version   :  0.80 04.05.2020                                                                               =
'==============================================================================================================
'==============================================================================================================
Dim x As String
   strPath = "D:\ExcelBackup" 'директория для резервной копии
   On Error Resume Next
   x = GetAttr(strPath) And 0
   If Err = 0 Then ' если данный путь существует, то сохраняем в него открываемую книгу
       strDate = Format(Now, "dd/mm/yy hh-mm")
       FileNameXls = strPath & "\" & ActiveWorkbook.Name & " " & strDate & ".xlsb"
       ActiveWorkbook.SaveCopyAs Filename:=FileNameXls
   Else 'если путь не существует или же он недоступен, то выводим соответствующее сообщение
       MsgBox "Создание резервной копии невозможно! Папка " & strPath & " недоступна или не существует!", vbCritical
   End If
End Sub


Public Sub Clearing()
'==============================================================================================================
'==============================================================================================================
' Module    : Clearing
' Author    : Антон Палихов aka Palant | http://palikhov.wordpress.com | e-mail: a.v.palikhov@hotmail.com
' Client    : NPC Ukrenergo
' Version   :  0.80 04.05.2020                                                                               =
'==============================================================================================================
'==============================================================================================================
Dim cRowsStatementCount As Double

Dim cRowsManualCount As Double

Dim cRowsControlCount As Double

Dim oRangeStatement As Range

Dim oRangeControl As Range

Dim oRangeManual As Range

Dim i, j As Integer
'Очистка данных на листе выписки
cRowsStatementCount = ThisWorkbook.Sheets("statement").Cells(Rows.Count, 28).End(xlUp).Row

Set oRangeStatement = ThisWorkbook.Sheets("statement").Range(Cells(2, 28).Address, Cells(cRowsStatementCount + 1, 38).Address)

oRangeStatement.ClearContents


'Очистка данных на листе ручной обработки

cRowsManualCount = ThisWorkbook.Sheets("manual_processing").Cells(Rows.Count, 2).End(xlUp).Row

Set oRangeManual = ThisWorkbook.Sheets("manual_processing").Range(Cells(2, 1).Address, Cells(cRowsManualCount + 1, 23).Address)

oRangeManual.ClearContents

For i = 2 To cRowsManualCount
    
    For j = 1 To 22
    
    ThisWorkbook.Sheets("manual_processing").Cells(i, j).Interior.Color = xlNone
    
    Next j
    
Next i

'Очистка данных на листе обработки контрольных значений
cRowsControlCount = ThisWorkbook.Sheets("control_processing").Cells(Rows.Count, 2).End(xlUp).Row

Set oRangeControl = ThisWorkbook.Sheets("control_processing").Range(Cells(2, 1).Address, Cells(cRowsControlCount + 1, 23).Address)

oRangeControl.ClearContents

'Очистка завершена. Выводим сообщение

MsgBox "Проміжні дані на аркушах виписки, ручної обробки та обробки контродьних значень очищені"

End Sub

Sub Clear_registry_additional_data()
'==============================================================================================================
'==============================================================================================================
' Module    : Clearing_REGISTRY
' Author    : Антон Палихов aka Palant | http://palikhov.wordpress.com | e-mail: a.v.palikhov@hotmail.com
' Client    : NPC Ukrenergo
' Version   :  0.80 04.05.2020                                                                               =
'==============================================================================================================
'==============================================================================================================
Worksheets("registry").Range("Af2:bz200000").ClearContents
End Sub



Public Sub Clearing1()
'==============================================================================================================
'==============================================================================================================
' Module    : Clearing
' Author    : Антон Палихов aka Palant | http://palikhov.wordpress.com | e-mail: a.v.palikhov@hotmail.com
' Client    : NPC Ukrenergo
' Version   :  0.80 04.05.2020                                                                               =
'==============================================================================================================
'==============================================================================================================
Dim cRowsStatementCount As Double

Dim cRowsManualCount As Double

Dim cRowsControlCount As Double

Dim oRangeStatement As Range

Dim oRangeControl As Range

Dim oRangeManual As Range

Dim i, j As Integer
'Очистка данных на листе выписки
cRowsStatementCount = ThisWorkbook.Sheets("statement").Cells(Rows.Count, 28).End(xlUp).Row

Set oRangeStatement = ThisWorkbook.Sheets("statement").Range(Cells(2, 28).Address, Cells(cRowsStatementCount + 1, 38).Address)

oRangeStatement.ClearContents


'Очистка данных на листе ручной обработки

cRowsManualCount = ThisWorkbook.Sheets("manual_processing").Cells(Rows.Count, 2).End(xlUp).Row

Set oRangeManual = ThisWorkbook.Sheets("manual_processing").Range(Cells(2, 1).Address, Cells(cRowsManualCount + 1, 23).Address)

oRangeManual.ClearContents

For i = 2 To cRowsManualCount
    
    For j = 1 To 22
    
    ThisWorkbook.Sheets("manual_processing").Cells(i, j).Interior.Color = xlNone
    
    Next j
    
Next i

'Очистка данных на листе обработки контрольных значений
cRowsControlCount = ThisWorkbook.Sheets("control_processing").Cells(Rows.Count, 2).End(xlUp).Row

Set oRangeControl = ThisWorkbook.Sheets("control_processing").Range(Cells(2, 1).Address, Cells(cRowsControlCount + 1, 23).Address)

oRangeControl.ClearContents

'Очистка завершена. Выводим сообщение

MsgBox "Проміжні дані на аркушах виписки, ручної обробки та обробки контродьних значень очищені"

End Sub

Sub Clear_registry_additional_data()
'
' Макрос4 Макрос
'

Worksheets("registry").Range("Af2:bz200000").ClearContents
End Sub



Public Sub Clearing2()
'==============================================================================================================
'==============================================================================================================
' Module    : Clearing
' Author    : Антон Палихов aka Palant | http://palikhov.wordpress.com | e-mail: a.v.palikhov@hotmail.com
' Client    : NPC Ukrenergo
' Version   :  0.80 04.05.2020                                                                               =
'==============================================================================================================
'==============================================================================================================
Dim cRowsStatementCount As Double

Dim cRowsManualCount As Double

Dim cRowsControlCount As Double

Dim oRangeStatement As Range

Dim oRangeControl As Range

Dim oRangeManual As Range

Dim i, j As Integer
'Очистка данных на листе выписки
cRowsStatementCount = ThisWorkbook.Sheets("statement").Cells(Rows.Count, 28).End(xlUp).Row

Set oRangeStatement = ThisWorkbook.Sheets("statement").Range(Cells(2, 28).Address, Cells(cRowsStatementCount + 1, 38).Address)

oRangeStatement.ClearContents


'Очистка данных на листе ручной обработки

cRowsManualCount = ThisWorkbook.Sheets("manual_processing").Cells(Rows.Count, 2).End(xlUp).Row

Set oRangeManual = ThisWorkbook.Sheets("manual_processing").Range(Cells(2, 1).Address, Cells(cRowsManualCount + 1, 23).Address)

oRangeManual.ClearContents

For i = 2 To cRowsManualCount
    
    For j = 1 To 22
    
    ThisWorkbook.Sheets("manual_processing").Cells(i, j).Interior.Color = xlNone
    
    Next j
    
Next i

'Очистка данных на листе обработки контрольных значений
cRowsControlCount = ThisWorkbook.Sheets("control_processing").Cells(Rows.Count, 2).End(xlUp).Row

Set oRangeControl = ThisWorkbook.Sheets("control_processing").Range(Cells(2, 1).Address, Cells(cRowsControlCount + 1, 23).Address)

oRangeControl.ClearContents

'Очистка завершена. Выводим сообщение

MsgBox "Проміжні дані на аркушах виписки, ручної обробки та обробки контродьних значень очищені"

End Sub

Sub Clear_registry_additional_data()
'
' Макрос4 Макрос
'

Worksheets("registry").Range("Af2:bz200000").ClearContents
End Sub



'Процедура вставки формул
Public Sub InsertFormula()

'==============================================================================================================
'==============================================================================================================
' Module    : InsertFormula
' Author    : Антон Палихов aka Palant | http://palikhov.wordpress.com | e-mail: a.v.palikhov@hotmail.com
' Client    : NPC Ukrenergo
' Version   :  0.80 04.05.2020                                                                               =
'==============================================================================================================
'==============================================================================================================


Dim tRowsCount As Double
Dim tRowsCount2 As Double
Dim i As Double
tRowsCount = ThisWorkbook.Sheets("REGISTRY").Cells(Rows.Count, 1).End(xlUp).Row
tRowsCount2 = ThisWorkbook.Sheets("REGISTRY").Cells(Rows.Count, 3).End(xlUp).Row

With Sheets("registry")
For i = tRowsCount + 1 To tRowsCount2

.Cells(i, 1).FormulaR1C1Local = "=VLOOKUP(RC[2];ДОГОВОРИ!R4C2:R2600C4;2;FALSE)"

.Cells(i, 2).FormulaR1C1Local = "=VLOOKUP(RC[1];ДОГОВОРИ!R4C2:R2600C4;3;FALSE)"

.Cells(i, 17).FormulaR1C1Local = "=ROUND(RC[-1]*1,2;2)"
.Cells(i, 32).FormulaR1C1Local = "=RC[-29]&RC[-14]"

Next i
End With
End Sub

Public Sub Coloration()

'==============================================================================================================
'==============================================================================================================
' Module    : Coloration
' Author    : Антон Палихов aka Palant | http://palikhov.wordpress.com | e-mail: a.v.palikhov@hotmail.com
' Client    : NPC Ukrenergo
' Version   :  0.80 04.05.2020                                                                               =
'==============================================================================================================
'==============================================================================================================
Dim i As Double
Dim j As Double
Dim k As Double
Dim f As Double
Dim tRowsCount As Double
tRowsCount = ThisWorkbook.Sheets("REGISTRY").Cells(Rows.Count, 3).End(xlUp).Row
MsgBox tRowsCount
j = 0
For i = 1 To tRowsCount

If ThisWorkbook.Sheets("REGISTRY").Cells(i + 1, 40).Value = 0 And ThisWorkbook.Sheets("REGISTRY").Cells(i + 1, 33).Value <> 0 Then ThisWorkbook.Sheets("Registry").Cells(i + 1, 17).Interior.Color = 52377

'If ThisWorkbook.Sheets("Registry").Cells(i + 1, 13).Value > 0 And ThisWorkbook.Sheets("Registry").Cells(i + 1, 34).Value <> O Then ThisWorkbook.Sheets("Registry").Cells(i + 1, 30).Value = "ПОМИЛКА. ВЖЕ НАЯВНА ДАТА ОПЛАТИ"
If ThisWorkbook.Sheets("REGISTRY").Cells(i + 1, 40).Value = 0 And ThisWorkbook.Sheets("REGISTRY").Cells(i + 1, 33).Value <> 0 And ThisWorkbook.Sheets("Registry").Cells(i + 1, 14).Value = 0 Then ThisWorkbook.Sheets("Registry").Cells(i + 1, 14).Value = ThisWorkbook.Sheets("Registry").Cells(i + 1, 35).Value
If ThisWorkbook.Sheets("REGISTRY").Cells(i + 1, 40).Value = 0 And ThisWorkbook.Sheets("REGISTRY").Cells(i + 1, 33).Value <> 0 Then j = j + 1
'MsgBox ThisWorkbook.Sheets("REGISTRY").Cells(i + 1, 39).Value
'MsgBox ThisWorkbook.Sheets("REGISTRY").Cells(i + 1, 32).Value
'MsgBox "i" & i
'MsgBox "j" & j


Next i

MsgBox "Співпадають суми по " & j & " інвойсів"

tRowsCount = ThisWorkbook.Sheets("manual_processing").Cells(Rows.Count, 3).End(xlUp).Row
For i = 1 To tRowsCount

If Application.IsError(Application.Find("Розподіл", ThisWorkbook.Sheets("manual_processing").Cells(i + 1, 20).Value, 1)) = True Then k = 0 Else k = WorksheetFunction.Find("Розподіл", ThisWorkbook.Sheets("manual_processing").Cells(i + 1, 20).Value, 1)

If k > 0 Then _

    For j = 1 To 22
    
    ThisWorkbook.Sheets("manual_processing").Cells(i + 1, j).Interior.Color = 8388736
    
    Next j

End If
If Application.IsError(Application.Find("борг", ThisWorkbook.Sheets("manual_processing").Cells(i + 1, 20).Value, 1)) = True Then f = 0 Else f = WorksheetFunction.Find("борг", ThisWorkbook.Sheets("manual_processing").Cells(i + 1, 20).Value, 1)

If f > 0 Then _

    For j = 1 To 22
    
    ThisWorkbook.Sheets("manual_processing").Cells(i + 1, j).Interior.Color = 16763904
    
    Next j

End If
Next i


End Sub

Public Sub Calendar_data_inserting()
'==============================================================================================================
'==============================================================================================================
' Module    : Calendar_data_inserting
' Author    : Антон Палихов aka Palant | http://palikhov.wordpress.com | e-mail: a.v.palikhov@hotmail.com
' Client    : NPC Ukrenergo
' Version   :  0.80 04.05.2020                                                                               =
'==============================================================================================================
'==============================================================================================================
Dim i As Double ' Счетчик

Dim cRowsRegistryCount As Double 'Количество строк с данными на листе реестра

Dim dYear As String

Dim dMonth As Integer

Dim tMonth As String

Dim dDay As Integer

Dim dTenday As String

cRowsRegistryCount = ThisWorkbook.Sheets("Registry").Cells(Rows.Count, 2).End(xlUp).Row

MsgBox cRowsRegistryCount

For i = 2 To cRowsRegistryCount

'Простановка года

If ThisWorkbook.Sheets("Registry").Cells(i, 6).Value <> "" Then GoTo Label1 'Если в столбце года стоят данные - переходим к следующей строке

If ThisWorkbook.Sheets("Registry").Cells(i, 11).Value = "" Then GoTo Label1 'Если в столбце даты нет данных - переходим к следующей строке

If ThisWorkbook.Sheets("Registry").Cells(i, 11).Value = 0 Then GoTo Label1 'Если в столбце даты нет данных - переходим к следующей строке

dYear = Year(ThisWorkbook.Sheets("Registry").Cells(i, 11).Value) 'Присвоить временной переменной значения года из столбца даты
 
ThisWorkbook.Sheets("Registry").Cells(i, 6).Value = dYear 'Записать в столбец года значение временной переменной

If ThisWorkbook.Sheets("Registry").Cells(i, 7).Value <> "" Then GoTo Label1 'Если в столбце месяца стоят данные - переходим к следующей строке

dMonth = Month(ThisWorkbook.Sheets("Registry").Cells(i, 11).Value) 'Присвоить временной переменной значения месяца из столбца даты
'На основании числа месяца определим текстовую строку его названия и запишем во временную перменную
If dMonth = 1 Then tMonth = "січень"
If dMonth = 2 Then tMonth = "лютий"
If dMonth = 3 Then tMonth = "березень"
If dMonth = 4 Then tMonth = "квітень"
If dMonth = 5 Then tMonth = "травень"
If dMonth = 6 Then tMonth = "червень"
If dMonth = 7 Then tMonth = "липень"
If dMonth = 8 Then tMonth = "серпень"
If dMonth = 9 Then tMonth = "вересень"
If dMonth = 10 Then tMonth = "жовтень"
If dMonth = 11 Then tMonth = "листопад"
If dMonth = 12 Then tMonth = "грудень"

ThisWorkbook.Sheets("Registry").Cells(i, 7).Value = tMonth 'Запишем в столбец месяца значение временной переменной

If ThisWorkbook.Sheets("Registry").Cells(i, 8).Value <> "" Then GoTo Label1 'Если в столбце декады стоят данные - переходим к следующей строке

dDay = Day(ThisWorkbook.Sheets("Registry").Cells(i, 11).Value) 'По дате определяем день месяца
'По дню месяца определяем декаду и записываем во временную переменную
If dDay > 0 And dDay < 11 Then dTenday = 1
If dDay > 10 And dDay < 21 Then dTenday = 2
If dDay > 20 And dDay < 32 Then dTenday = 3

ThisWorkbook.Sheets("Registry").Cells(i, 8).Value = dTenday 'Записываем значение временной переменной в столбец декады

tMonth = ""
dTenday = ""
dYear = ""

Label1: Next i
End Sub




Public Sub Invoices_add_to_Registry()

'==============================================================================================================
'==============================================================================================================
' Module    : Statement_Processing
' Author    : Антон Палихов aka Palant | http://palikhov.wordpress.com | e-mail: a.v.palikhov@hotmail.com
' Client    : NPC Ukrenergo
' Version   :  0.80 04.05.2020                                                                               =
'==============================================================================================================
'==============================================================================================================

Dim i As Integer

Dim oRowsCount As Integer

Dim tRowsCount As Double

Dim targetWS As Sheets

Dim dataWS As Sheets

Dim Credit_counter As Integer

oRowsCount = ThisWorkbook.Sheets("invoices").UsedRange.Rows.Count

'Set dataWS = ThisWorkbook.Sheets("invoices")

'Set targetWS = ThisWorkbook.Sheets("REGISTRY")

tRowsCount = ThisWorkbook.Sheets("REGISTRY").Cells(Rows.Count, 1).End(xlUp).Row

MsgBox oRowsCount
MsgBox tRowsCount
Credit_counter = 0

For i = 1 To oRowsCount
If ThisWorkbook.Sheets("invoices").Cells(i + 1, 11) = "DEDIT" Then _
  
    
    ThisWorkbook.Sheets("REGISTRY").Cells(tRowsCount + 1 + Credit_counter, 3).Value = ThisWorkbook.Sheets("invoices").Cells(i + 1, 2).Value ' PARTICIPANT_EIC
   
   ' MsgBox ThisWorkbook.Sheets("invoices").Cells(i + 1, 2).Value
   ' MsgBox ThisWorkbook.Sheets("REGISTRY").Cells(tRowsCount + i, 3).Value
     ThisWorkbook.Sheets("REGISTRY").Cells(tRowsCount + 1 + Credit_counter, 4).Value = ThisWorkbook.Sheets("invoices").Cells(i + 1, 3).Value 'SERVICE_TYPE 4
     ThisWorkbook.Sheets("REGISTRY").Cells(tRowsCount + 1 + Credit_counter, 8).Value = ThisWorkbook.Sheets("invoices").Cells(i + 1, 4).Value 'INVOICE_NUMBER 8
    'Set targetWS.Cells(tRowsCount + i, 9).Value = dataWS.Cells(i + 1, 2).Value 'CONTRACT_NUMBER 9
     ThisWorkbook.Sheets("REGISTRY").Cells(tRowsCount + 1 + Credit_counter, 10).Value = ThisWorkbook.Sheets("invoices").Cells(i + 1, 7).Value 'DATA_FROM 10
     ThisWorkbook.Sheets("REGISTRY").Cells(tRowsCount + 1 + Credit_counter, 11).Value = ThisWorkbook.Sheets("invoices").Cells(i + 1, 8).Value 'DATA_TO 11
     ThisWorkbook.Sheets("REGISTRY").Cells(tRowsCount + 1 + Credit_counter, 12).Value = ThisWorkbook.Sheets("invoices").Cells(i + 1, 11).Value 'TRANSACTION_TYPE 12
     ThisWorkbook.Sheets("REGISTRY").Cells(tRowsCount + 1 + Credit_counter, 15).Value = ThisWorkbook.Sheets("invoices").Cells(i + 1, 12).Value 'AMOUNT 15
    'CONTRACT_ID.
    
    Credit_counter = Credit_counter + 1
    Application.StatusBar = Credit_counter
End If

Next i

MsgBox Credit_counter

End Sub





Public Sub statement_clearing()

Dim cStatementRowsCount As Double
Dim oStatement As Range
Dim statws As Worksheet
Dim ostat As Range
'==============================================================================================================
'cStatementRowsCount = ThisWorkbook.Sheets("statement").Cells(Rows.Count, 2).End(xlUp).Row
'Set statws = ThisWorkbook.Sheets("statement")
'Set ostat = statws.Range(Cells(2, 1).Address, Cells(cStatementRowsCount, 79).Address)

    Worksheets("statement").Range("A2:bz200000").ClearContents
 'ostat.ClearContents
'==============================================================================================================
MsgBox "Проміжні дані на аркуші виписки очищені"
End Sub

Public Function RegExpExtract(Text As String, Pattern As String, Optional Item As Integer = 1) As String
    On Error GoTo ErrHandl
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = Pattern
    regex.Global = True
    If regex.Test(Text) Then
        Set matches = regex.Execute(Text)
        RegExpExtract = matches.Item(Item - 1)
        Exit Function
    End If
ErrHandl:
    RegExpExtract = CVErr(xlErrValue)
End Function


Public Sub statement_IMPORTING()
'==============================================================================================================
'==============================================================================================================
' Module    : Statement_Processing
' Author    : Антон Палихов aka Palant | http://palikhov.wordpress.com | e-mail: a.v.palikhov@hotmail.com
' Client    : NPC Ukrenergo
' Version   :  0.80 04.05.2020                                                                               =
'==============================================================================================================
'==============================================================================================================
'********************************************************************************************
'                                    ПЕРЕМЕННЫЕ ДЛЯ КОМАНД
'********************************************************************************************
Dim thatWB As Workbook, thisWB As Workbook
Dim thisWS As Worksheet, thatWS As Worksheet
Dim zOpenFileName As String
Dim inputData As String
Dim cRowsCount As Integer
Dim oRange As Range
Dim oSupport As Range
Dim iRowsCount As Integer
Dim WSName As String
Dim iRowsCount2 As Integer
'КАКОЙ ФАЙЛ НЕОБХОДИМО ОТКРЫТЬ
'inputData = InputBox("Enter name of file")
'ОТКРЫТИЕ ФАЙЛА
zOpenFileName = Application.GetOpenFilename
'ПРОВЕРКА НА ОШИБКИ
If zOpenFileName = "" Then Exit Sub
Application.ScreenUpdating = False
Set thisWB = ThisWorkbook 'ЭТА КНИГА - КНИГА НАЗНАЧЕНИЯ, В КОТОРУЮ НЕОБХОДИМО СКОПИРОВАТЬ ДАННЫЕ
Set thisWS = ThisWorkbook.Sheets("statement") 'ЛИСТ выписки, НА КОТОРЫЙ НЕОБХОДИМО ВСТАВИТЬ ДАННЫЕ ИЗ ФАЙЛА ИМПОРТА ИЗ ММС

cRowsCount = ThisWorkbook.Sheets("statement").Cells(Rows.Count, 2).End(xlUp).Row

Set oRange = thisWS.Range(Cells(2, 1).Address, Cells(cRowsCount + 1, 22).Address)

oRange.ClearContents

Set oSupport = thisWS.Range(Cells(2, 27).Address, Cells(cRowsCount + 1, 78).Address)

oSupport.ClearContents


Set thatWB = Workbooks.Open(zOpenFileName) ' файл источник
'MsgBox thatWB.Name
WSName = Left(thatWB.Name, Len(thatWB.Name) - 4)
'MsgBox WSName

Set thatWS = thatWB.Sheets("Page1") 'источник данных - название листа

Application.CutCopyMode = False

iRowsCount = thatWS.UsedRange.Rows.Count

Set IRange = thatWS.Range(Cells(6, 1), Cells(iRowsCount, 22))

iRowsCount2 = iRowsCount - 4

IRange.Copy
thisWB.Activate
thisWS.Range(Cells(2, 1).Address, Cells(iRowsCount2, 22).Address).PasteSpecial Paste:=xlPasteValues

With thisWS.Range(Cells(2, 1).Address, Cells(iRowsCount2, 22).Address)
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
End With

thatWB.Close
ThisWorkbook.Sheets("Main").Activate

Application.ScreenUpdating = True

MsgBox "Банківська виписка імпортована"



End Sub

Sub Statement_Processing()
    
    On Error GoTo Statement_Processing_Err
    
    '==============================================================================================================
    '==============================================================================================================
    ' Module    : Statement_Processing                                                                            =
    ' Author    : Антон Палихов aka Palant | http://palikhov.wordpress.com | e-mail: a.v.palikhov@hotmail.com     =
    ' Client    : NPC Ukrenergo                                                                                   =
     ' Version   :  0.80 04.05.2020                                                                               =
    '==============================================================================================================
    '==============================================================================================================
    
'********************************************************************************************
'*                                    ПЕРЕМЕННЫЕ ДЛЯ КОМАНД                                 *
'********************************************************************************************
    
    
    
    Dim InvoiceRegExp As New RegExp ' создаем экземпляр RegExp
    Dim DateRegExp As New RegExp ' создаем экземпляр RegExp
    Dim myRegExp As New RegExp ' создаем экземпляр RegExp
    Dim InvoiceMatch, DateMatch As Match ' один из совпавших образцов
    Dim colInvoiceMatches, colDateMatches As MatchCollection ' коллекция этих образцов
    Dim strTest As String ' тестируемая строка
    Dim i, j, k, l, p, q As Integer ' счетчики для циклов
    Dim aInvoice As Integer 'полученный номер инвойса
    Dim sum_statement As Double 'сумма платежа в строке на листе с данными выписки
    Dim Date_dogovor_i As Date 'дата заключения договора
    Dim Date_operation_i As Date 'дата проведения операции
    Dim Date_invoice_i As Date 'дата формирования инвойсаа
    Dim customer_i As String 'Номер контрагента
    Dim date_temp1 As Date
    Dim date_temp2 As Date
    Dim date_temp3 As Date
    Dim invoice_i As Double ' Номер инвойса
    Dim date_payment_i As Date ' Дата платежа
    Dim NumberOfDates As Integer 'Количество распознанных дат в назначении платежа
    Dim Income_counter As Integer 'Количество строк, в которых отправитель платежа Укрэнерго
    Dim Manual_counter As Integer 'Счетчик строк, попавших на лист ручной обработки
    Dim Automated_counter As Integer 'Счетчик строк, обработанных автоматически
    Dim Semi_automated_counter As Integer 'Счетчик строк выписки, обработанный по номеру контрагента и сумме в строке выписке
    Dim Row_Sum As Long
    Dim Statement_Rows As Integer 'Переменная для хранения количества строк платежей на листе с данными банковской выписки
    Dim invoice_match_i As Integer 'номер строки в которой находится искомый инвойс на листе с массивом инвойсов
    Dim matchi As Double 'Замена переменной для определения сопоставления строки платежей и строки инвойсов
    Dim oReg As Range 'Диапазон данн
    Dim payment_type_i As String 'Тип платежу
    Dim temp_invoice_i As Double
    Dim invoice_semimatch_i As Integer 'номер строки в которой находится сумма из листа выписки на листе инвойсов по этому контрагенту
    Dim matching_type As Boolean
    Dim temp_sum As Double ' Временная сумма
    Dim temp_text As String ' Временный текст
    Dim semi_invoice_i As String 'Переменная состоящая из номера контрагента и суммы платежа
    Dim CorrectRows As Integer 'Количество корректно обработанных строк инвойсов
    Dim semi_invoice_match_i As Integer 'номер строки по нестрогому соответствию
    Dim regws As Worksheet
    Dim cRegistryRowsCount As Double
    Dim semi_invoice_match_i2 As Double
    Dim semi_i As Integer
    
'********************************************************************************************
    'Запрос у пользователя о намерении продолжать?
    
   ' Select Case MsgBox("Цю операцію не можна буде відмінити " & _
  '  "Збережемо робочу книгу перед початком операції?", vbYesNoCancel)
  '  Case Is = vbYes
   ' ThisWorkbook.Save
   ' Case Is = vbCancel
  '  Exit Sub
  '  End Select
'Начальные значения переменных
    Manual_counter = 0
    Automated_counter = 0
    Income_counter = 0
    Semi_automated_counter = 0
    CorrectRows = 0
    Set regws = ThisWorkbook.Sheets("REGISTRY")
    Set oReg = regws.Range(Cells(31, 2).Address, Cells(cRowsCount + 1, 78).Address)
'********************************************************************************************
    'Начало обработки выписки
    
    Application.StatusBar = "Запущена автоматична обробка виписок" 'Вывод сообщения в статус-бар
    
    'Определение количества строк в выписке
    
    Statement_Rows = ActiveWorkbook.Sheets("statement").UsedRange.Rows.Count
    'Определение количества строк на  листе реестра
    cRegistryRowsCount = ThisWorkbook.Sheets("REGISTRY").Cells(Rows.Count, 2).End(xlUp).Row

    'уставнока диапазона который необходимо очистить
    Set oReg = regws.Range(Cells(2, 33).Address, Cells(cRegistryRowsCount, 78).Address)
    
    oReg.ClearContents
 '********************************************************************************************


    '********************************************************************************************
    '   Начало цикла по строкам выписки
    '********************************************************************************************
    
    For i = 1 To Statement_Rows
    
            'В статус баре указать процент обработки листа statement (i/Statement_rows)*100 %
            
            Application.StatusBar = "В обробці " & "рядок із" & Statement_Rows 'Вывод сообщения в статус-бар
            
            k = 0  'переменные для определения инвойса
            p = 0  'переменные для определения инвойса
            q = 0   'переменные для определения инвойса
            
            invoice_match_i = 0 'начальное значение сопоставления
            
            invoice_i = 0 'начальное значение инвойса
            
            semi_invoice_match_i2 = 0 'для каждой строки соответствие изначально равно 0

            If ThisWorkbook.Sheets("statement").Cells(i + 1, 2).Value = "№ документу" Then GoTo Label1:     'проверка на ошибку и переход к следующей строке
            If ThisWorkbook.Sheets("statement").Cells(i + 1, 2).Value = "" Then GoTo Label1:                'проверка на ошибку и переход к следующей строке
                     
            sum_statement = WorksheetFunction.Substitute(ThisWorkbook.Sheets("statement").Cells(i + 1, 6).Value, " UAH", "")    'Определение суммы платежа на листе выписки
            
            date_payment_i = ThisWorkbook.Sheets("statement").Cells(i + 1, 3).Value                                             'Определение даты платежа на листе выписки

            customer_i = ThisWorkbook.Sheets("statement").Cells(i + 1, 14).Value                                                'Определение контрагента на листе выписки
        




OutcomeProcessing:             'Определение того что платеж является исходящим и должен быть обработан при обработке реестра дебита
            '********************************************************************************************
            '   Проверка на исходящий платеж и обработка исходящего платежа
            '********************************************************************************************
                If ThisWorkbook.Sheets("statement").Cells(i + 1, 14).Value = "00100227" Or ThisWorkbook.Sheets("statement").Cells(i + 1, 14).Value = "100227" Then ThisWorkbook.Sheets("statement").Cells(i + 1, 28).Value = "Outcome" Else GoTo AutoProcessing 'Если ЕГРПОУ отправителя равен ЕГРПОУ Укрэнерго - значит это исходящий платеж
            
                Income_counter = Income_counter + 1 'Увеличение счетчика исходящих платежей
                
                GoTo Label1 ' Переход к следующей строке
                        
AutoProcessing:
             '============================================================================================
            ' Автоматическая обработка
            '============================================================================================
                'Попытка поиска инвойса
                
                InvoiceRegExp.Global = True ' если Global = True, то поиск ведется по всей строке, если False, то только до первого совпадения

                InvoiceRegExp.IgnoreCase = True ' игнорировать регистр
                
                InvoiceRegExp.Pattern = "\d{13}" ' шаблон для поиска. Инвойс содержит в себе 13 цифр
                
                Set colInvoiceMatches = InvoiceRegExp.Execute(ThisWorkbook.Sheets("statement").Cells(i + 1, 20)) ' создание коллекции с элементами которые соответствую паттерну

                If colInvoiceMatches.Count = 0 Then GoTo Outcome_Semi_Auto_Processing
                
                'Если инвойс не найдено в назначении платежа, попробовать обработать строку с помощью неточного сопоставления, если и это не получится - скопировать строку на лист ручной обработки
                                
                If colInvoiceMatches.Count = 1 Then invoice_i = colInvoiceMatches.Item(0).Value                 'Если в коллекции только один элемент - тогда считаем его инвойсом
                
                If colInvoiceMatches.Count = 1 Then GoTo DATE_DETECTION                                         'Если только один элемент - инвойс определен - переходим к определнию дат в назначении платежа
                                

                
'********************************************************************************************
INVOICE_DETECTION:
                '********************************************************************************************
                '   ПОПЫТКА АНАЛИТИЧЕСКОГО ОПРЕДЕЛЕНИЯ - КАКАЯ ИЗ ДВУХ ПОСЛЕДОВАТЕЛЬНОСТЕЙ 13 ЦИФР ЯВЛЯЕТСЯ НОМЕРОМ ИНВОЙСА?
                '********************************************************************************************

                'ДЛЯ НАЧАЛА ВЫВЕДЕМ В 60 И 61 СТОЛБЕЦ ПОЛУЧЕННЫЕ ЗНАЧЕНИЯ

                ThisWorkbook.Sheets("statement").Cells(i + 1, 60).Value = colInvoiceMatches.Item(0).Value           'ПЕРВАЯ ПОСЛЕДОВАТЕЛЬНОСТЬ

                ThisWorkbook.Sheets("statement").Cells(i + 1, 61).Value = colInvoiceMatches.Item(1).Value           'ВТОРАЯ ПОСЛЕДОВАТЕЛЬНОСТЬ
                
                p = WorksheetFunction.Find(colInvoiceMatches.Item(0).Value, ThisWorkbook.Sheets("statement").Cells(i + 1, 20).Value, 1)     'Определение количества символов слева от первой последовательности
                
                ThisWorkbook.Sheets("statement").Cells(i + 1, 62).Value = p                                                                 'Вывод этого значения в 62 столбец
                
                q = WorksheetFunction.Find(colInvoiceMatches.Item(1).Value, ThisWorkbook.Sheets("statement").Cells(i + 1, 20).Value, 1)     'Определение количества символов слева от второй последовательности
                
                ThisWorkbook.Sheets("statement").Cells(i + 1, 63).Value = q                                                                 'Вывод этого значения в 63 столбец
                
                ThisWorkbook.Sheets("statement").Cells(i + 1, 64).Value = Right(Left(ThisWorkbook.Sheets("statement").Cells(i + 1, 20).Value, p - 1), 1)    'В 64 столбец выводим символ стоящий слева от первой последовательности
                
                ThisWorkbook.Sheets("statement").Cells(i + 1, 65).Value = Right(Left(ThisWorkbook.Sheets("statement").Cells(i + 1, 20).Value, q - 1), 1)    'В 65 столбец выводим символ стоящий слева от второй последовательности
                
                If Right(Left(ThisWorkbook.Sheets("statement").Cells(i + 1, 20).Value, p - 1), 1) = "X" Then invoice_i = colInvoiceMatches.Item(1).Value    'Если окружающий символ первой последовательности равен X то первая последовательность - X идентификатор. Инвойс - вторая последовательность
                If Right(Left(ThisWorkbook.Sheets("statement").Cells(i + 1, 20).Value, p - 1), 1) = "Х" Then invoice_i = colInvoiceMatches.Item(1).Value    'Некоторын пишут кириллическую Х
                
                If Right(Left(ThisWorkbook.Sheets("statement").Cells(i + 1, 20).Value, q - 1), 1) = "X" Then invoice_i = colInvoiceMatches.Item(0).Value    'Если окружающий символ второй последовательности равен X то вторая последовательность - X идентификатор. Инвойс - первая последовательность
                If Right(Left(ThisWorkbook.Sheets("statement").Cells(i + 1, 20).Value, q - 1), 1) = "Х" Then invoice_i = colInvoiceMatches.Item(0).Value    'Некоторын пишут кириллическую Х
                ThisWorkbook.Sheets("statement").Cells(i + 1, 66).Value = invoice_i 'выводим в 66 столбец значение инвойса
                

'********************************************************************************************
DATE_DETECTION:
               
                '********************************************************************************************
                '   ПОИСК ДАТ В НАЗНАЧЕНИИ ПЛАТЕЖА
                '********************************************************************************************

                DateRegExp.Global = True ' если Global = True, то поиск ведется по всей строке, _если False, то только до первого совпадения

                DateRegExp.IgnoreCase = True ' игнорировать регистр

                DateRegExp.Pattern = "(\d{1,2}[\.\-]){2}(\d{4}|\d{2})|(\d{1,2}[\/\-]){2}(\d{4}|\d{2})|(\d{1,2}[\-\-]){2}(\d{4}|\d{2})" ' шаблон для поиска даты
            
                Set colDateMatches = DateRegExp.Execute(ThisWorkbook.Sheets("statement").Cells(i + 1, 20))

                NumberOfDates = colDateMatches.Count
            
                'If NumberOfDates < 2 Then _                    'Если найдено меньше двух дат - указать это в поле примечание

                    
                If NumberOfDates = 2 Then _
                
                        'Если найдено две даты, то это вероятнее всего дата операции и дата инвойса - в примечание идет запись "Оплата ххх числа, дата операции ххх, дата инвойса ххх"
                                        
                        date_temp1 = colDateMatches(0).Value    'Дата операции =  минимальная из двух дат

                        date_temp2 = colDateMatches(1).Value
                        
                        Date_operation_i = WorksheetFunction.Min(date_temp1, date_temp2)
                        
                        Date_invoice_i = WorksheetFunction.Max(date_temp1, date_temp2) 'Дата инвойса = максимальная из двух дат
                        
                End If
                        
                If NumberOfDates = 3 Then _
                
                'Если найдено три даты, то это вероятнее всего дата операции, дата инвойса, дата заключения договора

                    date_temp1 = colDateMatches(0).Value
                    date_temp2 = colDateMatches(1).Value
                    date_temp3 = colDateMatches(2).Value
                    Date_dogovor_i = WorksheetFunction.Min(date_temp1, date_temp2, date_temp3) 'Дата заключения договора= минимальная из трех дат
                    Date_invoice_i = WorksheetFunction.Max(date_temp1, date_temp2, date_temp3) 'Дата инвойса = максимальная из трех дат
                    'Дата операции = средняя дата
                    If colDateMatches(0).Value > Date_dogovor_i And colDateMatches(0).Value < Date_invoice_i Then Date_operation_i = colDateMatches(0).Value
                    If colDateMatches(1).Value > Date_dogovor_i And colDateMatches(1).Value < Date_invoice_i Then Date_operation_i = colDateMatches(1).Value
                    If colDateMatches(2).Value > Date_dogovor_i And colDateMatches(2).Value < Date_invoice_i Then Date_operation_i = colDateMatches(2).Value
                        
                End If
            
            
            '********************************************************************************************
            '   Определение типа платежа
            '********************************************************************************************
            
            '********************************************************************************************
            '   ОПРЕДЕЛЕНИЕ КОНТРАГЕНТА, СУММЫ ПЛАТЕЖА, даты платежа
            '********************************************************************************************

            not_strong_invoice_i = customer_i & sum_statement 'при нестрогом поиске соответствия сопоставлять будем одинаковые суммы у одинаковых контрагентов. Для этого необходимо создать одномерные массивы контрагент+сумма

            ThisWorkbook.Sheets("statement").Cells(i + 1, 36).Value = not_strong_invoice_i 'запишем в 36 столбец значение контрагент + сумма платежа
            
            '********************************************************************************************
            '   ПОИСК АНАЛОГИЧНОГО ИНВОЙСА В МАССИВЕ ИНВОЙСОВ, ОБРАБОТКА ДАННЫХ
            '********************************************************************************************
            
            'В случае если номер инвойса определен, ищем этот номер инвойса в массиве инвойсов
            
            
            invoice_match_i = 0 'Начальное значение сопоставления - 0

            matchi = 0 'и для второй переменной тоже
            
        'ПРОВЕРКА - ЕСЛИ НАйденный инвойс начинается с 0, то и на листе реестра он стоит с 0, а значит форматирование экселя уже съело этот 0 поэтому первый символ необходимо отбросить и перезаписать значение инвойса

            ThisWorkbook.Sheets("statement").Cells(i + 1, 55).Value = Left(invoice_i, 1) 'В 55 СТОЛБЕЦ ВЫВЕДЕМ ПЕРВЫЙ СИМВОЛ ИНВОЙСА
                      
            If Left(invoice_i, 1) = "0" Then invoice_i = Right(invoice_i, 12) * 1
            
            ThisWorkbook.Sheets("statement").Cells(i + 1, 50).Value = invoice_i 'ВЫВОДИМ определенное ЗНАЧЕНИЕ ИНВОЙСА В СТОЛБЕЦ 50

            ThisWorkbook.Sheets("statement").Cells(i + 1, 51).Value = invoice_match_i 'ВЫВОДИМ В СТОЛБЕЦ 51 ЗНАЧЕНИЕ ПЕРЕМЕННОЙ invoice_match_i до начала поиска сопоставления

            ThisWorkbook.Sheets("statement").Cells(i + 1, 52).Value = Application.IsError(Application.Match(invoice_i, ThisWorkbook.Sheets("REGISTRY").Range("H1:H50000").Value, 0)) 'ВЫВОДИМ В СТОЛБЕЦ 52 РЕЗУЛЬТАТ ПРОВЕРКИ ПОИСКА ИНВОЙСА НА ЛИСТЕ РЕЕСТРА. В СЛУЧАЕ УСПЕХА В 52 СТОЛБЕЦ ДОЛЖНО ЗАПИСАТЬСЯ FALSE
            
            If Application.IsError(Application.Match(invoice_i, ThisWorkbook.Sheets("REGISTRY").Range("I1:I100000").Value, 0)) = True Then matchi = 0 Else matchi = WorksheetFunction.Match(invoice_i, ThisWorkbook.Sheets("REGISTRY").Range("I1:I100000").Value, 0) 'ЕСЛИ ПРОВЕРКА ПОИСКА ВЫДАЕТ ОШИБКУ, ТО СТРОГО СОПОСТАВЛЕНИЯ ДОБИТЬСЯ НЕ УДАЛОСЬ. ПРИСВАИВАЕМ ПЕРЕМЕННОЙ ЗНАЧЕНИЕ 0. В ПРОТИВНОМ СЛУЧАЕ ПРИСВАИВАЕМ ПЕРЕМЕННОЙ ЗНАЧЕНИЕ СОПОСТАВЛЛЕНОЙ СТРОКИ
            
            ThisWorkbook.Sheets("statement").Cells(i + 1, 56).Value = matchi
           
                        ThisWorkbook.Sheets("statement").Cells(i + 1, 53).Value = WorksheetFunction.Match(invoice_i, ThisWorkbook.Sheets("REGISTRY").Range("I1:I100000").Value, 0) 'ВЫВОДИМ В СТОЛБЕЦ 53 ЗНАЧЕНИЕ НАЙДЕННОЙ СТРОКИ
                    
            ThisWorkbook.Sheets("statement").Cells(i + 1, 54).Value = matchi 'ВЫВОДИМ В СТОЛБЕЦ 54 ЗНАЧЕНИЕ ПЕРЕМЕННОЙ
                      
            If matchi > 0 Then GoTo Auto Else GoTo Outcome_Semi_Auto_Processing 'ЕСЛИ ПЕРЕМЕННАЯ БОЛЬШЕ 0 ЗНАЧИТ ЕСТЬ СОПОСТАВЛЕНИЕ - ПЕРЕХОДИМ НА ЗАКЛАДКУ АВТОМАТИЧЕСКОГО СОПОСТАВЛЕНИЯ. В ПРОТИВНОМ СЛУЧАЕ ПЕРЕХОДИМ НА ЗАКЛАДКУ ПОЛУАВТОМАТИЧЕСКОГО СОПОСТАВЛЕНИЯ ЧТОБЫ ПОПЫТАТЬСЯ СОПОСТАВИТЬ ПО НЕСТРОГОМУ ПРИЗНАКУ
            
            'В случае если найден -> идем по пути строгогосоответствия. переход  по закладкам не нужен
            
'********************************************************************************************
Auto:
            
            ThisWorkbook.Sheets("statement").Cells(i + 1, 31).Value = invoice_i 'Записываем в столбец AE на листе statement значение номера инвойса
            
            temp_sum = 0
            
            temp_sum = ThisWorkbook.Sheets("Registry").Cells(matchi, 33).Value 'Временно записываем в переменную значение из ячейки

            If temp_sum <> 0 Then ThisWorkbook.Sheets("Registry").Cells(matchi, 37).Value = "Повтор інвойсу"
            
            ThisWorkbook.Sheets("Registry").Cells(matchi, 36).Value = sum_statement
            
            ThisWorkbook.Sheets("Registry").Cells(matchi, 38).Value = not_strong_invoice_i
            
            ThisWorkbook.Sheets("Registry").Cells(matchi, 33).Value = temp_sum + sum_statement 'Используем временное значение для определения новой суммы
            
            'Дописать в примечание реквизиты платежа - дату операции, дату платежа и сумму платежа

            temp_text = ThisWorkbook.Sheets("Registry").Cells(matchi, 34).Value 'Временно записываем в переменную текст из ячейки
            
            ThisWorkbook.Sheets("Registry").Cells(matchi, 35).Value = date_payment_i
            
            ThisWorkbook.Sheets("Registry").Cells(matchi, 34).Value = temp_text & "; " & date_payment_i & " - " & sum_statement 'В ячейке прописываем текст состоящий из временной переменной и добавлением из текущей строки выписки
            
            ThisWorkbook.Sheets("statement").Cells(i + 1, 28).Value = "Auto"
            
            ThisWorkbook.Sheets("Registry").Cells(matchi, 39).Value = "Auto"
            
            ThisWorkbook.Sheets("Registry").Cells(matchi, 40).FormulaR1C1Local = "=RC[-7]-RC[-23]"
            
            ThisWorkbook.Sheets("statement").Cells(i + 1, 29).Value = date_payment_i
            
            ThisWorkbook.Sheets("statement").Cells(i + 1, 30).Value = sum_statement
            
            ThisWorkbook.Sheets("statement").Cells(i + 1, 32).Value = matchi 'Запись в колонку AF значения соответствия строке на листе Реестр
            Automated_counter = Automated_counter + 1 'Увеличиваем счетчик обработанных строк в автоматическом режиме
            k = k + 1
            
            ThisWorkbook.Sheets("statement").Cells(i + 1, 38).Value = k
            
            GoTo Label1 'Переходим к следующей строке
            'В случае если не найден
            
        
            
Outcome_Semi_Auto_Processing:
            
            
            
            semi_invoice_i = WorksheetFunction.Text(customer_i & sum_statement, "0,00")
            
            ThisWorkbook.Sheets("statement").Cells(i + 1, 36).Value = semi_invoice_i

            ThisWorkbook.Sheets("statement").Cells(i + 1, 29).Value = date_payment_i
            
            ThisWorkbook.Sheets("statement").Cells(i + 1, 30).Value = sum_statement
            
            ThisWorkbook.Sheets("statement").Cells(i + 1, 43).Value = Application.IsError(Application.Match(semi_invoice_i, ThisWorkbook.Sheets("REGISTRY").Range("AF1:AF100000").Value, 0))
            ThisWorkbook.Sheets("statement").Cells(i + 1, 44).Value = semi_invoice_match_i2
            If Application.IsError(Application.Match(semi_invoice_i, ThisWorkbook.Sheets("REGISTRY").Range("AF1:AF100000").Value, 0)) = True Then semi_i = 1 Else semi_i = 0
            
            If Application.IsError(Application.Match(semi_invoice_i, ThisWorkbook.Sheets("REGISTRY").Range("AF1:AF100000").Value, 0)) = True Then semi_invoice_match_i2 = 0 Else semi_invoice_match_i2 = Application.Match(semi_invoice_i, ThisWorkbook.Sheets("REGISTRY").Range("AF1:AF100000").Value, 0)
            ThisWorkbook.Sheets("statement").Cells(i + 1, 45).Value = semi_invoice_match_i2
            ThisWorkbook.Sheets("statement").Cells(i + 1, 47).Value = Application.Match(semi_invoice_i, ThisWorkbook.Sheets("REGISTRY").Range("AF1:AF100000").Value, 0)
            ThisWorkbook.Sheets("statement").Cells(i + 1, 46).Value = semi_invoice_match_i2
            If semi_invoice_match_i2 > 0 Then GoTo SemiAuto Else GoTo ManualProcessing

SemiAuto:
            
            If ThisWorkbook.Sheets("statement").Cells(i + 1, 2).Value = "№ документу" Then GoTo Label1: 'Проверка на ошибки
            
            temp_sum = 0
            
            temp_sum = ThisWorkbook.Sheets("Registry").Cells(semi_invoice_match_i2, 33).Value 'Временно записываем в переменную значение из ячейки
                        
            ThisWorkbook.Sheets("Registry").Cells(semi_invoice_match_i2, 37).Value = "Співставлення по контрагенту та сумі платежу. Необхідний додатковий контроль"
            
            If temp_sum <> 0 Then ThisWorkbook.Sheets("Registry").Cells(semi_invoice_match_i2, 37).Value = "Повтор інвойсу. Співставлення по контрагенту та сумі платежу. Необхідний додатковий контроль"
            
            ThisWorkbook.Sheets("Registry").Cells(semi_invoice_match_i2, 36).Value = sum_statement
            
            ThisWorkbook.Sheets("Registry").Cells(semi_invoice_match_i2, 38).Value = semi_invoice_i
            
            ThisWorkbook.Sheets("Registry").Cells(semi_invoice_match_i2, 33).Value = temp_sum + sum_statement 'Используем временное значение для определения новой суммы
            
            'Дописать в примечание реквизиты платежа - дату операции, дату платежа и сумму платежа
            
            temp_text = ThisWorkbook.Sheets("Registry").Cells(semi_invoice_match_i2, 34).Value 'Временно записываем в переменную текст из ячейки
            
            ThisWorkbook.Sheets("Registry").Cells(semi_invoice_match_i2, 35).Value = date_payment_i
            
            ThisWorkbook.Sheets("Registry").Cells(semi_invoice_match_i2, 34).Value = temp_text & "; " & date_payment_i & " - " & sum_statement 'В ячейке прописываем текст состоящий из временной переменной и добавлением из текущей строки выписки
            
            If ThisWorkbook.Sheets("statement").Cells(i + 1, 1).Value = "№ документу" Then GoTo Label1:
            
            ThisWorkbook.Sheets("statement").Cells(i + 1, 28).Value = "SemiAuto"
            
            ThisWorkbook.Sheets("statement").Cells(i + 1, 32).Value = semi_invoice_match_i2
            
            ThisWorkbook.Sheets("Registry").Cells(semi_invoice_match_i2, 39).Value = "SemiAuto"
            
            Semi_automated_counter = Semi_automated_counter + 1
            
            ThisWorkbook.Sheets("Registry").Cells(semi_invoice_match_i2, 40).FormulaR1C1Local = "=RC[-7]-RC[-23]"
            
            'If ThisWorkbook.Sheets("Registry").Cells(semi_invoice_match_i2, 32).Value = ThisWorkbook.Sheets("Registry").Cells(semi_invoice_match_i2, 16).Value Then ThisWorkbook.Sheets("Registry").Cells(semi_invoice_match_i2, 16).Interior.Color = 52377
            
            GoTo Label1
            '********************************************************************************************
            '   КОПИРОВАНИЕ СТРОК НА ЛИСТ РУЧНОЙ ОБРАБОТКИ
            '********************************************************************************************

            '============================================================================================
            ' Подготовка данных для ручной обработки
            '============================================================================================

ManualProcessing:
            
            If ThisWorkbook.Sheets("statement").Cells(i + 1, 2).Value = "№ документу" Then GoTo Label1:
            
            If ThisWorkbook.Sheets("statement").Cells(i + 1, 2).Value = "" Then GoTo Label1:

            ThisWorkbook.Sheets("statement").Cells(i + 1, 29).Value = date_payment_i
            
            ThisWorkbook.Sheets("statement").Cells(i + 1, 30).Value = sum_statement

            For j = 1 To 22
            
            ThisWorkbook.Sheets("manual_processing").Cells(Manual_counter + 2, j).Value = ThisWorkbook.Sheets("statement").Cells(i + 1, j).Value
            
            Next j
            
            Manual_counter = Manual_counter + 1
            
            ThisWorkbook.Sheets("statement").Cells(i + 1, 28).Value = "Manual"
            
            
            '********************************************************************************************
Label1:     Next i
    
    MsgBox "Вхідних платежів " & Income_counter & " рядків. В автоматичному режимі опрацьовано " & Automated_counter & " рядків. Полуавтоматичне співставлення " & Semi_automated_counter & " рядків. Для ручної обробки залишилось " & Manual_counter & " рядків"
    
    'Обнулить статус бар
    Application.StatusBar = False
        Exit Sub

    '********************************************************************************************
Statement_Processing_Err:
     '   MsgBox Err.Description & vbCrLf & _
      '         "в UAEN03.Statement_Processing.Statement_Processing", _
       '        vbExclamation + vbOKOnly, App.Title
     '   Resume Next
    ' MsgBox "error"
     Resume Next
     
     
End Sub



    '==============================================================================================================
    '==============================================================================================================
    ' Module    : csvimport
    ' Author    : Антон Палихов aka Palant | http://palikhov.wordpress.com | e-mail: a.v.palikhov@hotmail.com
    ' Client    : NPC Ukrenergo
    '==============================================================================================================
    '==============================================================================================================
    Sub load_csv()
    Dim fStr As String

    With Application.FileDialog(msoFileDialogFilePicker)
        .Show
        If .SelectedItems.Count = 0 Then
            MsgBox "Cancel Selected"
            Exit Sub
        End If
        'fStr is the file path and name of the file you selected.
        fStr = .SelectedItems(1)
    End With

    With ThisWorkbook.Sheets("invoices").QueryTables.Add(Connection:= _
    "TEXT;" & fStr, Destination:=Range("invoices!$A$1"))
        .Name = "CAPTURE"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 866
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = True
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False

    End With
End Sub

    '==============================================================================================================
    '==============================================================================================================
    ' Module    : csvimport
    ' Author    : Антон Палихов aka Palant | http://palikhov.wordpress.com | e-mail: a.v.palikhov@hotmail.com
    ' Client    : NPC Ukrenergo
    '==============================================================================================================
    '==============================================================================================================
Public Sub load_csv()
    Dim fStr As String

    With Application.FileDialog(msoFileDialogFilePicker)
        .Show
        If .SelectedItems.Count = 0 Then
            MsgBox "Cancel Selected"
            Exit Sub
        End If
        'fStr is the file path and name of the file you selected.
        fStr = .SelectedItems(1)
    End With

    With ThisWorkbook.Sheets("invoices").QueryTables.Add(Connection:= _
    "TEXT;" & fStr, Destination:=Range("invoices!$A$1"))
        .Name = "CAPTURE"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 866
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = True
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False

    End With
End Sub





'Процедура вставки формул
Public Sub InsertFormula1()

'==============================================================================================================
'==============================================================================================================
' Module    : InsertFormula
' Author    : Антон Палихов aka Palant | http://palikhov.wordpress.com | e-mail: a.v.palikhov@hotmail.com
' Client    : NPC Ukrenergo
'==============================================================================================================
'==============================================================================================================


Dim tRowsCount As Double
Dim tRowsCount2 As Double
Dim i As Double
tRowsCount = ThisWorkbook.Sheets("REGISTRY").Cells(Rows.Count, 1).End(xlUp).Row
tRowsCount2 = ThisWorkbook.Sheets("REGISTRY").Cells(Rows.Count, 3).End(xlUp).Row

With Sheets("registry")
For i = tRowsCount + 1 To tRowsCount2

.Cells(i, 1).FormulaR1C1Local = "=VLOOKUP(RC[2];ДОГОВОРИ!R4C2:R1600C4;2;FALSE)"

.Cells(i, 2).FormulaR1C1Local = "=VLOOKUP(RC[1];ДОГОВОРИ!R4C2:R1600C4;3;FALSE)"

.Cells(i, 16).FormulaR1C1Local = "=ROUND(RC[-1]*1,2;2)"
.Cells(i, 31).FormulaR1C1Local = "=RC[-29]&RC[-14]"

Next i
End With
End Sub

Public Sub Coloration1()

'==============================================================================================================
'==============================================================================================================
' Module    : Coloration
' Author    : Антон Палихов aka Palant | http://palikhov.wordpress.com | e-mail: a.v.palikhov@hotmail.com
' Client    : NPC Ukrenergo
'==============================================================================================================
'==============================================================================================================
Dim i As Double
Dim j As Double
Dim k As Double
Dim f As Double
Dim tRowsCount As Double
tRowsCount = ThisWorkbook.Sheets("REGISTRY").Cells(Rows.Count, 3).End(xlUp).Row
MsgBox tRowsCount
j = 0
For i = 1 To tRowsCount

If ThisWorkbook.Sheets("REGISTRY").Cells(i + 1, 39).Value = 0 And ThisWorkbook.Sheets("REGISTRY").Cells(i + 1, 32).Value <> 0 Then ThisWorkbook.Sheets("Registry").Cells(i + 1, 16).Interior.Color = 52377

'If ThisWorkbook.Sheets("Registry").Cells(i + 1, 13).Value > 0 And ThisWorkbook.Sheets("Registry").Cells(i + 1, 34).Value <> O Then ThisWorkbook.Sheets("Registry").Cells(i + 1, 30).Value = "ПОМИЛКА. ВЖЕ НАЯВНА ДАТА ОПЛАТИ"
If ThisWorkbook.Sheets("REGISTRY").Cells(i + 1, 39).Value = 0 And ThisWorkbook.Sheets("REGISTRY").Cells(i + 1, 32).Value <> 0 And ThisWorkbook.Sheets("Registry").Cells(i + 1, 13).Value = 0 Then ThisWorkbook.Sheets("Registry").Cells(i + 1, 13).Value = ThisWorkbook.Sheets("Registry").Cells(i + 1, 34).Value
If ThisWorkbook.Sheets("REGISTRY").Cells(i + 1, 39).Value = 0 And ThisWorkbook.Sheets("REGISTRY").Cells(i + 1, 32).Value <> 0 Then j = j + 1
'MsgBox ThisWorkbook.Sheets("REGISTRY").Cells(i + 1, 39).Value
'MsgBox ThisWorkbook.Sheets("REGISTRY").Cells(i + 1, 32).Value
'MsgBox "i" & i
'MsgBox "j" & j


Next i

MsgBox "Співпадають суми по " & j & " інвойсів"

tRowsCount = ThisWorkbook.Sheets("manual_processing").Cells(Rows.Count, 3).End(xlUp).Row
For i = 1 To tRowsCount

If Application.IsError(Application.Find("Розподіл", ThisWorkbook.Sheets("manual_processing").Cells(i + 1, 20).Value, 1)) = True Then k = 0 Else k = WorksheetFunction.Find("Розподіл", ThisWorkbook.Sheets("manual_processing").Cells(i + 1, 20).Value, 1)

If k > 0 Then _

    For j = 1 To 22
    
    ThisWorkbook.Sheets("manual_processing").Cells(i + 1, j).Interior.Color = 8388736
    
    Next j

End If
If Application.IsError(Application.Find("борг", ThisWorkbook.Sheets("manual_processing").Cells(i + 1, 20).Value, 1)) = True Then f = 0 Else f = WorksheetFunction.Find("борг", ThisWorkbook.Sheets("manual_processing").Cells(i + 1, 20).Value, 1)

If f > 0 Then _

    For j = 1 To 22
    
    ThisWorkbook.Sheets("manual_processing").Cells(i + 1, j).Interior.Color = 16763904
    
    Next j

End If
Next i


End Sub

Public Sub Calendar_data_inserting1()
'==============================================================================================================
'==============================================================================================================
' Module    : Calendar_data_inserting
' Author    : Антон Палихов aka Palant | http://palikhov.wordpress.com | e-mail: a.v.palikhov@hotmail.com
' Client    : NPC Ukrenergo
'==============================================================================================================
'==============================================================================================================
Dim i As Double ' Счетчик

Dim cRowsRegistryCount As Double 'Количество строк с данными на листе реестра

Dim dYear As String

Dim dMonth As Integer

Dim tMonth As String

Dim dDay As Integer

Dim dTenday As String

cRowsRegistryCount = ThisWorkbook.Sheets("Registry").Cells(Rows.Count, 2).End(xlUp).Row

MsgBox cRowsRegistryCount

For i = 2 To cRowsRegistryCount

'Простановка года

If ThisWorkbook.Sheets("Registry").Cells(i, 5).Value <> "" Then GoTo Label1 'Если в столбце года стоят данные - переходим к следующей строке

If ThisWorkbook.Sheets("Registry").Cells(i, 10).Value = "" Then GoTo Label1 'Если в столбце даты нет данных - переходим к следующей строке

If ThisWorkbook.Sheets("Registry").Cells(i, 10).Value = 0 Then GoTo Label1 'Если в столбце даты нет данных - переходим к следующей строке

dYear = Year(ThisWorkbook.Sheets("Registry").Cells(i, 10).Value) 'Присвоить временной переменной значения года из столбца даты
 
ThisWorkbook.Sheets("Registry").Cells(i, 5).Value = dYear 'Записать в столбец года значение временной переменной

If ThisWorkbook.Sheets("Registry").Cells(i, 6).Value <> "" Then GoTo Label1 'Если в столбце месяца стоят данные - переходим к следующей строке

dMonth = Month(ThisWorkbook.Sheets("Registry").Cells(i, 10).Value) 'Присвоить временной переменной значения месяца из столбца даты
'На основании числа месяца определим текстовую строку его названия и запишем во временную перменную
If dMonth = 1 Then tMonth = "січень"
If dMonth = 2 Then tMonth = "лютий"
If dMonth = 3 Then tMonth = "березень"
If dMonth = 4 Then tMonth = "квітень"
If dMonth = 5 Then tMonth = "травень"
If dMonth = 6 Then tMonth = "червень"
If dMonth = 7 Then tMonth = "липень"
If dMonth = 8 Then tMonth = "серпень"
If dMonth = 9 Then tMonth = "вересень"
If dMonth = 10 Then tMonth = "жовтень"
If dMonth = 11 Then tMonth = "листопад"
If dMonth = 12 Then tMonth = "грудень"

ThisWorkbook.Sheets("Registry").Cells(i, 6).Value = tMonth 'Запишем в столбец месяца значение временной переменной

If ThisWorkbook.Sheets("Registry").Cells(i, 7).Value <> "" Then GoTo Label1 'Если в столбце декады стоят данные - переходим к следующей строке

dDay = Day(ThisWorkbook.Sheets("Registry").Cells(i, 10).Value) 'По дате определяем день месяца
'По дню месяца определяем декаду и записываем во временную переменную
If dDay > 0 And dDay < 11 Then dTenday = 1
If dDay > 10 And dDay < 21 Then dTenday = 2
If dDay > 20 And dDay < 32 Then dTenday = 3

ThisWorkbook.Sheets("Registry").Cells(i, 7).Value = dTenday 'Записываем значение временной переменной в столбец декады

tMonth = ""
dTenday = ""
dYear = ""

Label1: Next i
End Sub


Public Sub Clearing4()
'==============================================================================================================
'==============================================================================================================
' Module    : Clearing
' Author    : Антон Палихов aka Palant | http://palikhov.wordpress.com | e-mail: a.v.palikhov@hotmail.com
' Client    : NPC Ukrenergo
'==============================================================================================================
'==============================================================================================================
Dim cRowsStatementCount As Double

Dim cRowsManualCount As Double

Dim cRowsControlCount As Double

Dim oRangeStatement As Range

Dim oRangeControl As Range

Dim oRangeManual As Range

Dim i, j As Integer
'Очистка данных на листе выписки
cRowsStatementCount = ThisWorkbook.Sheets("statement").Cells(Rows.Count, 28).End(xlUp).Row

Set oRangeStatement = ThisWorkbook.Sheets("statement").Range(Cells(2, 28).Address, Cells(cRowsStatementCount + 1, 38).Address)

oRangeStatement.ClearContents


'Очистка данных на листе ручной обработки

cRowsManualCount = ThisWorkbook.Sheets("manual_processing").Cells(Rows.Count, 2).End(xlUp).Row

Set oRangeManual = ThisWorkbook.Sheets("manual_processing").Range(Cells(2, 1).Address, Cells(cRowsManualCount + 1, 23).Address)

oRangeManual.ClearContents

For i = 2 To cRowsManualCount
    
    For j = 1 To 22
    
    ThisWorkbook.Sheets("manual_processing").Cells(i, j).Interior.Color = xlNone
    
    Next j
    
Next i

'Очистка данных на листе обработки контрольных значений
cRowsControlCount = ThisWorkbook.Sheets("control_processing").Cells(Rows.Count, 2).End(xlUp).Row

Set oRangeControl = ThisWorkbook.Sheets("control_processing").Range(Cells(2, 1).Address, Cells(cRowsControlCount + 1, 23).Address)

oRangeControl.ClearContents

'Очистка завершена. Выводим сообщение

MsgBox "Проміжні дані на аркушах виписки, ручної обробки та обробки контродьних значень очищені"

End Sub

Sub Clear_registry_additional_data()
'
' Макрос4 Макрос
'

Worksheets("registry").Range("Af2:bz200000").ClearContents
End Sub










