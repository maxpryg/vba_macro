Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
' коллекция с идентификаторами накладных
' нужна для создания ведомости передачи груза
Public refCollection As Collection


Sub Send_Mail_Mass()
    Dim objOutlookApp As Object, objMail As Object
    Dim sTo As String, sSubject As String, sBody As String, sAttachment As String
    Dim lr As Long, lLastR As Long

    Application.ScreenUpdating = False
    On Error Resume Next
    
    'пробуем подключиться к Outlook, если он уже открыт
    Set objOutlookApp = GetObject(, "Outlook.Application")
    Err.Clear 'Outlook закрыт, очищаем ошибку
    
    If objOutlookApp Is Nothing Then
        Set objOutlookApp = CreateObject("Outlook.Application")
    End If
    
    'произошла ошибка создания объекта - выход
    If Err.number <> 0 Then Set objOutlookApp = Nothing: Set objMail = Nothing: Exit Sub
    'objOutlookApp.Session.Logon "user","1234",False, True

    lLastR = Cells(Rows.Count, 1).End(xlUp).Row
    'цикл от второй строки(начало данных с адресами) до последней ячейки таблицы
    For lr = 2 To lLastR
        Set objMail = objOutlookApp.CreateItem(0)   'создаем новое сообщение
        
        
        
        'тело сообщения
        Dim mailbody As String
            mailbody = "<p>Доброго дня,</p>" & _
                    "<p>Ми відправили Ваше замовлення:</p>" & _
                    "<p>&nbsp;&nbsp;&nbsp;&nbsp;кількість місць - " & Cells(lr, 7).Value & "</p>" & _
                    "<p>&nbsp;&nbsp;&nbsp;&nbsp;загальна вага - " & Cells(lr, 8).Value & "</p>" & _
                    "<p>&nbsp;&nbsp;&nbsp;&nbsp;номер ТТН - " & Cells(lr, 4).Value & "</p>" & _
                    "<p>Для відстежування замовлення можете перейти за посиланням <a href='https://www.sat.ua/ru/treking/tracking/'>ТК САТ</a></p>"

        'создаем сообщение
        With objMail
            .To = Cells(lr, 1).Value 'адрес получателя
            '.Subject = Cells(lr, 2).Value 'тема сообщения
            .Subject = "САТ - номер ТТН, " & Cells(lr, 3).Value 'тема сообщения
            '.Body = Cells(lr, 4).Value 'текст сообщения
            '.Body = "Доброго дня," & vbNewLine & "Ми відправили Ваше замовлення:" & vbNewLine
            .BodyFormat = olFormatHTML
            .HTMLBody = "<html><head></head><body>" & mailbody & "</body></html>"

            'вложение(если ячейка не пустая и путь к файлу указан правильно)
            'If Cells(lr, 4).Value <> "" Then
            '   If Dir(Cells(lr, 4).Value, 16) <> "" Then
            '        .Attachments.Add Cells(lr, 4).Value
            '    End If
            'End If
            
            .Display 'Send, Display, если необходимо просмотреть сообщение, а не отправлять без просмотра
        End With
    Next lr

    Set objOutlookApp = Nothing: Set objMail = Nothing
    Application.ScreenUpdating = True
End Sub


'VBA function to send HTTP POST to a server:
Function httpPost(method, url, body)
    ' создание подключения
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        ' метод и адрес запроса
        '.Open "POST", "https://urm.sat.ua/openws/hs/api/v2.0/documents/nng/json/save", False
         .Open method, url, False
        ' добавляем заголовки запроса
        .setRequestHeader "accountref", "d9192b3b-28d7-11eb-941c-00505601031c" ' ключи аккаунта
        .setRequestHeader "apikey", "d86777dd-33dd-4233-9f71-3ee9c0d3d831" ' ключ к API
        .setRequestHeader "app", "cabinet" ' типа заходим через кабинет
        .send body ' отправляем тело запроса
        'httpPost = .responseBody ' получаем ответ запроса
    
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''START BLOCK'''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' читаем ответ запроса побайтово, чтобы правильно прочитать кирилицу
        Dim responseText As Variant
        Set fileStream = CreateObject("ADODB.Stream")
        fileStream.Open
        fileStream.Type = 1 'Binary
        fileStream.Write .responseBody
        fileStream.Position = 0
        fileStream.Type = 2 'Text
        fileStream.Charset = "UTF-8"
        text = fileStream.ReadText
        fileStream.Close
        'MsgBox text
        '''''''''''END BLOCK''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    End With
    
'    Dim Json As Object, nngReturnState As String
'    Set Json = JsonConverter.ParseJson(text)
    'httpPost = Json("success")
    httpPost = text

End Function


Sub saveStickers()

    Dim nngRef As Variant
    Dim body As String
    
    ' создаем словарь и добавляем в него пары ключи:значение с данными для создания ведомости
    Dim dictBody As Object
    Set dictBody = CreateObject("Scripting.Dictionary")
    
'    Set refCollection = New Collection
'    refCollection.Add ("e5e99f0c-d39b-4378-91cb-93b8737eb9e1")
    
    body = "["
    For Each nngRef In refCollection
        dictBody.Add "ref", nngRef
        body = body + JsonConverter.ConvertToJson(dictBody) + ","
        dictBody.Remove "ref"
    Next nngRef
    body = body + "]"
    Debug.Print (body)
    
    Set WinHttpReq = CreateObject("WinHTTP.WinHTTPrequest.5.1")
    With WinHttpReq
        ' метод и адрес запроса
        .Open "POST", "https://urm.sat.ua/openws/hs/api/v2.0/documents/nng/json/print", False
        ' добавляем заголовки запроса
        .setRequestHeader "accountref", "d9192b3b-28d7-11eb-941c-00505601031c" ' ключи аккаунта
        .setRequestHeader "apikey", "d86777dd-33dd-4233-9f71-3ee9c0d3d831" ' ключ к API
        .setRequestHeader "app", "cabinet" ' типа заходим через кабинет
        .send body ' отправляем тело запроса
        'httpPost = .responseBody ' получаем ответ запроса
        
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Type = 1
        oStream.Open
        oStream.Write WinHttpReq.responseBody
        oStream.SaveToFile "C:\Users\E9926628\Desktop\SAT_stickers.pdf", 2
        'oStream.SaveToFile ("C:\Users\E9926628\Desktop\SAT_stickers.pdf")
        oStream.Close

    End With
     
    MsgBox ("Стикера на рабочем столе. Смотри файл - SAT_stickers.pdf")
    
End Sub


Sub createTransferList()

    Dim today As String
    today = Date
    
    ' создаем словарь и добавляем в него пары ключи:значение с данными для создания ведомости
    Dim dictBody As Object
    Set dictBody = CreateObject("Scripting.Dictionary")

    dictBody.Add "contact", "Максим"
    dictBody.Add "phone", "+380675559537"
    dictBody.Add "sender", "14c4099f-c829-4183-b04a-b0c6fb8bac8f"
    dictBody.Add "senderAddress", "Дніпровський район, вулиця Олекси Довбуша, 37"
    dictBody.Add "writeMode", "save"
    dictBody.Add "dateCompletion", today
    dictBody.Add "nngList", refCollection
    
    'конвертируем словарь в JSON строку и передаем функции для отправки на сервер
    Dim body As String, response As String
    body = JsonConverter.ConvertToJson(dictBody)
    'Debug.Print JsonConverter.ConvertToJson(dictBody, Whitespace:=2)
    response = httpPost("POST", "https://urm.sat.ua/openws/hs/api/v2.0/documents/spreadSheet/json/save", body)
    
End Sub


Sub createNNG()
''''''''' TODO''''''''''''''''''''''''''''''
    ' счетчик ошибок, раскраска строк с ошибками, вывод окна с количеством ошибок +++
    ' плательщик - МЫ, ОНИ +
    ' очистка фильтров и форматирования ошибок ++
    ' безналичная форма оплаты для получателей +
    ' добавить прогресс бар +
    ' создиние ведомости в случае положительного овтета +
    ' сохранять где-то файл перед запуском макроса +
    ' адрессная доставка +
    ' выводить окно с ошибками +


    ' сохраняем файл
    ActiveWorkbook.Save

    ' переменные для накладной на груз
    Dim recipient As String, rspRecipient As String, townRecipient As String, contactRecipient As String, recipientPhone As String
    Dim seatsAmount As String, weight As String, payerType As String, paymentMethod As String, phoneForSMS As String
    Dim additionalInformation As String
    
    ' активируем нужный листи и фильтруем строки, где количество мест не равно нулю
    ThisWorkbook.Worksheets("Contacts").Activate
    ThisWorkbook.Worksheets("Contacts").ListObjects("Contacts").Range.AutoFilter Field:=2, Criteria1:=">0", Operator:=xlAnd


    ''''''''''''''''Progress Bar''''''''''''''''''''''''''''''''''''''''''''''''''
    '(Step 1) Display your Progress Bar
    Dim pctdone As Single
    ufProgress.LabelProgress.Width = 0
    ufProgress.Show
    ''''''''''''''''Progress Bar'''''''''''''''''''''''''''''''''''''''''''''''''

    ' цикл по отфильтрованным строкам
    ' начинаем со второй строки, первая - названия столбцов
    Dim nng As Range, nngList As Range
    Dim falseCounter As Integer, nngCounter As Long ' falseCounter - счетчик ответов false, nngCounter - счетчик ННГ
    Dim nngQty As Long ' количество накладных
    
    Set nngList = Range("A2", Cells(Rows.Count, "A").End(xlUp)).SpecialCells(xlCellTypeVisible) ' список накладных к созданию
    nngQty = nngList.Count ' колиество накладных, которые будут созданы
    
    ' коллекция с идентификаторами накладных
    ' нужна для создания ведомости передачи груза
    Set refCollection = New Collection
    
    For Each nng In nngList
        nngCounter = nngCounter + 1 ' считаем накладные
        
        ''''''''''''''''Progress Bar''''''''''''''''''''''''''''''''''''''''''''''''''
        '(Step 2) Periodically update progress bar
        pctdone = nngCounter / nngQty
        With ufProgress
            .LabelCaption.Caption = "Создаю ННГ " & nngCounter & " из " & nngQty
            .LabelProgress.Width = pctdone * (.FrameProgress.Width)
        End With
        DoEvents
        ''''''''''''''''Progress Bar''''''''''''''''''''''''''''''''''''''''''''''''''
        
        ' создаем словарь и добавляем в него пары ключи:значение с данными для создания ННГ
        Dim dictBody As Object
        Set dictBody = CreateObject("Scripting.Dictionary")
        
         'значения, которые одинаковые во всех ННГ - константы
        dictBody.Add "contactSender", "Максим"
        dictBody.Add "senderPhone", "+380675559537"
        dictBody.Add "sender", "14c4099f-c829-4183-b04a-b0c6fb8bac8f"
        dictBody.Add "senderAddress", "Дніпровський район, вулиця Олекси Довбуша, 37"
        dictBody.Add "townSender", "8d7f5ea4-9436-11dd-98c6-001cc0108cd1"
        dictBody.Add "description", "44f066e6-dede-497b-aadb-ad149a993558"
        dictBody.Add "cargoType", "Базовый"
        dictBody.Add "writeMode", "save"
        
        'переменные значения ННГ
        dictBody.Add "recipient", Cells(nng.Row, 14).Value
        dictBody.Add "rspRecipient", Cells(nng.Row, 16).Value
        dictBody.Add "townRecipient", Cells(nng.Row, 15).Value
        dictBody.Add "contactRecipient", Cells(nng.Row, 10).Value
        dictBody.Add "recipientPhone", Cells(nng.Row, 9).Value
        dictBody.Add "seatsAmount", Cells(nng.Row, 2).Value
        dictBody.Add "weight", Cells(nng.Row, 3).Value
        'dictBody.Add "additionalInformation", ""
        dictBody.Add "phoneForSMS", Cells(nng.Row, 9).Value
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''' START BLOCK Плательщик и форма оплаты'''''''''''
        ' добавляем плательщика и форму оплаты
        ' если плательщик - "мы" или "ми", тогда оплата  по безналу
        If StrComp(Cells(nng.Row, 4).Value, "мы", vbTextCompare) = 0 Or StrComp(Cells(nng.Row, 4).Value, "ми", vbTextCompare) = 0 Then
            dictBody.Add "payerType", "Отправитель"
            dictBody.Add "paymentMethod", "NonCash"
        ElseIf StrComp(Cells(nng.Row, 4).Value, "они", vbTextCompare) = 0 Then
            dictBody.Add "payerType", "Получатель"
            dictBody.Add "paymentMethod", "Cash"
            
            ' если плательщик "они", то оплата определяеться по тексту из ячейки формы оплаты
            ' если ячейка не содержит текста, то оплата наличкой
            If Cells(nng.Row, 5) <> "" Then
                dictBody.Add "additionalInformation", "Платник - ОТРИМУВАЧ; Спосіб оплати - БЕЗГОТІВКОВИЙ"
            End If
        End If
        ''''''''''' END  BLOCK Плательщик и форма оплаты' ''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''' START BLOCK Доставка до дверей'''''''''''''''''''''''''''''
        ' если стоит "да" в графе доставка
        If StrComp(Cells(nng.Row, 7).Value, "да", vbTextCompare) = 0 Then
            dictBody.Add "delivery", True
            dictBody.Add "recipientAddress", Cells(nng.Row, 8).Value
        End If
        ''''''''''' END  BLOCK Доставка до дверей' '''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'конвертируем словарь в JSON строку и передаем функции для отправки на сервер
        Dim body As String, response As String
        body = JsonConverter.ConvertToJson(dictBody)
        response = httpPost("POST", "https://urm.sat.ua/openws/hs/api/v2.0/documents/nng/json/save", body)

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''' START BLOCK Обрабатываем ответ от сервера '''''''''''''''''''
        Dim Json As Object, nngReturnState As String, nngRef As String
        Set Json = JsonConverter.ParseJson(response)
        nngReturnState = Json("success")
        
        ' Если накладная успешно создана, добавляем ее идентфикатор в коллекцию накладных
        If nngReturnState = True Then
            nngRef = Json("data")(1)("ref")
            refCollection.Add nngRef
        Else
        ' считаем накладные созданные с ошибками и раскрашиваем строки с этими накладными
            falseCounter = falseCounter + 1
            Range(Cells(nng.Row, 1), Cells(nng.Row, 10)).Interior.Color = 2569911
            MsgBox Cells(nng.Row, 1).Value & vbNewLine & response, vbExclamation, Title:="ОШИБКА при создании " & Cells(nng.Row, 1).Value
        End If
        '''''''''''' END BLOCK Обрабатываем ответ от сервера '''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        ''''''''''''''''Progress Bar''''''''''''''''''''''''''''''''''''''''''''''''''
        '(Step 3) Close the progress bar when you're done
        If nngCounter = nngQty Then Unload ufProgress
        ''''''''''''''''Progress Bar''''''''''''''''''''''''''''''''''''''''''''''''''
        
    Next nng

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''' START BLOCK Запрос на создание ведомости '''''''''''''''''''''
    Dim answer As Variant
    answer = MsgBox("Создать ведомость?", vbYesNo + vbQuestion)
    If answer = vbYes Then
        ' создать ведомость  передачи груза
            createTransferList
    ElseIf answer = vbNo Then
        ' не создавать ведомость
        'Exit Sub
    End If
    '''''''''''' END BLOCK Запрос на создание ведомости ''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''' START BLOCK Запрос на печать стикеров '''''''''''''''''''''
    Dim answer2 As Variant
    answer2 = MsgBox("Сохраняем стикера?", vbYesNo + vbQuestion)
    If answer2 = vbYes Then
        ' создать ведомость  передачи груза
            saveStickers
    ElseIf answer2 = vbNo Then
        ' не создавать ведомость
        'Exit Sub
    End If
    '''''''''''' END BLOCK Запрос на создание ведомости ''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If falseCounter > 0 Then
        MsgBox "При создании накладных произошло ошибок - " & falseCounter & "." & vbNewLine & "Строки с ошибками выделены красным цветом."
    Else
        MsgBox "Накладные созданы. Проверяй в кабинете."
    End If
    
End Sub


Sub testJsonFile()
    ' Advanced example: Read .json file and load into sheet (Windows-only)
    ' (add reference to Microsoft Scripting Runtime)
    ' {"values":[{"a":1,"b":2,"c": 3},...]}
    
    Dim FSO As New FileSystemObject
    Dim JsonTS As TextStream
    Dim JsonText As String
    Dim Parsed As Dictionary
    
    
    ' Read .json file
    Set JsonTS = FSO.OpenTextFile("C:\Users\E9926628\Desktop\example.json", ForReading)
    JsonText = JsonTS.ReadAll
    JsonTS.Close

    Debug.Print JsonText
    ' Parse json to Dictionary
    ' "values" is parsed as Collection
    ' each item in "values" is parsed as Dictionary
    Set Parsed = JsonConverter.ParseJson(JsonText)

    ' Prepare and write values to sheet
    Dim Values As Variant
    ReDim Values(Parsed("values").Count, 3)

    Dim Value As Dictionary
    Dim i As Long

    i = 0
    For Each Value In Parsed("values")
      Values(i, 0) = Value("a")
      Values(i, 1) = Value("b")
      Values(i, 2) = Value("c")
      i = i + 1
    Next Value

    Sheets("example").Range(Cells(1, 1), Cells(Parsed("values").Count, 3)) = Values
End Sub


Sub testArrayToJson()

    Dim dictBody As Object
    Set dictBody = CreateObject("Scripting.Dictionary")
    
    Dim arr(0 To 1) As Long
    
    Dim arrMarks(0 To 3) As Long

    ' Set the value of position 0
    arrMarks(0) = 1

    ' Set the value of position 3
    arrMarks(3) = 46
    
    '
    
     Dim collMarks As New Collection
     collMarks.Add "14c4099f-c829-4183-b04a-b0c6fb8bac8f"
     collMarks.Add "14c4099f-c829-4183-b04a-b0c6fb"
     dictBody.Add "nngList", refCollection
    
    Dim body As String
    body = JsonConverter.ConvertToJson(dictBody)
    MsgBox JsonConverter.ConvertToJson(body, Whitespace:=2)
    
End Sub

Sub saveFile()
    Dim today As String, satFilename As String
    today = Format(Now(), "DD-MM-YY")
    satFilename = "C:\Users\E9926628\Documents\D\SAT_bak\SAT_" & today
    ActiveWorkbook.SaveAs filename:=satFilename, FileFormat:=xlOpenXMLWorkbookMacroEnabled
End Sub


Sub duplicateRecipient()
    ' полностью дублирует активную строку
    Dim cellRow As String
    cellRow = ActiveCell.Row
    Rows(cellRow).EntireRow.Insert
    Rows(cellRow + 1).Copy Rows(cellRow)
    
End Sub

Sub searchRecipient()
    ' окно для поиска и перехода к получателю
    Dim userInput As String, foundCell As Range, recipientNameColumn As Range
    Columns("A:A").Select
    userInput = Application.InputBox("Введи название получателя", Type:=3)
    If userInput = "False" Then Exit Sub: Rem user pressed Cancel
'    Set foundCell = Cells.Find(What:=userInput, After:=ActiveCell, LookIn:=xlValues, LookAt:= _
'    xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
    
    Set foundCell = Selection.Find(What:=userInput, After:=ActiveCell, LookIn:=xlValues, LookAt:= _
    xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
    
    If foundCell Is Nothing Then MsgBox "Not found": Exit Sub
    Application.Goto foundCell.Offset(, 1)
    
End Sub
