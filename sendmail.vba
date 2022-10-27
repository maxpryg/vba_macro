Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Option Explicit


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
    If Err.Number <> 0 Then Set objOutlookApp = Nothing: Set objMail = Nothing: Exit Sub
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



Sub findMistakes()
    ' закрашивает красным цветом строки с предполагаемыми ошибками в накладных на груз САТа
    
    Dim n As Integer, lastRow As Long, errCounter As Integer
    
    ' последняя строка таблицы
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
    Dim recipient As String
    Dim rspNum As String
    
    errCounter = 0 ' инициализирум счетчик ошибок

    ' цикл по строкам с накладными на груз
    ' начинаем со второй, первая - названия столбцов
    For n = 2 To lastRow
        
        recipient = Cells(n, 3) ' получатель
        rspNum = Cells(n, 5) ' номер склада
        
        Select Case True
            ' Выбираем нужных получаталей и проверяй номер склада, куда отправили
            Case recipient Like "*Юнона*" And rspNum <> "Львов № 3"
                ThisWorkbook.Worksheets("Розсилка").Range(Cells(n, 1), Cells(n, 15)).Interior.Color = 2569911
                errCounter = errCounter + 1
   
         End Select

    Next n
    
    MsgBox ("Проверка складов завершена. Найдено  " & errCounter & " ошибок")
    
End Sub
