' ====================================================================
' VBA МАКРОС ДЛЯ EXCEL - DADATA API
' Получение адреса, широты и долготы по кадастровому номеру
' ====================================================================

Sub GetCadastreData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cadNum As String
    Dim apiKey As String
    Dim apiUrl As String
    Dim xmlHttp As Object
    Dim json As String
    Dim value As String
    Dim lat As String
    Dim lon As String
    
    ' ============ ВАЖНО: Установите ваш API ключ ============
    apiKey = "YOUR_DADATA_API_KEY"
    apiUrl = "https://suggestions.dadata.ru/suggestions/api/4_1/rs/findById/address"
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Проверяем наличие заголовков в первой строке
    If lastRow < 2 Then
        MsgBox "Добавьте кадастровые номера в колонку A, начиная со строки 2!", vbExclamation
        Exit Sub
    End If
    
    ' Начинаем со строки 2 (строка 1 - заголовки)
    For i = 2 To lastRow
        cadNum = Trim(ws.Cells(i, 1).Value)
        
        If cadNum <> "" Then
            On Error Resume Next
            
            ' Создаем HTTP запрос
            Set xmlHttp = CreateObject("MSXML2.XMLHTTP")
            
            With xmlHttp
                .Open "POST", apiUrl, False
                .setRequestHeader "Content-Type", "application/json"
                .setRequestHeader "Authorization", "Token " & apiKey
                .Send "{""query"": """ & cadNum & """}"
                
                If .Status = 200 Then
                    json = .ResponseText
                Else
                    ws.Cells(i, 2).Value = "ERROR: " & .Status
                    ws.Cells(i, 3).Value = ""
                    ws.Cells(i, 4).Value = ""
                    GoTo NextIteration
                End If
            End With
            
            ' Парсим JSON и извлекаем нужные поля
            value = ExtractJSONValue(json, "value")
            lat = ExtractJSONValue(json, "geo_lat")
            lon = ExtractJSONValue(json, "geo_lon")
            
            ' Записываем в ячейки
            ws.Cells(i, 2).Value = value
            ws.Cells(i, 3).Value = lat
            ws.Cells(i, 4).Value = lon
            
            ' Статус в консоль
            Debug.Print "Строка " & i & ": " & cadNum & " - OK"
            
NextIteration:
            ' Задержка между запросами (DaData имеет лимиты)
            Application.Wait (Now + TimeValue("0:00:00.5"))
            
            On Error GoTo 0
        End If
    Next i
    
    MsgBox "Обработка завершена! Проверьте результаты в колонках B, C, D", vbInformation
End Sub


Function ExtractJSONValue(jsonText As String, fieldName As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim result As String
    Dim searchStr As String
    
    ' Ищем поле в JSON (значение в кавычках)
    searchStr = """" & fieldName & """:"""
    startPos = InStr(jsonText, searchStr)
    
    If startPos > 0 Then
        ' Переходим к началу значения
        startPos = startPos + Len(searchStr)
        ' Находим конец значения (следующая кавычка)
        endPos = InStr(startPos, jsonText, """")
        
        If endPos > 0 Then
            result = Mid(jsonText, startPos, endPos - startPos)
            ExtractJSONValue = result
            Exit Function
        End If
    End If
    
    ' Если значение - число (для geo_lat, geo_lon)
    searchStr = """" & fieldName & """:"
    startPos = InStr(jsonText, searchStr)
    
    If startPos > 0 Then
        startPos = startPos + Len(searchStr)
        ' Находим конец числа (запятая или скобка)
        endPos = InStr(startPos, jsonText, ",")
        If endPos = 0 Then
            endPos = InStr(startPos, jsonText, "}")
        End If
        
        If endPos > 0 Then
            result = Trim(Mid(jsonText, startPos, endPos - startPos))
            result = Replace(result, """", "")
            ExtractJSONValue = result
            Exit Function
        End If
    End If
    
    ExtractJSONValue = "NOT_FOUND"
End Function
