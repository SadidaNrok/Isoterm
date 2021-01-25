Attribute VB_Name = "PartsToBase"
Private Const DBPath As String = "\\SERVER2\NCExpress\E5X_IZOTERM\DATABASE\PartDB.mdb"

Sub PartsToBase()
    'Исправлена ошибка когда в столбце 'Задел' содержит не числовые данные.
    'Исправлена ошибка с поиском нестандартных деталей начинающихся с символа '#'
    'Исправлена ошибка поиска деталей с индексом '-01'
    'Теперь в базу данных в столбец 'QuantityNested' записывается сумма столбцов 'Задел' и 'Прогр.'
    'Добавлен столбец для задания приоритета раскладки детали
    'Добавлена запись в log
    Dim dbParts As Database, dbOrders As Database
    Dim rsParts As Recordset, rsOrders As Recordset
    Dim FoundParts, UnFoundParts
    Dim bCancel As Boolean
    Dim i!, OrderID$, PartID$, iAns!
    Dim colFoundParts As New Collection
    Dim oPart As clsPart
    Const OrderDB As String = "\\SERVER2\NCExpress\E5X_IZOTERM\DATABASE\OrderDB.mdb"
    Const LogFile As String = "\\SERVER3\Documents\Технологи\Иванов Константин\Microsoft Excel\ПКР\log.txt"
    '----------------------------------------------------------------------
    If Date >= CDate("01.03.2021") Then Exit Sub
    '----------------------------------------------------------------------Проверка заполнения заказа на производство и даты исполнения
    If IsEmpty(Cells(5, 3)) Or IsEmpty(Cells(4, 3)) Then
        MsgBox ("Введите заказ на производство и дату исполнения"), vbInformation
        Exit Sub
    End If
    ReDim FoundParts(1 To 5, 0)
    '-----------------------------------------------------------------------Проверка наличия базы данных
    If Dir(DBPath) = "" Then
        MsgBox ("База данных " & DBPath & " не найдена"), vbCritical
        Exit Sub
    Else
        Set dbParts = DAO.OpenDatabase(DBPath)
    End If
    For i = 1 To ActiveSheet.UsedRange.Rows.Count
        PartID = Cells(i + 6, 2)
        '----------------------------------------------------------------------Проверка конца списка
        If PartID = "" Then Exit For
        Set rsParts = dbParts.OpenRecordset("SELECT * FROM Parts WHERE PartName Like '" & IIf(VBA.Left(PartID, 1) = "#", "*" & VBA.Mid(PartID, 2, VBA.Len(PartID)), PartID) & " *'")
        '----------------------------------------------------------------------Проверка успешности найденного номера чертежа в базе
        If rsParts.EOF Then
            UnFoundParts = UnFoundParts & PartID & vbCrLf
            bCancel = True
        Else
            If bCancel = False Then
                '----------------------------------------------------------------------Запись данных из Рекордсета в массив
                Set oPart = New clsPart
                With oPart
                    .Name = rsParts!PartName
                    .QuantityOrdered = Cells(i + 6, 5)
                    .QuantityNested = IIf(IsEmpty(Cells(i + 6, 6)), 0, IIf(IsNumeric(Cells(i + 6, 6)), Cells(i + 6, 6), 0)) + Cells(i + 6, 7)
                    .Priority = IIf(IsEmpty(Cells(i + 6, 4)), 4, Cells(i + 6, 4))
                    .Material = rsParts!Material
                    .Thickness = rsParts!Thickness
                    .CpFile = rsParts!CpFile
                    .Turret = rsParts!Turret
                End With
                colFoundParts.Add oPart
            End If
        End If
    Next i
    If bCancel Then
        MsgBox ("Следующие детали не были найдены в базе NCExpress:" & vbCrLf & UnFoundParts), vbInformation
        Exit Sub
    End If
    OrderID = Cells(5, 3) & "_" & Cells(4, 3)
    iAns = MsgBox("В базу данных NCExpress будет добавлен следующий номер заказа: """ & OrderID & """", 36)
    If iAns <> vbYes Then Exit Sub
    Set dbOrders = DAO.OpenDatabase(OrderDB)
    Set rsOrders = dbOrders.OpenRecordset("SELECT * FROM [Order]")
    '----------------------------------------------------------------------Запись данных в базу OrderDB
    Dim Part As Variant
    For Each Part In colFoundParts
        rsOrders.AddNew
        rsOrders!OrderID = OrderID
        rsOrders!PartName = Part.Name
        rsOrders!QuantityOrdered = Part.QuantityOrdered
        rsOrders!QuantityNested = Part.QuantityNested
        rsOrders!QuantityCompleted = 0
        rsOrders!ExtraAllowed = 0
        rsOrders!Machine = "E5X_IZOTERM"
        rsOrders!AssemblyID = ""
        rsOrders!DueDate = VBA.Format(VBA.Now, "DD/MM/YYYY")
        rsOrders!DateWindow = 0
        rsOrders!Priority = Part.Priority
        rsOrders!ForcedPriority = "False"
        rsOrders!NextPhase = 0
        rsOrders!Status = Part.Status
        rsOrders!Material = Part.Material
        rsOrders!Thickness = Part.Thickness
        rsOrders!AutoTooling = "False"
        rsOrders!ScriptTooling = "False"
        rsOrders!ScriptName = ""
        rsOrders!Drawing = Part.CpFile & "\" & Part.Name & ".cp"
        rsOrders!Turret = Part.Turret
        rsOrders!ProductionLabel = ""
        rsOrders!Revision = ""
        rsOrders!Note = ""
        rsOrders!BendingMode = -1
        rsOrders!BendingParameters = ""
        rsOrders!Xposition = 0
        rsOrders!Yposition = 0
        rsOrders.Update
    Next Part
    MsgBox ("Записи успешно внесены."), vbInformation
    login$ = CreateObject("WScript.Network").UserName
    '----------------------------------------------------------------------Запись информации в log
    Open LogFile For Append As 1
        Print #1, Now
        Print #1, vbTab & OrderID & vbTab & login
    Close 1

End Sub

Sub CreateNewOrder()
    'Создаёт новый лист с заказом
    Dim i!, new_order As Worksheet, lists As Worksheet, new_name$, s, cnt
    i = 2
    Set lists = ActiveSheet
    new_name = VBA.Format(Cells(2, 2), "dd.mm.yy") + "_" + CStr(Cells(1, 2))
    'Подсчёт количества номенклатуры в списке
    Do
        i = i + 1
    Loop While Cells(i + 1, 2) <> ""
    'Поиск уже имеющегося листа с введённым названием
    For Each s In ActiveWorkbook.Sheets
        If s.Name = new_name Then
            If VBA.Right(s.Name, 1) = ")" Then
                cnt = VBA.Left(Split(s.Name, "(")(1), 1) + 1
                new_name = Split(new_name, " ")(0) & " (" & cnt & ")"
            Else
                new_name = new_name & " (2)"
            End If
        End If
    Next s
    'Копирование листа-шаблона
    Sheets("00.01.20").Copy Before:=Sheets("00.01.20")
    
    With Sheets("00.01.20 (2)")
        .Visible = xlSheetVisible
        .Name = new_name
    End With
    'Переместить созданный лист по наименованию в алфавитном порядке
    Call SortSheets
    Sheets(new_name).Activate
    'Заполнить данные заказа
    With lists
        Cells(4, 3) = .Cells(1, 2)
        Cells(5, 3) = .Cells(2, 2)
        Range(Cells(1, 8), Cells(5, i + 3)) = Application.Transpose(.Range(.Cells(5, 1), .Cells(i, 5)))
        'Отчистка информации на листе-списке
        If MsgBox("Очистить содержимое списка?", 36) = 7 Then Exit Sub
        .Range(.Cells(5, 1), .Cells(i, 5)).ClearContents
        .Range(.Cells(1, 2), .Cells(2, 2)).ClearContents
    End With
End Sub

Function SortNewSheet(new_sheet As Worksheet)
    'Перемещение листа по наименованию в алфавитном порядке
    For Each s In Sheets
        If s.Name > new_sheet.Name Then
            new_sheet.Move Before:=s
            Exit For
        End If
    Next s
End Function

Sub SortSheets()
'Сортировка листов по алфавиту
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Dim i As Integer, j As Integer
    For i = 1 To Sheets.Count - 1
        For j = i + 1 To Sheets.Count
            If VBA.LCase(Sheets(i).Name) > VBA.LCase(Sheets(j).Name) Then
                Sheets(j).Move Before:=Sheets(i)
            End If
        Next j
    Next i
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
End Sub

