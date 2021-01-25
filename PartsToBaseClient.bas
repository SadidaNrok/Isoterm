Attribute VB_Name = "PartsToBase"
Sub PartsToBase()
   Dim IsOpen As Boolean, WasOpened As Boolean
   Dim ActiveBook As Workbook
   Const CorePath$ = "\\SERVER3\Documents\Технологи\Иванов Константин\Microsoft Excel\ПКР\Core.xlsm"
   Const CoreName$ = "Core.xlsm"
   Const ModuleName = "!PartsToBase.PartsToBase"
   
   Set ActiveBook = ThisWorkbook
   IsOpen = IsCoreOpen(CoreName)
   Application.ScreenUpdating = False
   If Not IsOpen Then OpenCore (CorePath)
   ActiveBook.Activate
   Application.Run (CoreName + ModuleName)
   If Not IsOpen Then CloseCore (CoreName)
   Application.ScreenUpdating = True
End Sub
Function IsCoreOpen(CoreName As String) As Boolean
   Dim WBook As Workbook
   For Each WBook In Workbooks
       If WBook.Name = CoreName Then
           IsCoreOpen = True
           Exit Function
       End If
   Next WBook
   IsOpenCore = False
End Function
Sub OpenCore(CoreName As String)
   Workbooks.Open CoreName, UpdateLinks:=True
   WasOpened = True
End Sub
Sub CloseCore(CoreName As String)
   Workbooks(CoreName).Saved = True
   Workbooks(CoreName).Close
End Sub
