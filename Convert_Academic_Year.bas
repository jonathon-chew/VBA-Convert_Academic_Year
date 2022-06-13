Attribute VB_Name = "findacademicYear"
Sub findacademicYear()

lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row

columnToConvert = 9

For x = 2 To lastRow
If Cells(x, columnToConvert) >= DateValue("1/8/2017") And Cells(x, columnToConvert) <= DateValue("31/7/2018") Then
    Cells(x, columnToConvert) = "2017/2018"
    GoTo NextStudent
ElseIf Cells(x, columnToConvert) >= DateValue("1/8/2018") And Cells(x, columnToConvert) <= DateValue("31/7/2019") Then
        Cells(x, columnToConvert) = "2018/2019"
        GoTo NextStudent
ElseIf Cells(x, columnToConvert) >= DateValue("1/8/2016") And Cells(x, columnToConvert) <= DateValue("31/7/2017") Then
        Cells(x, columnToConvert) = "2016/2017"
        GoTo NextStudent
ElseIf Cells(x, columnToConvert) >= DateValue("1/8/2019") And Cells(x, columnToConvert) <= DateValue("31/7/2020") Then
        Cells(x, columnToConvert) = "2019/2020"
        GoTo NextStudent
ElseIf Cells(x, columnToConvert) >= DateValue("1/8/2020") And Cells(x, columnToConvert) <= DateValue("31/7/2021") Then
        Cells(x, columnToConvert) = "2020/2021"
        GoTo NextStudent
ElseIf Cells(x, columnToConvert) >= DateValue("1/8/2021") And Cells(x, columnToConvert) <= DateValue("31/7/2022") Then
    Cells(x, columnToConvert) = "2021/2022"
    GoTo NextStudent
End If

NextStudent:

Next x


End Sub

