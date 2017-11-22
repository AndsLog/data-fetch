Public Sub gogo()
Application.ScreenUpdating = False
Sheets("工作表1").Select

Range("A1").Select
Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
Selection.ClearContents

    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;http://law.moj.gov.tw/News/news_result.aspx?SearchRange=G&k1=%E6%89%80%E5%BE%97%E7%A8%85%E6%B3%95" _
        , Destination:=Range("$A$1"))
        .WebSelectionType = xlAllTables
        .WebFormatting = xlWebFormattingNone
        .Refresh BackgroundQuery:=False
    End With
    
Sheets("工作表2").Select

Range("A1").Select

Application.ScreenUpdating = True

MsgBox "工作表1載入完成", vbOKOnly, "通知"

End Sub
