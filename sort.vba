Sub sort()

Sheets("工作表2").Select

Dim old_day, new_day
Dim str
str = Split(Cells(6, "A").Value, ".")
For i = 0 To UBound(str)
    old_day = old_day & str(i)
Next

Application.ScreenUpdating = False
Sheets("工作表1").Select
V = Range("A4").End(xlDown).Row
Set D_N = CreateObject("Scripting.Dictionary")

Dim Day()
Dim Note()

ReDim Day(V - 2)
ReDim Note(V - 2)

For i = 0 To V - 2
    D_N.Add Cells(i + 4, "B"), Cells(i + 4, "E")
Next

Sheets("工作表2").Select

For Each a In D_N '摘要中無所得稅法的字樣，就刪除該項資料
    If InStr(1, D_N(a).Value, "所得稅法") = 0 Then
        D_N.Remove (a)
    End If
Next

Day() = D_N.keys
Note() = D_N.items

For t = 0 To UBound(Day)
    Cells(t + 6, "A") = Day(t)
    Cells(t + 6, "B") = Note(t)
Next

Range("A1").Select
Application.ScreenUpdating = True

str = Split(Day(0).Value, ".")
For i = 0 To UBound(str)
    new_day = new_day & str(i)
Next


If Val(new_day) > Val(old_day) Then
    MsgBox "法規資料有更動", vbOKOnly, "通知"
Else
End If

MsgBox "查詢完畢", vbOKOnly, "通知"

End Sub
