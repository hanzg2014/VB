### 東京気温変動分析 1962 ~ 2012

#### VBA程序

```
Rem Whether a year is a lunar year or not
Function IsLunarYear(y As Integer) As Boolean
    If (y Mod 4 = 0 Or y Mod 400 = 0) And y Mod 100 <> 0 Then
        IsLunarYear = True
    Else
        IsLunarYear = False
    End If
End Function

Rem Yearly average temperature
Sub AverageTemperature()
    Worksheets(1).Activate
    Dim i As Integer
    Dim j As Integer
    Dim days As Integer
    Dim sum As Single
    Dim avg As Single

    Rem year 1962 ~ 2012
    For i = 0 To 50
        j = i * 12 + 3
        Rem sum of all the temperatures except that of Feb.
        sum_31 = (Range("D" & j).Value + Range("D" & (j + 2)).Value + Range("D" & (j + 4)).Value + Range("D" & (j + 6)).Value + Range("D" & (j + 7)).Value + Range("D" & (j + 9)).Value + Range("D" & (j + 11)).Value) * 31
        sum_30 = (Range("D" & (j + 3)).Value + Range("D" & (j + 5)).Value + Range("D" & (j + 8)).Value + Range("D" & (j + 10)).Value) * 30
        Rem if this year is a lunar year
        If IsLunarYear(Range("L" & j).Value) Then
            sum = sum_30 + sum_31 + Range("D" & (j + 1)).Value * 29
            days = 366
            agg = sum / days
        Else
            sum = sum_30 + sum_31 + Range("D" & (j + 1)).Value * 28
            days = 365
            agg = sum / days
        End If

        Range("M" & (i + 3)).Value = agg
        
        Rem Range("N" & (i + 3)).Formula = "=SUM(G" & index_begin & ": G" & (index_begin + 12) & ")"
    Next i
End Sub


Rem Total amount of the precipitations by year
Sub Precipitation()
    Worksheets(1).Activate
    Dim i As Integer
    Rem 1962 ~ 2012
    For i = 0 To 50
        index_begin = i * 12 + 3
        Range("N" & (i + 3)).Formula = "=SUM(G" & index_begin & ": G" & (index_begin + 12) & ")"
    Next i
End Sub

Rem Highest temperature by year
Sub TemperatureHigh()
    Worksheets(1).Activate
    Dim i As Integer
    Rem 1962 ~ 2012
    For i = 0 To 50
        index_begin = i * 12 + 3
        Range("P" & (i + 3)).Formula = "=MAX(E" & index_begin & ": E" & (index_begin + 12) & ")"
    Next i
End Sub


Rem Lowest temperature by year
Sub TemperatureLow()
    Worksheets(1).Activate
    Dim i As Integer
    Rem 1962 ~ 2012
    For i = 0 To 50
        index_begin = i * 12 + 3
        Range("Q" & (i + 3)).Formula = "=MIN(F" & index_begin & ": F" & (index_begin + 12) & ")"
    Next i
    
End Sub


```

---

#### 參考

Excel VBA 程式設計教學：函數（Function）與子程序（Sub）
https://blog.gtwang.org/programming/excel-vba-function-and-sub/2/

6小時，寫了一篇適合Excel小白學的VBA入門教程 
https://www.zixundingzhi.com/Excel/940aa5f9b7e09704.html

Getting Started with VBA in Office
https://msdn.microsoft.com/VBA/office-shared-vba/articles/getting-started-with-vba-in-office
'''&amp; i 改成 & i 就好了'''
