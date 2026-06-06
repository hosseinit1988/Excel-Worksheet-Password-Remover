Option Explicit

Sub FastPasswordBreaker_AllSheets()
    Dim ws As Worksheet
    Dim i As Integer, j As Integer, k As Integer
    Dim l As Integer, m As Integer, n As Integer
    Dim pwd As String
    Dim totalSheets As Integer, unlockedSheets As Integer
    Dim sheetIndex As Integer
    Dim startTime As Double, endTime As Double
    
    totalSheets = ActiveWorkbook.Worksheets.Count
    unlockedSheets = 0
    startTime = Timer
    
    ' اضافه کردن شیت برای گزارش
    Dim reportSheet As Worksheet
    On Error Resume Next
    Set reportSheet = ActiveWorkbook.Worksheets("PasswordReport")
    If reportSheet Is Nothing Then
        Set reportSheet = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
        reportSheet.Name = "PasswordReport"
        reportSheet.Range("A1") = "شماره"
        reportSheet.Range("B1") = "نام شیت"
        reportSheet.Range("C1") = "رمز پیدا شده"
        reportSheet.Range("D1") = "زمان (ثانیه)"
        reportSheet.Range("A1:D1").Font.Bold = True
    Else
        reportSheet.Cells.Clear
        reportSheet.Range("A1") = "شماره"
        reportSheet.Range("B1") = "نام شیت"
        reportSheet.Range("C1") = "رمز پیدا شده"
        reportSheet.Range("D1") = "زمان (ثانیه)"
        reportSheet.Range("A1:D1").Font.Bold = True
    End If
    On Error GoTo 0
    
    sheetIndex = 1
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> "PasswordReport" Then
            Application.StatusBar = "در حال پردازش: " & ws.Name & " (" & sheetIndex & "/" & totalSheets & ")"
            Dim sheetStartTime As Double
            sheetStartTime = Timer
            
            If ws.ProtectContents = True Then
                Dim found As Boolean
                Dim foundPwd As String
                found = False
                
                ' روش سریع: فقط ترکیب A, B به اضافه 4 کاراکتر انتهایی
                For i = 65 To 66  ' A, B
                    If found Then Exit For
                    For j = 65 To 66
                        If found Then Exit For
                        For k = 65 To 66
                            If found Then Exit For
                            For l = 65 To 66
                                If found Then Exit For
                                For m = 65 To 66
                                    If found Then Exit For
                                    For n = 65 To 66
                                        ' ترکیب با کاراکترهای رایج
                                        Dim c1 As Integer, c2 As Integer
                                        For c1 = 48 To 57  ' اعداد 0-9
                                            For c2 = 48 To 57
                                                pwd = Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(n) & Chr(c1) & Chr(c2)
                                                
                                                On Error Resume Next
                                                ws.Unprotect pwd
                                                On Error GoTo 0
                                                
                                                If ws.ProtectContents = False Then
                                                    found = True
                                                    foundPwd = pwd
                                                    unlockedSheets = unlockedSheets + 1
                                                    Exit For
                                                End If
                                            Next c2
                                            If found Then Exit For
                                        Next c1
                                        If found Then Exit For
                                    Next n
                                    If found Then Exit For
                                Next m
                                If found Then Exit For
                            Next l
                            If found Then Exit For
                        Next k
                        If found Then Exit For
                    Next j
                    If found Then Exit For
                Next i
                
                ' ذخیره در گزارش
                reportSheet.Cells(sheetIndex + 1, 1) = sheetIndex
                reportSheet.Cells(sheetIndex + 1, 2) = ws.Name
                If found Then
                    reportSheet.Cells(sheetIndex + 1, 3) = foundPwd
                    reportSheet.Cells(sheetIndex + 1, 3).Interior.Color = RGB(144, 238, 144)
                Else
                    reportSheet.Cells(sheetIndex + 1, 3) = "پیدا نشد"
                    reportSheet.Cells(sheetIndex + 1, 3).Interior.Color = RGB(255, 200, 200)
                End If
                reportSheet.Cells(sheetIndex + 1, 4) = Round(Timer - sheetStartTime, 2)
            Else
                ' شیت رمز ندارد
                unlockedSheets = unlockedSheets + 1
                reportSheet.Cells(sheetIndex + 1, 1) = sheetIndex
                reportSheet.Cells(sheetIndex + 1, 2) = ws.Name
                reportSheet.Cells(sheetIndex + 1, 3) = "بدون رمز"
                reportSheet.Cells(sheetIndex + 1, 4) = 0
            End If
            sheetIndex = sheetIndex + 1
        End If
    Next ws
    
    endTime = Timer
    Application.StatusBar = False
    
    ' تنظیم ظاهر گزارش
    reportSheet.Columns("A:D").AutoFit
    reportSheet.Activate
    
    MsgBox "✅ عملیات با موفقیت به پایان رسید!" & vbCrLf & _
           "─────────────────────" & vbCrLf & _
           "کل شیت‌ها: " & totalSheets & vbCrLf & _
           "شیت‌های باز شده: " & unlockedSheets & vbCrLf & _
           "شیت‌های قفل مانده: " & (totalSheets - unlockedSheets) & vbCrLf & _
           "─────────────────────" & vbCrLf & _
           "زمان کل: " & Round(endTime - startTime, 2) & " ثانیه" & vbCrLf & _
           "گزارش کامل در شیت 'PasswordReport' ذخیره شد.", vbInformation
End Sub
