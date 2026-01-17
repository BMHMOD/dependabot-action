Sub FillYardsFromStock_Enhanced()
    '========================================================
    ' ماكرو محسّن لملء Internal & External Yard من STOCK
    ' مع إمكانية التخصيص والتعديل السهل
    '========================================================
    
    Dim wbStock As Workbook, wbInternal As Workbook, wbExternal As Workbook
    Dim wsStock As Worksheet, wsInternal As Worksheet, wsExternal As Worksheet
    Dim lastRow As Long, i As Long
    Dim startTime As Double
    
    On Error GoTo ErrorHandler
    startTime = Timer
    
    ' إيقاف التحديث لتسريع المعالجة
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    '========================================================
    ' 1. فتح الملفات
    '========================================================
    Set wbStock = Workbooks.Open(SelectFile("اختر ملف STOCK"))
    Set wsStock = wbStock.Sheets(1)
    
    Set wbInternal = Workbooks.Open(SelectFile("اختر ملف Internal Yard"))
    Set wsInternal = wbInternal.Sheets(1)
    
    Set wbExternal = Workbooks.Open(SelectFile("اختر ملف External Yard"))
    Set wsExternal = wbExternal.Sheets(1)
    
    '========================================================
    ' 2. مسح البيانات القديمة
    '========================================================
    wsInternal.Range("C6:G100").ClearContents ' توسيع النطاق للأمان
    wsExternal.Range("C6:F30").ClearContents
    
    '========================================================
    ' 3. إعداد المصفوفات والتعيينات
    '========================================================
    
    ' Internal Yard Blocks - كل block له 3 صفوف (import, export, transshipment/storage)
    Dim dictInternal As Object
    Set dictInternal = CreateObject("Scripting.Dictionary")
    dictInternal.Add "M", 6
    dictInternal.Add "A", 9
    dictInternal.Add "B", 12
    dictInternal.Add "C", 15
    dictInternal.Add "D", 18
    dictInternal.Add "H", 21
    dictInternal.Add "F", 24
    dictInternal.Add "Y777", 29
    dictInternal.Add "S22", 35
    dictInternal.Add "S003", 38
    dictInternal.Add "S666", 41
    dictInternal.Add "INSP", 44
    dictInternal.Add "S002", 47  ' إضافة blocks إضافية
    dictInternal.Add "S03", 50
    dictInternal.Add "S333", 53
    
    ' External Yard Areas - تعديل حسب احتياجك
    Dim dictExternal As Object
    Set dictExternal = CreateObject("Scripting.Dictionary")
    ' Format: "YardName" -> Array(StartRow, "Area1|Area2|Block1|Block2")
    dictExternal.Add "التجارية", Array(6, "S444|S068|S032")
    dictExternal.Add "المفروزة", Array(8, "S900|RORO1|BR")
    dictExternal.Add "68", Array(10, "S600|S700")
    dictExternal.Add "Other1", Array(12, "RAIL|SCALE")
    dictExternal.Add "Other2", Array(14, "XRAY|RORO5")
    
    '========================================================
    ' 4. معالجة البيانات من STOCK
    '========================================================
    lastRow = wsStock.Cells(wsStock.Rows.Count, "A").End(xlUp).Row
    
    Dim modeVal As String, blockVal As String, feVal As String
    Dim cntrLen As String, areaVal As String
    Dim progressCounter As Long
    
    ' عرض شريط التقدم
    Application.StatusBar = "جاري معالجة البيانات... 0%"
    
    For i = 2 To lastRow
        ' قراءة القيم
        modeVal = UCase(Trim(CStr(wsStock.Cells(i, 16).Value))) ' Column P: Mode
        blockVal = UCase(Trim(CStr(wsStock.Cells(i, 7).Value))) ' Column G: Block
        feVal = UCase(Trim(CStr(wsStock.Cells(i, 13).Value))) ' Column M: FE
        cntrLen = Trim(CStr(wsStock.Cells(i, 10).Value)) ' Column J: Cntr Len
        areaVal = UCase(Trim(CStr(wsStock.Cells(i, 6).Value))) ' Column F: Area
        
        ' تحديث شريط التقدم كل 500 صف
        progressCounter = progressCounter + 1
        If progressCounter Mod 500 = 0 Then
            Application.StatusBar = "جاري معالجة البيانات... " & _
                Format(progressCounter / lastRow, "0%")
        End If
        
        '========================================================
        ' معالجة Internal Yard
        '========================================================
        If dictInternal.exists(blockVal) Then
            Call UpdateYardCount(wsInternal, dictInternal(blockVal), _
                               modeVal, cntrLen, feVal, True)
        End If
        
        '========================================================
        ' معالجة External Yard
        '========================================================
        Dim yardKey As Variant
        For Each yardKey In dictExternal.Keys
            Dim yardData As Variant
            yardData = dictExternal(yardKey)
            
            Dim yardStartRow As Long
            Dim yardAreas As String
            yardStartRow = yardData(0)
            yardAreas = "|" & yardData(1) & "|"
            
            ' تحقق إذا كان Area أو Block ضمن هذه الساحة
            If InStr(1, yardAreas, "|" & areaVal & "|", vbTextCompare) > 0 Or _
               InStr(1, yardAreas, "|" & blockVal & "|", vbTextCompare) > 0 Then
                
                Call UpdateYardCount(wsExternal, yardStartRow, _
                                   modeVal, cntrLen, feVal, False)
                Exit For ' إيجاد الساحة، الخروج
            End If
        Next yardKey
    Next i
    
    '========================================================
    ' 5. حفظ وإنهاء
    '========================================================
    Application.StatusBar = "جاري حفظ الملفات..."
    
    wbInternal.Save
    wbExternal.Save
    
    ' إعادة تفعيل الإعدادات
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    
    Dim elapsed As Double
    elapsed = Timer - startTime
    
    MsgBox "تم الانتهاء بنجاح!" & vbCrLf & _
           "عدد الصفوف المعالجة: " & Format(lastRow - 1, "#,##0") & vbCrLf & _
           "الوقت المستغرق: " & Format(elapsed, "0.0") & " ثانية", _
           vbInformation, "اكتمل"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    MsgBox "حدث خطأ: " & Err.Description & vbCrLf & _
           "في السطر: " & Err.Source, vbCritical
End Sub

'========================================================
' Sub روتين لتحديث العدادات
'========================================================
Private Sub UpdateYardCount(ws As Worksheet, baseRow As Long, _
                           modeVal As String, cntrLen As String, _
                           feVal As String, isInternal As Boolean)
    
    Dim targetRow As Long
    Dim targetCol As Long
    
    ' تحديد الصف حسب Mode
    Select Case modeVal
        Case "IMPORT"
            targetRow = baseRow
        Case "EXPORT"
            targetRow = baseRow + 1
        Case "STORAGE", "TRANSSHIPMENT"
            targetRow = baseRow + 2
        Case Else
            Exit Sub
    End Select
    
    ' تحديد العمود حسب Container Length و Full/Empty
    If cntrLen = "20" And feVal = "F" Then
        targetCol = 3 ' Column C: 20F
    ElseIf cntrLen = "40" And feVal = "F" Then
        targetCol = 4 ' Column D: 40F
    ElseIf cntrLen = "20" And feVal = "E" Then
        targetCol = 5 ' Column E: 20E
    ElseIf cntrLen = "40" And feVal = "E" Then
        targetCol = 6 ' Column F: 40E
    ElseIf cntrLen = "45" And isInternal Then
        targetCol = 7 ' Column G: 45 (Internal only)
    Else
        Exit Sub
    End If
    
    ' زيادة العداد
    If IsEmpty(ws.Cells(targetRow, targetCol).Value) Then
        ws.Cells(targetRow, targetCol).Value = 1
    Else
        ws.Cells(targetRow, targetCol).Value = ws.Cells(targetRow, targetCol).Value + 1
    End If
End Sub

'========================================================
' Function لاختيار الملف
'========================================================
Private Function SelectFile(prompt As String) As String
    Dim filePath As Variant
    filePath = Application.GetOpenFilename("Excel Files (*.xlsx;*.xlsm), *.xlsx;*.xlsm", , prompt)
    If filePath = False Then
        MsgBox "لم يتم اختيار ملف. إلغاء العملية.", vbExclamation
        End
    End If
    SelectFile = CStr(filePath)
End Function
