Sub TestSimpleCount()
    '========================================================
    ' نسخة تجريبية بسيطة لاختبار المنطق فقط
    ' تعرض النتائج في MsgBox بدون ملء الملفات
    '========================================================
    
    Dim wbStock As Workbook
    Dim wsStock As Worksheet
    Dim lastRow As Long, i As Long
    
    ' فتح ملف STOCK
    Dim stockPath As Variant
    stockPath = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx", , "اختر ملف STOCK للاختبار")
    If stockPath = False Then Exit Sub
    
    Set wbStock = Workbooks.Open(stockPath)
    Set wsStock = wbStock.Sheets(1)
    
    lastRow = wsStock.Cells(wsStock.Rows.Count, "A").End(xlUp).Row
    
    ' عدادات للاختبار - Block M فقط
    Dim m_import_20F As Long, m_import_40F As Long
    Dim m_export_20F As Long, m_export_40F As Long
    Dim m_storage_20F As Long, m_storage_40F As Long
    
    ' عدادات للاختبار - External Yard (S444 area)
    Dim ext_import_20F As Long, ext_import_40F As Long
    Dim ext_export_20F As Long, ext_export_40F As Long
    
    Application.ScreenUpdating = False
    
    ' قراءة البيانات
    For i = 2 To lastRow
        Dim modeVal As String, blockVal As String, feVal As String
        Dim cntrLen As String, areaVal As String
        
        modeVal = UCase(Trim(CStr(wsStock.Cells(i, 16).Value))) ' Mode
        blockVal = UCase(Trim(CStr(wsStock.Cells(i, 7).Value))) ' Block
        feVal = UCase(Trim(CStr(wsStock.Cells(i, 13).Value))) ' FE
        cntrLen = Trim(CStr(wsStock.Cells(i, 10).Value)) ' Cntr Len
        areaVal = UCase(Trim(CStr(wsStock.Cells(i, 6).Value))) ' Area
        
        ' اختبار Block M
        If blockVal = "M" And feVal = "F" Then
            If modeVal = "IMPORT" Then
                If cntrLen = "20" Then m_import_20F = m_import_20F + 1
                If cntrLen = "40" Then m_import_40F = m_import_40F + 1
            ElseIf modeVal = "EXPORT" Then
                If cntrLen = "20" Then m_export_20F = m_export_20F + 1
                If cntrLen = "40" Then m_export_40F = m_export_40F + 1
            ElseIf modeVal = "STORAGE" Then
                If cntrLen = "20" Then m_storage_20F = m_storage_20F + 1
                If cntrLen = "40" Then m_storage_40F = m_storage_40F + 1
            End If
        End If
        
        ' اختبار External Yard - S444 area
        If areaVal = "S444" And feVal = "F" Then
            If modeVal = "IMPORT" Then
                If cntrLen = "20" Then ext_import_20F = ext_import_20F + 1
                If cntrLen = "40" Then ext_import_40F = ext_import_40F + 1
            ElseIf modeVal = "EXPORT" Then
                If cntrLen = "20" Then ext_export_20F = ext_export_20F + 1
                If cntrLen = "40" Then ext_export_40F = ext_export_40F + 1
            End If
        End If
    Next i
    
    Application.ScreenUpdating = True
    
    ' عرض النتائج
    Dim results As String
    results = "نتائج الاختبار من ملف STOCK" & vbCrLf & String(40, "=") & vbCrLf & vbCrLf
    
    results = results & "INTERNAL YARD - Block M:" & vbCrLf
    results = results & "  Import:  20F=" & m_import_20F & ", 40F=" & m_import_40F & vbCrLf
    results = results & "  Export:  20F=" & m_export_20F & ", 40F=" & m_export_40F & vbCrLf
    results = results & "  Storage: 20F=" & m_storage_20F & ", 40F=" & m_storage_40F & vbCrLf & vbCrLf
    
    results = results & "EXTERNAL YARD - Area S444:" & vbCrLf
    results = results & "  Import: 20F=" & ext_import_20F & ", 40F=" & ext_import_40F & vbCrLf
    results = results & "  Export: 20F=" & ext_export_20F & ", 40F=" & ext_export_40F & vbCrLf & vbCrLf
    
    results = results & "إجمالي الصفوف المعالجة: " & Format(lastRow - 1, "#,##0")
    
    MsgBox results, vbInformation, "نتائج الاختبار"
    
    wbStock.Close False
End Sub
