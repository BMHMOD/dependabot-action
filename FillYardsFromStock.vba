Sub FillYardsFromStock()
    '========================================================
    ' هذا الماكرو يقوم بملء ملفات Internal Yard و External Yard
    ' من بيانات ملف STOCK
    '========================================================
    
    Dim wbStock As Workbook
    Dim wbInternal As Workbook
    Dim wbExternal As Workbook
    Dim wsStock As Worksheet
    Dim wsInternal As Worksheet
    Dim wsExternal As Worksheet
    
    Dim lastRow As Long
    Dim i As Long
    Dim modeVal As String
    Dim blockVal As String
    Dim feVal As String
    Dim cntrLen As String
    Dim areaVal As String
    
    ' متغيرات للعد
    Dim count20F As Long, count40F As Long, count20E As Long, count40E As Long, count45 As Long
    
    On Error GoTo ErrorHandler
    
    ' فتح الملفات - يجب تعديل المسارات حسب موقع ملفاتك
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    MsgBox "الرجاء اختيار ملف STOCK", vbInformation
    Dim stockPath As Variant
    stockPath = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx", , "اختر ملف STOCK")
    If stockPath = False Then Exit Sub
    Set wbStock = Workbooks.Open(stockPath)
    Set wsStock = wbStock.Sheets(1)
    
    MsgBox "الرجاء اختيار ملف Internal Yard", vbInformation
    Dim internalPath As Variant
    internalPath = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx", , "اختر ملف Internal Yard")
    If internalPath = False Then
        wbStock.Close False
        Exit Sub
    End If
    Set wbInternal = Workbooks.Open(internalPath)
    Set wsInternal = wbInternal.Sheets(1)
    
    MsgBox "الرجاء اختيار ملف External Yard", vbInformation
    Dim externalPath As Variant
    externalPath = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx", , "اختر ملف External Yard")
    If externalPath = False Then
        wbStock.Close False
        wbInternal.Close False
        Exit Sub
    End If
    Set wbExternal = Workbooks.Open(externalPath)
    Set wsExternal = wbExternal.Sheets(1)
    
    ' مسح البيانات القديمة من Internal Yard (الأعمدة C:G من الصف 6)
    ' مسح Internal Yard
    wsInternal.Range("C6:G52").ClearContents
    
    ' مسح External Yard (الأعمدة C:F من الصف 6)
    wsExternal.Range("C6:F15").ClearContents
    
    '========================================================
    ' معالجة Internal Yard
    '========================================================
    lastRow = wsStock.Cells(wsStock.Rows.Count, "A").End(xlUp).Row
    
    ' تعريف مصفوفة للBlocks في Internal Yard
    ' الصف 6: M (import), 7: M (export), 8: M (transshipment)
    ' الصف 9: A (import), 10: A (export), 11: A (transshipment)
    ' الصف 12: B (import), 13: B (export), 14: B (transshipment)
    ' الصف 15: C (import), 16: C (export), 17: C (transshipment)
    ' الصف 18: D (import), 19: D (export), 20: D (transshipment)
    ' الصف 21: H (import), 22: H (export), 23: H (transshipment)
    ' الصف 24: F (import), 25: F (export/ref), 26: F (transshipment)
    ' والباقي حسب الترتيب
    
    Dim blockRows As Object
    Set blockRows = CreateObject("Scripting.Dictionary")
    
    ' Internal Yard Blocks mapping (Block -> StartRow)
    blockRows.Add "M", 6
    blockRows.Add "A", 9
    blockRows.Add "B", 12
    blockRows.Add "C", 15
    blockRows.Add "D", 18
    blockRows.Add "H", 21
    blockRows.Add "F", 24
    blockRows.Add "Y777", 29  ' حسب البنية
    blockRows.Add "S22", 35
    blockRows.Add "S003", 38
    blockRows.Add "S666", 41
    blockRows.Add "INSP", 44
    
    ' قراءة البيانات من STOCK ومعالجتها
    For i = 2 To lastRow ' البدء من الصف 2 (بعد الهيدر)
        modeVal = Trim(UCase(wsStock.Cells(i, "P").Value)) ' Mode
        blockVal = Trim(UCase(wsStock.Cells(i, "G").Value)) ' Block
        feVal = Trim(UCase(wsStock.Cells(i, "M").Value)) ' FE (F or E)
        cntrLen = Trim(wsStock.Cells(i, "J").Value) ' Cntr Len
        areaVal = Trim(UCase(wsStock.Cells(i, "F").Value)) ' Area
        
        ' تحقق من أن Block موجود في Internal Yard
        If blockRows.exists(blockVal) And blockVal <> "" Then
            Dim baseRow As Long
            baseRow = blockRows(blockVal)
            
            ' تحديد الصف حسب Mode
            Dim targetRow As Long
            If modeVal = "IMPORT" Then
                targetRow = baseRow
            ElseIf modeVal = "EXPORT" Then
                targetRow = baseRow + 1
            ElseIf modeVal = "STORAGE" Then
                targetRow = baseRow + 2 ' transshipment row
            Else
                targetRow = 0
            End If
            
            If targetRow > 0 Then
                ' تحديد العمود حسب FE و Cntr Len
                If cntrLen = "20" And feVal = "F" Then
                    wsInternal.Cells(targetRow, 3).Value = wsInternal.Cells(targetRow, 3).Value + 1 ' Column C: 20F
                ElseIf cntrLen = "40" And feVal = "F" Then
                    wsInternal.Cells(targetRow, 4).Value = wsInternal.Cells(targetRow, 4).Value + 1 ' Column D: 40F
                ElseIf cntrLen = "20" And feVal = "E" Then
                    wsInternal.Cells(targetRow, 5).Value = wsInternal.Cells(targetRow, 5).Value + 1 ' Column E: 20E
                ElseIf cntrLen = "40" And feVal = "E" Then
                    wsInternal.Cells(targetRow, 6).Value = wsInternal.Cells(targetRow, 6).Value + 1 ' Column F: 40E
                ElseIf cntrLen = "45" Then
                    wsInternal.Cells(targetRow, 7).Value = wsInternal.Cells(targetRow, 7).Value + 1 ' Column G: 45
                End If
            End If
        End If
    Next i
    
    '========================================================
    ' معالجة External Yard
    '========================================================
    ' External Yards حسب Area أو Block معين
    ' الصف 6-7: ساحة التجارية (S444, S068, S032)
    ' الصف 8-9: ساحة المفروزة
    ' الصف 10-11: ساحة 68
    ' وهكذا...
    
    ' مصفوفة للساحات الخارجية
    Dim externalYards As Object
    Set externalYards = CreateObject("Scripting.Dictionary")
    
    ' External Yard mapping (YardName -> StartRow, AreaList)
    ' يجب تعديل هذا حسب الساحات الفعلية في ملفك
    externalYards.Add "التجارية", Array(6, "S444,S068,S032")
    externalYards.Add "المفروزة", Array(8, "S900,RORO1")
    externalYards.Add "68", Array(10, "S600")
    ' أضف المزيد حسب احتياجك
    
    ' إعادة قراءة البيانات للساحات الخارجية
    For i = 2 To lastRow
        modeVal = Trim(UCase(wsStock.Cells(i, "P").Value)) ' Mode
        blockVal = Trim(UCase(wsStock.Cells(i, "G").Value)) ' Block
        feVal = Trim(UCase(wsStock.Cells(i, "M").Value)) ' FE
        cntrLen = Trim(wsStock.Cells(i, "J").Value) ' Cntr Len
        areaVal = Trim(UCase(wsStock.Cells(i, "F").Value)) ' Area
        
        ' تحقق من الساحات الخارجية
        Dim yardKey As Variant
        For Each yardKey In externalYards.Keys
            Dim yardInfo As Variant
            yardInfo = externalYards(yardKey)
            Dim startRow As Long
            Dim areaList As String
            startRow = yardInfo(0)
            areaList = yardInfo(1)
            
            ' تحقق إذا كان Area أو Block ضمن هذه الساحة
            If InStr(1, areaList, areaVal, vbTextCompare) > 0 Or _
               InStr(1, areaList, blockVal, vbTextCompare) > 0 Then
                
                ' تحديد الصف حسب Mode
                Dim extTargetRow As Long
                If modeVal = "IMPORT" Then
                    extTargetRow = startRow
                ElseIf modeVal = "EXPORT" Then
                    extTargetRow = startRow + 1
                Else
                    extTargetRow = 0
                End If
                
                If extTargetRow > 0 Then
                    ' تحديد العمود
                    If cntrLen = "20" And feVal = "F" Then
                        wsExternal.Cells(extTargetRow, 3).Value = wsExternal.Cells(extTargetRow, 3).Value + 1 ' Column C: 20F
                    ElseIf cntrLen = "40" And feVal = "F" Then
                        wsExternal.Cells(extTargetRow, 4).Value = wsExternal.Cells(extTargetRow, 4).Value + 1 ' Column D: 40F
                    ElseIf cntrLen = "20" And feVal = "E" Then
                        wsExternal.Cells(extTargetRow, 5).Value = wsExternal.Cells(extTargetRow, 5).Value + 1 ' Column E: 20E
                    ElseIf cntrLen = "40" And feVal = "E" Then
                        wsExternal.Cells(extTargetRow, 6).Value = wsExternal.Cells(extTargetRow, 6).Value + 1 ' Column F: 40E
                    End If
                End If
            End If
        Next yardKey
    Next i
    
    ' حفظ الملفات
    wbInternal.Save
    wbExternal.Save
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "تم ملء الملفات بنجاح!", vbInformation, "اكتمل"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "حدث خطأ: " & Err.Description, vbCritical
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
