Attribute VB_Name = "Module_ComplianceStock"
' ================================================================
' COMPLIANCE STOCK - Professional Dashboard VBA
' PT. Gajah Tunggal Tbk | PPC Department
' ================================================================
' FITUR:
'   1. Dashboard ringkasan status stok (kritis/warning/aman)
'   2. Auto-highlight sel sisa jam berdasarkan threshold
'   3. Filter cepat per mesin (RTBA1, RTBA2, dst)
'   4. Export PDF per kelompok mesin
'   5. Auto-refresh warna setiap buka file
'   6. Tombol navigasi antar sheet material
'   7. Summary popup per ban (double-click)
' ================================================================

Option Explicit

' ── Konstanta warna tema GT ──────────────────────────────────────
Private Const CLR_KRITIS   As Long = 16711680   ' Merah     (#FF0000) < 8 jam
Private Const CLR_WARNING  As Long = 16763904   ' Oranye    (#FFC000) 8-16 jam
Private Const CLR_AMAN     As Long = 5296274    ' Hijau tua (#50C878) 16-24 jam
Private Const CLR_HEADER   As Long = 1644912    ' Biru GT   (#191970) header utama
Private Const CLR_SUBHDR   As Long = 4210752    ' Abu gelap (#404040)
Private Const CLR_WHITE    As Long = 16777215   ' Putih
Private Const CLR_LTBLUE   As Long = 13434828   ' Biru muda  background highlight
Private Const CLR_GOLD     As Long = 16763904   ' Emas untuk prioritas

' ── Nama sheet utama ────────────────────────────────────────────
Private Const SHEET_MAIN   As String = "Print Schedule Building"

' ── Baris data mulai ────────────────────────────────────────────
Private Const ROW_START    As Long = 17
Private Const ROWS_PER_SKU As Long = 3

' ── Kolom jam material (Q=17 s/d AD=30) ─────────────────────────
Private Const COL_JAM_START As Long = 17   ' Q - Tread
Private Const COL_JAM_END   As Long = 30   ' AD - Steel Chafer kanan

' ================================================================
' 1. ENTRY POINT — dipanggil dari tombol di sheet
' ================================================================

Sub RefreshDashboard()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_MAIN)

    Call ApplyHeaderFormat(ws)
    Call ColorizeJamCells(ws)
    Call BuildSummaryPanel(ws)
    Call HighlightPrioritasRows(ws)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Dashboard diperbarui! " & Now(), vbInformation, "Compliance Stock GT"
End Sub

Sub FilterByMesin()
    ' Tampilkan inputbox untuk filter mesin
    Dim mesin As String
    mesin = InputBox("Masukkan kode mesin (contoh: RTBA1, RTBA2, RTBB1, dll)" & vbCrLf & _
                     "Kosongkan untuk tampilkan semua:", "Filter Mesin", "")

    Application.ScreenUpdating = False
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_MAIN)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    Dim r As Long
    For r = ROW_START To lastRow
        Dim colAL As String
        colAL = Trim(ws.Cells(r, 38).Value) ' Kolom AL = kode mesin building

        If mesin = "" Then
            ws.Rows(r).Hidden = False
        ElseIf InStr(1, UCase(colAL), UCase(mesin)) > 0 Then
            ws.Rows(r).Hidden = False
        ElseIf ws.Cells(r, 38).Value = "" Then
            ' Baris sub (stock/jam) ikut parent
            ws.Rows(r).Hidden = False
        Else
            ws.Rows(r).Hidden = True
        End If
    Next r

    Application.ScreenUpdating = True
    If mesin = "" Then
        MsgBox "Semua baris ditampilkan.", vbInformation, "Filter"
    Else
        MsgBox "Filter mesin: " & UCase(mesin) & " diterapkan.", vbInformation, "Filter"
    End If
End Sub

Sub ExportPDF()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_MAIN)

    Dim savePath As String
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:="ComplianceStock_" & Format(Now, "yyyymmdd_HHmm"), _
        FileFilter:="PDF Files (*.pdf), *.pdf", _
        Title:="Simpan sebagai PDF")

    If savePath = "False" Then Exit Sub

    ws.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=savePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True

    MsgBox "PDF berhasil disimpan ke:" & vbCrLf & savePath, vbInformation, "Export PDF"
End Sub

Sub ShowKritisOnly()
    ' Tampilkan hanya baris yang ada material kritis (< 8 jam)
    Application.ScreenUpdating = False
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_MAIN)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    Dim r As Long, hasKritis As Boolean
    Dim blockStart As Long

    r = ROW_START
    Do While r <= lastRow
        hasKritis = False
        ' Baris ke-2 dari blok 3 baris = baris jam (offset +1 dari baris utama)
        Dim rJam As Long
        rJam = r + 1

        If rJam <= lastRow Then
            Dim c As Long
            For c = COL_JAM_START To COL_JAM_END
                Dim cellVal As String
                cellVal = Trim(CStr(ws.Cells(rJam, c).Value))
                If cellVal <> "-" And cellVal <> "" Then
                    Dim jamVal As Double
                    If IsNumeric(Replace(cellVal, " Jam", "")) Then
                        jamVal = CDbl(Replace(cellVal, " Jam", ""))
                        If jamVal < 8 Then
                            hasKritis = True
                            Exit For
                        End If
                    End If
                End If
            Next c
        End If

        ' Sembunyikan atau tampilkan 3 baris blok
        Dim i As Long
        For i = 0 To ROWS_PER_SKU - 1
            If r + i <= lastRow Then
                ws.Rows(r + i).Hidden = Not hasKritis
            End If
        Next i

        r = r + ROWS_PER_SKU
    Loop

    Application.ScreenUpdating = True
    MsgBox "Menampilkan baris dengan material KRITIS (< 8 jam).", vbInformation, "Filter Kritis"
End Sub

Sub ShowAllRows()
    Application.ScreenUpdating = False
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_MAIN)
    ws.Rows.Hidden = False
    Application.ScreenUpdating = True
    MsgBox "Semua baris ditampilkan.", vbInformation, "Reset Filter"
End Sub

' ================================================================
' 2. COLORIZE JAM CELLS
' ================================================================

Sub ColorizeJamCells(ws As Worksheet)
    ' Warnai sel sisa jam sesuai threshold
    ' < 8 jam  = MERAH   (kritis)
    ' 8-16 jam = ORANYE  (warning)
    ' >16 jam  = HIJAU   (aman)
    ' "-"      = abu-abu (tidak ada material)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    Dim r As Long, c As Long
    For r = ROW_START To lastRow
        ' Baris jam = setiap baris ke-2 dalam blok 3
        ' Baris ke-1 = header/SKU, ke-2 = stok, ke-3 = jam
        ' Berdasarkan struktur excel: baris jam = offset ke-2 dari baris SKU
        Dim rowType As Long
        rowType = ((r - ROW_START) Mod ROWS_PER_SKU)

        If rowType = 1 Then ' Baris ke-2 = baris jam (0-indexed: 0=SKU, 1=jam, 2=stok? cek lagi)
            For c = COL_JAM_START To COL_JAM_END
                Dim cell As Range
                Set cell = ws.Cells(r, c)
                Call ApplyJamColor(cell)
            Next c
        End If
    Next r
End Sub

Private Sub ApplyJamColor(cell As Range)
    Dim txt As String
    txt = Trim(CStr(cell.Value))

    If txt = "-" Or txt = "" Or txt = "0" Then
        ' Tidak ada material — background abu muda
        cell.Interior.Color = 14671839  ' #DFDFDF
        cell.Font.Color = 8421504       ' Abu medium
        cell.Font.Bold = False
        Exit Sub
    End If

    ' Ekstrak angka dari "21.1 Jam" atau angka biasa
    Dim numStr As String
    numStr = Replace(txt, " Jam", "")
    numStr = Replace(numStr, ",", ".")

    If Not IsNumeric(numStr) Then
        cell.Interior.ColorIndex = xlNone
        Exit Sub
    End If

    Dim jamVal As Double
    jamVal = CDbl(numStr)

    Select Case True
        Case jamVal < 8
            ' KRITIS - Merah terang + teks putih bold
            cell.Interior.Color = CLR_KRITIS
            cell.Font.Color = CLR_WHITE
            cell.Font.Bold = True
        Case jamVal < 16
            ' WARNING - Oranye + teks hitam bold
            cell.Interior.Color = CLR_WARNING
            cell.Font.Color = 0
            cell.Font.Bold = True
        Case Else
            ' AMAN - Hijau + teks putih
            cell.Interior.Color = CLR_AMAN
            cell.Font.Color = CLR_WHITE
            cell.Font.Bold = False
    End Select
End Sub

' ================================================================
' 3. HEADER FORMAT
' ================================================================

Private Sub ApplyHeaderFormat(ws As Worksheet)
    ' Row 14 = Header utama
    With ws.Rows(14)
        .Interior.Color = CLR_HEADER
        .Font.Color = CLR_WHITE
        .Font.Bold = True
        .Font.Size = 10
        .Font.Name = "Arial"
        .RowHeight = 30
    End With

    ' Row 15 = Sub-header curing/building
    With ws.Rows(15)
        .Interior.Color = CLR_SUBHDR
        .Font.Color = CLR_WHITE
        .Font.Bold = True
        .Font.Size = 9
        .Font.Name = "Arial"
    End With

    ' Row 16 = Sub-header shift
    With ws.Rows(16)
        .Interior.Color = 7829367   ' Biru medium
        .Font.Color = CLR_WHITE
        .Font.Bold = True
        .Font.Size = 8
    End With
End Sub

' ================================================================
' 4. HIGHLIGHT BARIS PRIORITAS
' ================================================================

Private Sub HighlightPrioritasRows(ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    Dim r As Long
    For r = ROW_START To lastRow Step ROWS_PER_SKU
        Dim statusCell As Range
        Set statusCell = ws.Cells(r, 16) ' Kolom P = Status

        Dim statusVal As String
        statusVal = LCase(Trim(CStr(statusCell.Value)))

        Select Case statusVal
            Case "prioritas"
                ' Background kuning muda untuk seluruh baris utama
                ws.Range(ws.Cells(r, 1), ws.Cells(r, 16)).Interior.Color = 16776960 ' Kuning
                ws.Range(ws.Cells(r, 1), ws.Cells(r, 16)).Font.Color = 0
                ws.Range(ws.Cells(r, 1), ws.Cells(r, 16)).Font.Bold = True
            Case "bo"
                ' BO = Back Order — highlight merah muda
                ws.Range(ws.Cells(r, 1), ws.Cells(r, 16)).Interior.Color = 16751001 ' Merah muda
                ws.Range(ws.Cells(r, 1), ws.Cells(r, 16)).Font.Color = 0
            Case Else
                ' Normal — baris bergantian putih/abu sangat muda
                If ((r - ROW_START) / ROWS_PER_SKU Mod 2) = 0 Then
                    ws.Range(ws.Cells(r, 1), ws.Cells(r, 16)).Interior.Color = 16777215 ' Putih
                Else
                    ws.Range(ws.Cells(r, 1), ws.Cells(r, 16)).Interior.Color = 15921906 ' Abu sangat muda
                End If
        End Select
    Next r
End Sub

' ================================================================
' 5. SUMMARY PANEL — ringkasan jumlah kritis/warning/aman
' ================================================================

Private Sub BuildSummaryPanel(ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    Dim cntKritis As Long, cntWarning As Long, cntAman As Long, cntTotal As Long

    Dim r As Long, c As Long
    For r = ROW_START To lastRow
        Dim rowType As Long
        rowType = ((r - ROW_START) Mod ROWS_PER_SKU)
        If rowType = 1 Then ' Baris jam
            For c = COL_JAM_START To COL_JAM_END
                Dim txt As String
                txt = Replace(Trim(CStr(ws.Cells(r, c).Value)), " Jam", "")
                If IsNumeric(txt) Then
                    Dim v As Double
                    v = CDbl(txt)
                    cntTotal = cntTotal + 1
                    If v < 8 Then
                        cntKritis = cntKritis + 1
                    ElseIf v < 16 Then
                        cntWarning = cntWarning + 1
                    Else
                        cntAman = cntAman + 1
                    End If
                End If
            Next c
        End If
    Next r

    ' Tulis summary di area AH1:AJ6 (sudah ada di kanan sheet)
    With ws.Cells(1, 37)
        .Value = "SUMMARY STOCK"
        .Font.Bold = True
        .Font.Color = CLR_WHITE
        .Interior.Color = CLR_HEADER
    End With
    ws.Cells(2, 37).Value = "Kritis (<8 jam)"
    ws.Cells(2, 38).Value = cntKritis
    ws.Cells(2, 37).Interior.Color = CLR_KRITIS
    ws.Cells(2, 38).Interior.Color = CLR_KRITIS
    ws.Cells(2, 37).Font.Color = CLR_WHITE
    ws.Cells(2, 38).Font.Color = CLR_WHITE
    ws.Cells(2, 38).Font.Bold = True

    ws.Cells(3, 37).Value = "Warning (8-16 jam)"
    ws.Cells(3, 38).Value = cntWarning
    ws.Cells(3, 37).Interior.Color = CLR_WARNING
    ws.Cells(3, 38).Interior.Color = CLR_WARNING

    ws.Cells(4, 37).Value = "Aman (>16 jam)"
    ws.Cells(4, 38).Value = cntAman
    ws.Cells(4, 37).Interior.Color = CLR_AMAN
    ws.Cells(4, 38).Interior.Color = CLR_AMAN
    ws.Cells(4, 37).Font.Color = CLR_WHITE
    ws.Cells(4, 38).Font.Color = CLR_WHITE

    ws.Cells(5, 37).Value = "TOTAL ITEM"
    ws.Cells(5, 38).Value = cntTotal
    ws.Cells(5, 37).Font.Bold = True
End Sub

' ================================================================
' 6. POPUP DETAIL PER BARIS (double-click)
' ================================================================

Sub ShowDetailPopup(r As Long)
    ' Dipanggil dari event double-click sheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_MAIN)

    ' Snap ke baris awal blok
    Dim blockRow As Long
    blockRow = ROW_START + (Int((r - ROW_START) / ROWS_PER_SKU) * ROWS_PER_SKU)

    Dim sku  As String: sku  = ws.Cells(blockRow, 2).Value
    Dim size As String: size = ws.Cells(blockRow, 3).Value
    Dim qcur As String: qcur = ws.Cells(blockRow, 4).Value
    Dim stat As String: stat = ws.Cells(blockRow, 16).Value

    Dim rJam As Long: rJam = blockRow + 1

    Dim matNames(13) As String
    matNames(0)  = "Tread"
    matNames(1)  = "Sidewall"
    matNames(2)  = "BEC"
    matNames(3)  = "Bead Finish"
    matNames(4)  = "Apex"
    matNames(5)  = "Body Ply"
    matNames(6)  = "Inner Liner"
    matNames(7)  = "Belt 1"
    matNames(8)  = "Belt 2"
    matNames(9)  = "Belt 3"
    matNames(10) = "Belt 4"
    matNames(11) = "Zero Belt"
    matNames(12) = "St.Chafer L"
    matNames(13) = "St.Chafer R"

    Dim msg As String
    msg = "═══════════════════════════════════" & vbCrLf
    msg = msg & " DETAIL COMPLIANCE STOCK" & vbCrLf
    msg = msg & "═══════════════════════════════════" & vbCrLf
    msg = msg & " SKU    : " & sku & vbCrLf
    msg = msg & " Size   : " & size & vbCrLf
    msg = msg & " Qty/hr : " & qcur & vbCrLf
    msg = msg & " Status : " & UCase(stat) & vbCrLf
    msg = msg & "───────────────────────────────────" & vbCrLf
    msg = msg & " MATERIAL           JAM   STATUS" & vbCrLf
    msg = msg & "───────────────────────────────────" & vbCrLf

    Dim i As Long
    For i = 0 To 13
        Dim colIdx As Long
        colIdx = COL_JAM_START + i
        If colIdx <= COL_JAM_END Then
            Dim jamTxt As String
            jamTxt = Trim(CStr(ws.Cells(rJam, colIdx).Value))
            Dim icon As String
            icon = StatusIcon(jamTxt)
            msg = msg & " " & PadRight(matNames(i), 18) & PadRight(jamTxt, 7) & icon & vbCrLf
        End If
    Next i

    msg = msg & "═══════════════════════════════════"
    MsgBox msg, vbInformation, "Detail: " & sku
End Sub

Private Function StatusIcon(jamTxt As String) As String
    Dim num As String
    num = Replace(jamTxt, " Jam", "")
    If Not IsNumeric(num) Or jamTxt = "-" Then
        StatusIcon = "  —"
    ElseIf CDbl(num) < 8 Then
        StatusIcon = "  ⚠ KRITIS"
    ElseIf CDbl(num) < 16 Then
        StatusIcon = "  ⚡ WARNING"
    Else
        StatusIcon = "  ✔ AMAN"
    End If
End Function

Private Function PadRight(s As String, n As Long) As String
    PadRight = s & Space(IIf(Len(s) < n, n - Len(s), 0))
End Function

' ================================================================
' 7. ADD BUTTONS KE SHEET (jalankan sekali)
' ================================================================

Sub AddDashboardButtons()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_MAIN)

    ' Hapus tombol lama jika ada
    Dim shp As Shape
    For Each shp In ws.Shapes
        If InStr(shp.Name, "btn_") > 0 Then shp.Delete
    Next shp

    ' Definisi tombol: Array(nama, caption, sub, kiri, atas, lebar, tinggi, warna)
    Dim buttons(5) As Variant
    buttons(0) = Array("btn_refresh",  "🔄 REFRESH",       "RefreshDashboard", 10,  5, 120, 28, CLR_HEADER)
    buttons(1) = Array("btn_kritis",   "🔴 LIHAT KRITIS",  "ShowKritisOnly",   140, 5, 120, 28, CLR_KRITIS)
    buttons(2) = Array("btn_showall",  "🔁 TAMPIL SEMUA",  "ShowAllRows",      270, 5, 120, 28, CLR_SUBHDR)
    buttons(3) = Array("btn_filter",   "🔍 FILTER MESIN",  "FilterByMesin",    400, 5, 120, 28, 7829367)
    buttons(4) = Array("btn_pdf",      "📄 EXPORT PDF",    "ExportPDF",        530, 5, 120, 28, 5287936)
    buttons(5) = Array("btn_nav",      "📊 MATERIAL SHEET","NavigateMaterial", 660, 5, 120, 28, 9127187)

    Dim i As Long
    For i = 0 To 5
        Dim b As Variant
        b = buttons(i)

        Dim btn As Shape
        Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
            b(3), b(4), b(5), b(6))

        btn.Name = b(0)

        With btn.TextFrame2.TextRange
            .Text = b(1)
            .Font.Size = 9
            .Font.Bold = True
            .Font.Fill.ForeColor.RGB = CLR_WHITE
            .ParagraphFormat.Alignment = msoAlignCenter
        End With

        With btn.Fill
            .ForeColor.RGB = b(7)
            .BackColor.RGB = b(7)
        End With

        With btn.Line
            .ForeColor.RGB = b(7)
            .Weight = 0.5
        End With

        btn.Shadow.Visible = msoTrue
        btn.Shadow.OffsetX = 1.5
        btn.Shadow.OffsetY = 1.5
        btn.Shadow.Blur = 3

        btn.OnAction = b(2)
    Next i

    MsgBox "Tombol dashboard berhasil ditambahkan!", vbInformation, "Setup"
End Sub

' ================================================================
' 8. NAVIGASI SHEET MATERIAL
' ================================================================

Sub NavigateMaterial()
    Dim matList As Variant
    matList = Array("Tread", "Sidewall", "BEC", "Bead Finish", "Apex", _
                    "BodyPly", "InnerLiner", "Steel Belt", "Zerobelt")

    Dim choice As String
    choice = InputBox("Pilih sheet material:" & vbCrLf & _
        "1. Tread" & vbCrLf & "2. Sidewall" & vbCrLf & "3. BEC" & vbCrLf & _
        "4. Bead Finish" & vbCrLf & "5. Apex" & vbCrLf & "6. BodyPly" & vbCrLf & _
        "7. InnerLiner" & vbCrLf & "8. Steel Belt" & vbCrLf & "9. Zerobelt" & vbCrLf & _
        vbCrLf & "Ketik angka:", "Navigasi Sheet Material", "1")

    If choice = "" Then Exit Sub
    If Not IsNumeric(choice) Then Exit Sub

    Dim idx As Long
    idx = CLng(choice) - 1
    If idx < 0 Or idx > UBound(matList) Then Exit Sub

    Dim targetSheet As String
    targetSheet = matList(idx)

    On Error GoTo SheetNotFound
    ThisWorkbook.Sheets(targetSheet).Activate
    Exit Sub

SheetNotFound:
    MsgBox "Sheet '" & targetSheet & "' tidak ditemukan.", vbExclamation, "Navigasi"
End Sub

' ================================================================
' 9. AUTO-RUN SAAT BUKA FILE (ditempatkan di ThisWorkbook)
' ================================================================
' Salin kode berikut ke modul ThisWorkbook:
'
' Private Sub Workbook_Open()
'     Application.ScreenUpdating = False
'     Call Module_ComplianceStock.ColorizeJamCells(ThisWorkbook.Sheets("Print Schedule Building"))
'     Call Module_ComplianceStock.BuildSummaryPanel(ThisWorkbook.Sheets("Print Schedule Building"))  ' private - skip
'     Application.ScreenUpdating = True
' End Sub
'
' Atau cukup panggil: Call RefreshDashboard
' ================================================================

' ================================================================
' 10. CONDITIONAL FORMAT LEGEND — buat kotak legend di sheet
' ================================================================

Sub AddLegend()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_MAIN)

    ' Tulis legend di baris 11-13 kolom AH-AJ
    Dim legendRow As Long: legendRow = 11

    ws.Cells(legendRow, 37).Value = "LEGENDA WARNA"
    ws.Cells(legendRow, 37).Font.Bold = True
    ws.Cells(legendRow, 37).Interior.Color = CLR_HEADER
    ws.Cells(legendRow, 37).Font.Color = CLR_WHITE

    ws.Cells(legendRow + 1, 37).Value = "Merah = Kritis < 8 Jam"
    ws.Cells(legendRow + 1, 37).Interior.Color = CLR_KRITIS
    ws.Cells(legendRow + 1, 37).Font.Color = CLR_WHITE
    ws.Cells(legendRow + 1, 37).Font.Bold = True

    ws.Cells(legendRow + 2, 37).Value = "Oranye = Warning 8-16 Jam"
    ws.Cells(legendRow + 2, 37).Interior.Color = CLR_WARNING

    ws.Cells(legendRow + 3, 37).Value = "Hijau = Aman > 16 Jam"
    ws.Cells(legendRow + 3, 37).Interior.Color = CLR_AMAN
    ws.Cells(legendRow + 3, 37).Font.Color = CLR_WHITE

    ws.Cells(legendRow + 4, 37).Value = "Abu = Tidak Ada Material"
    ws.Cells(legendRow + 4, 37).Interior.Color = 14671839

    MsgBox "Legend ditambahkan.", vbInformation, "Legend"
End Sub

' ================================================================
' 11. QUICK STATS — hitung dan tampilkan statistik cepat
' ================================================================

Sub QuickStats()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_MAIN)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    Dim cntSKU As Long, cntKritis As Long, cntWarn As Long, cntAman As Long
    Dim minJam As Double: minJam = 9999
    Dim minMat As String, minSKU As String

    Dim r As Long, c As Long
    For r = ROW_START To lastRow Step ROWS_PER_SKU
        cntSKU = cntSKU + 1
        Dim rJam As Long: rJam = r + 1
        If rJam > lastRow Then Exit For

        For c = COL_JAM_START To COL_JAM_END
            Dim txt As String
            txt = Replace(Trim(CStr(ws.Cells(rJam, c).Value)), " Jam", "")
            If IsNumeric(txt) Then
                Dim v As Double: v = CDbl(txt)
                If v < 8 Then
                    cntKritis = cntKritis + 1
                ElseIf v < 16 Then
                    cntWarn = cntWarn + 1
                Else
                    cntAman = cntAman + 1
                End If
                If v < minJam Then
                    minJam = v
                    minMat = ws.Cells(14, c).Value & " (Col " & c & ")"
                    minSKU = ws.Cells(r, 2).Value
                End If
            End If
        Next c
    Next r

    Dim msg As String
    msg = "═══════════ QUICK STATS ══════════" & vbCrLf
    msg = msg & " Total SKU Aktif  : " & cntSKU & vbCrLf
    msg = msg & " Total Item KRITIS: " & cntKritis & " item" & vbCrLf
    msg = msg & " Total Item WARN  : " & cntWarn & " item" & vbCrLf
    msg = msg & " Total Item AMAN  : " & cntAman & " item" & vbCrLf
    msg = msg & "──────────────────────────────────" & vbCrLf
    msg = msg & " Stok Paling Kritis:" & vbCrLf
    msg = msg & "   SKU  : " & minSKU & vbCrLf
    msg = msg & "   Mat  : " & minMat & vbCrLf
    msg = msg & "   Sisa : " & Format(minJam, "0.0") & " Jam" & vbCrLf
    msg = msg & "═════════════════════════════════"

    MsgBox msg, vbInformation, "Quick Stats - " & Format(Now, "dd/mm/yyyy HH:mm")
End Sub
