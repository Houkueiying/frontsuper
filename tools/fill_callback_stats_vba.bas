Option Explicit

' ===== 設定（一般不用改）=====
Private Const SRC_FOLDER As String = "D:\電訪追蹤情形\11510"
Private Const STATS_SHEET As String = "115年回收率"
Private Const SRC_DOCTOR_COL As Long = 7       ' G 欄：醫師
Private Const SRC_NOTSEEN_COL As Long = 31     ' AE 欄：[未看診]
Private Const NOTSEEN_TEXT As String = "[未看診]"
Private Const DOCTOR_FIRST_ROW As Long = 4     ' A4
Private Const DOCTOR_LAST_ROW As Long = 21     ' A21
Private Const ROW_TOTAL As Long = 2            ' 實際看診人數
Private Const ROW_CALLBACK As Long = 22        ' 病人實際衛教人數
Private Const HEADER_ROW As Long = 1           ' 日期標題列
Private Const FORMULA_LAST_ROW As Long = 30    ' 新增欄位時，複製公式到哪一列

Public Sub 預覽()
    RunFill False
End Sub

Public Sub 正式填入()
    If MsgBox("確定要寫入統計表嗎？缺少的日期欄會自動新增，寫入前會先自動備份。", vbYesNo + vbQuestion, "正式填入") = vbYes Then
        RunFill True
    End If
End Sub

Private Sub RunFill(ByVal writeMode As Boolean)
    Dim stats As Worksheet
    Dim doctorRow As Object, dateCol As Object, unmatched As Object, done As Object
    Dim totalD As Object, notSeenD As Object, cbD As Object, recsD As Object
    Dim files As Collection, fname As Variant, f As Variant
    Dim wb As Workbook, ws As Worksheet
    Dim key As String, log As String, bakPath As String, warn As String, tmp As String
    Dim k As Variant, nmk As Variant
    Dim col As Long, seen As Long, r As Long, insAt As Long
    Dim doWrite As Boolean

    On Error Resume Next
    Set stats = ThisWorkbook.Worksheets(STATS_SHEET)
    On Error GoTo 0
    If stats Is Nothing Then
        MsgBox "找不到工作表「" & STATS_SHEET & "」，請在統計表檔案中執行。", vbExclamation
        Exit Sub
    End If
    If Dir(SRC_FOLDER, vbDirectory) = "" Then
        MsgBox "找不到資料夾：" & SRC_FOLDER, vbExclamation
        Exit Sub
    End If

    ' 醫師名字 → 列號（A4~A21）
    Set doctorRow = CreateObject("Scripting.Dictionary")
    For r = DOCTOR_FIRST_ROW To DOCTOR_LAST_ROW
        key = NormName(stats.Cells(r, 1).Value)
        If key <> "" Then doctorRow(key) = r
    Next r

    ' 日期 → 欄號（第 1 列）
    Set dateCol = BuildDateCols(stats, warn)
    log = warn

    Set totalD = CreateObject("Scripting.Dictionary")
    Set notSeenD = CreateObject("Scripting.Dictionary")
    Set cbD = CreateObject("Scripting.Dictionary")
    Set recsD = CreateObject("Scripting.Dictionary")
    Set unmatched = CreateObject("Scripting.Dictionary")
    Set done = CreateObject("Scripting.Dictionary")

    ' 先收集檔名
    Set files = New Collection
    fname = Dir(SRC_FOLDER & "\*.xlsx")
    Do While fname <> ""
        If Left(fname, 2) <> "~$" Then files.Add fname
        fname = Dir()
    Loop

    ' 逐檔讀取，依檔名日期彙總
    For Each f In files
        key = MDKeyFromFileName(CStr(f))
        If key = "" Then
            log = log & "[略過] 檔名看不出日期：" & f & vbLf
        Else
            Set wb = Workbooks.Open(SRC_FOLDER & "\" & f, ReadOnly:=True)
            Set ws = wb.Worksheets(1)
            CountSheet ws, key, totalD, notSeenD, cbD, recsD
            log = log & "[讀取] " & f & " → " & key & vbLf
            wb.Close SaveChanges:=False
        End If
    Next f

    ' 寫入前先備份（存到本機資料夾，因為統計表若在 OneDrive 上，無法直接備份到同一路徑）
    If writeMode Then
        Dim bakFolder As String
        bakFolder = Environ("USERPROFILE") & "\電訪統計備份"
        If Dir(bakFolder, vbDirectory) = "" Then MkDir bakFolder
        bakPath = bakFolder & "\統計表_備份_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsm"
        On Error Resume Next
        ThisWorkbook.SaveCopyAs bakPath
        If Err.Number <> 0 Then
            On Error GoTo 0
            MsgBox "備份失敗，已停止，沒有寫入任何資料。" & vbLf & Err.Description, vbCritical, "備份失敗"
            Exit Sub
        End If
        On Error GoTo 0
        log = log & "[備份] " & bakPath & vbLf
    End If

    ' 補上統計表缺少的日期欄（預覽只列出，正式填入才真的新增）
    Do
        key = NextMissingKey(totalD, dateCol, done)
        If key = "" Then Exit Do
        done(key) = True
        If writeMode Then
            insAt = InsertPos(dateCol, key)
            stats.Columns(insAt).Insert
            stats.Cells(HEADER_ROW, insAt).Value = Split(key, "/")(0) & "月" & Split(key, "/")(1) & "日"
            CopyFormulas stats, insAt
            Set dateCol = BuildDateCols(stats, tmp)
            log = log & "[已新增欄位] " & key & " → 第 " & insAt & " 欄" & vbLf
        Else
            log = log & "[將新增欄位] " & key & vbLf
        End If
    Loop

    ' 產生結果
    For Each k In totalD.Keys
        seen = totalD(k) - notSeenD(k)
        col = 0
        If dateCol.Exists(k) Then col = dateCol(k)
        doWrite = (writeMode And col > 0)

        log = log & k & " → 實際看診 " & seen & "、病人實際衛教 " & cbD(k) & vbLf
        If doWrite Then
            stats.Cells(ROW_TOTAL, col).Value = seen
            stats.Cells(ROW_CALLBACK, col).Value = cbD(k)
        End If

        For Each nmk In recsD(k).Keys
            If doctorRow.Exists(nmk) Then
                log = log & "    " & nmk & " → 第" & doctorRow(nmk) & "列 = " & recsD(k)(nmk) & vbLf
                If doWrite Then stats.Cells(doctorRow(nmk), col).Value = recsD(k)(nmk)
            Else
                unmatched(nmk) = True
            End If
        Next nmk
    Next k

    If unmatched.Count > 0 Then
        log = log & "[警告] 以下醫師不在 A4~A21，未填入：" & Join(unmatched.Keys, "、") & vbLf
    End If

    If writeMode Then
        ThisWorkbook.Save
        log = log & vbLf & "【已正式寫入】"
    Else
        log = log & vbLf & "【預覽】未寫入任何內容"
    End If

    Debug.Print log
    MsgBox Left(log, 1500), vbInformation, "電訪統計"
End Sub

' 統計表日期欄：回傳 {"月/日": 欄號}；重複日期會記錄到 warn
Private Function BuildDateCols(stats As Worksheet, ByRef warn As String) As Object
    Dim d As Object, col As Long, lastCol As Long, key As String
    Set d = CreateObject("Scripting.Dictionary")
    lastCol = stats.Cells(HEADER_ROW, stats.Columns.Count).End(xlToLeft).Column
    For col = 2 To lastCol
        key = MDKeyFromHeader(stats.Cells(HEADER_ROW, col).Value)
        If key <> "" Then
            If d.Exists(key) Then
                warn = warn & "[警告] 日期 " & key & " 重複，使用第 " & d(key) & " 欄" & vbLf
            Else
                d(key) = col
            End If
        End If
    Next col
    Set BuildDateCols = d
End Function

' 找出下一個缺少的日期（日期由小到大）
Private Function NextMissingKey(totalD As Object, dateCol As Object, done As Object) As String
    Dim k As Variant, best As String, bestNum As Long
    bestNum = 999999
    For Each k In totalD.Keys
        If Not dateCol.Exists(k) And Not done.Exists(k) Then
            If KeyNum(CStr(k)) < bestNum Then
                bestNum = KeyNum(CStr(k))
                best = CStr(k)
            End If
        End If
    Next k
    NextMissingKey = best
End Function

' 新日期要插入的位置：插在日期比它早、最接近的欄位之後
Private Function InsertPos(dateCol As Object, ByVal key As String) As Long
    Dim k As Variant, nk As Long, bestNum As Long, bestCol As Long, minCol As Long
    nk = KeyNum(key)
    bestNum = -1
    bestCol = 0
    minCol = 999999
    For Each k In dateCol.Keys
        If dateCol(k) < minCol Then minCol = dateCol(k)
        If KeyNum(CStr(k)) < nk And KeyNum(CStr(k)) > bestNum Then
            bestNum = KeyNum(CStr(k))
            bestCol = dateCol(k)
        End If
    Next k
    If bestCol > 0 Then
        InsertPos = bestCol + 1
    ElseIf minCol < 999999 Then
        InsertPos = minCol
    Else
        InsertPos = 2
    End If
End Function

' 新欄位的公式（第 23 列以後）從左邊那欄複製，格式也一併沿用
Private Sub CopyFormulas(stats As Worksheet, ByVal insAt As Long)
    Dim src As Long, r As Long
    If insAt > 2 Then
        src = insAt - 1
    Else
        src = insAt + 1
    End If
    For r = ROW_CALLBACK + 1 To FORMULA_LAST_ROW
        If stats.Cells(r, src).HasFormula Then
            stats.Cells(r, insAt).FormulaR1C1 = stats.Cells(r, src).FormulaR1C1
        End If
    Next r
    stats.Columns(insAt).ColumnWidth = stats.Columns(src).ColumnWidth
End Sub

' "月/日" 轉成可比較的數字，例如 10/10 → 1010
Private Function KeyNum(ByVal key As String) As Long
    Dim p As Variant
    p = Split(key, "/")
    KeyNum = CLng(p(0)) * 100 + CLng(p(1))
End Function

' 統計每個名單檔：總筆數、未看診、醫師筆數、黃底（已電訪）
Private Sub CountSheet(ws As Worksheet, ByVal key As String, totalD As Object, _
                       notSeenD As Object, cbD As Object, recsD As Object)
    Dim patCol As Long, lastR As Long, lastC As Long, i As Long, c As Long
    Dim nm As String, hit As Boolean
    Dim recs As Object

    If Not totalD.Exists(key) Then
        totalD(key) = 0
        notSeenD(key) = 0
        cbD(key) = 0
        Set recsD(key) = CreateObject("Scripting.Dictionary")
    End If
    Set recs = recsD(key)

    patCol = FindPatientCol(ws)
    lastC = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    lastR = ws.Cells(ws.Rows.Count, SRC_DOCTOR_COL).End(xlUp).Row

    For i = 2 To lastR
        nm = NormName(ws.Cells(i, SRC_DOCTOR_COL).Value)
        If nm <> "" Then
            totalD(key) = totalD(key) + 1
            If recs.Exists(nm) Then
                recs(nm) = recs(nm) + 1
            Else
                recs(nm) = 1
            End If

            If InStr(1, CStr(ws.Cells(i, SRC_NOTSEEN_COL).Value), NOTSEEN_TEXT, vbBinaryCompare) > 0 Then
                notSeenD(key) = notSeenD(key) + 1
            End If

            hit = False
            If patCol > 0 Then
                hit = IsYellow(ws.Cells(i, patCol))
            Else
                For c = 1 To lastC
                    If IsYellow(ws.Cells(i, c)) Then
                        hit = True
                        Exit For
                    End If
                Next c
            End If
            If hit Then cbD(key) = cbD(key) + 1
        End If
    Next i
End Sub

' 找標題含「姓名」的欄位（沒有就找「病人」）
Private Function FindPatientCol(ws As Worksheet) As Long
    Dim c As Long, lastC As Long, t As String
    lastC = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastC
        t = CStr(ws.Cells(1, c).Value)
        If InStr(t, "姓名") > 0 Then FindPatientCol = c: Exit Function
    Next c
    For c = 1 To lastC
        t = CStr(ws.Cells(1, c).Value)
        If InStr(t, "病人") > 0 Then FindPatientCol = c: Exit Function
    Next c
    FindPatientCol = 0
End Function

' 判斷是否為黃底
Private Function IsYellow(ByVal cl As Range) As Boolean
    Dim col As Long, rr As Long, gg As Long, bb As Long
    If cl.Interior.ColorIndex = xlNone Then Exit Function
    col = cl.Interior.Color
    rr = col Mod 256
    gg = (col \ 256) Mod 256
    bb = (col \ 65536) Mod 256
    IsYellow = (rr >= 220 And gg >= 200 And bb <= 120)
End Function

' 去除空白，方便比對名字
Private Function NormName(ByVal v As Variant) As String
    NormName = Replace(Replace(Trim(CStr(v)), " ", ""), vbTab, "")
End Function

' 統計表標題 → "月/日"（支援日期格式與「10月1日」文字）
Private Function MDKeyFromHeader(ByVal v As Variant) As String
    Dim re As Object, m As Object, s As String
    If VarType(v) = vbDate Then
        MDKeyFromHeader = Month(v) & "/" & Day(v)
        Exit Function
    End If
    s = CStr(v)
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "(\d{1,2})\s*月\s*(\d{1,2})"
    If re.Test(s) Then
        Set m = re.Execute(s)(0)
        MDKeyFromHeader = CLng(m.SubMatches(0)) & "/" & CLng(m.SubMatches(1))
    End If
End Function

' 檔名 → "月/日"（支援 115.10.1 民國點號格式、1151001 民國、20261001 西元）
Private Function MDKeyFromFileName(ByVal fn As String) As String
    Dim re As Object, mt As Object, s As String
    Dim y As Long, mo As Long, d As Long
    Set re = CreateObject("VBScript.RegExp")

    ' 1) 民國 115.10.1 （也接受 115-10-01、115_10_1）
    re.Pattern = "(?:^|\D)(\d{3})[.\-_](\d{1,2})[.\-_](\d{1,2})(?!\d)"
    If re.Test(fn) Then
        Set mt = re.Execute(fn)(0)
        y = CLng(mt.SubMatches(0)) + 1911
        mo = CLng(mt.SubMatches(1))
        d = CLng(mt.SubMatches(2))
    Else
        ' 2) 1151001（民國）或 20261001（西元）
        re.Pattern = "(?:^|\D)(\d{7}|\d{8})(?:\D|$)"
        If Not re.Test(fn) Then Exit Function
        s = re.Execute(fn)(0).SubMatches(0)
        If Len(s) = 7 Then
            y = CLng(Left(s, 3)) + 1911
            mo = CLng(Mid(s, 4, 2))
            d = CLng(Right(s, 2))
        Else
            y = CLng(Left(s, 4))
            mo = CLng(Mid(s, 5, 2))
            d = CLng(Right(s, 2))
        End If
    End If
    If mo < 1 Or mo > 12 Or d < 1 Or d > 31 Then Exit Function
    MDKeyFromFileName = mo & "/" & d
End Function
