Attribute VB_Name = "DomesticCheck"
Option Explicit

'=====================================================================
' 국내 구간 검증 매크로 (CheckDomestic)
'  사용법: 검증용 xlsm에 Alt+F11 > 파일 > 가져오기 > 이 파일 선택 (처음 한 번)
'          Alt+F8 > CheckDomestic 실행 > 1) 내 파일(빠로 xlsm) 선택 > 2) 오차장 xlsx 선택
'  기존 CheckHospital(베트남 구간)과 별개로 동작하며, 결과는 "국내검증결과" 시트로 만들어집니다.
'  표시본은 "원본이름_국내검증표시"로 저장됩니다 (빠로 파일/오차장 원본은 수정하지 않음).
'
'  국내 구간 = 병원 시트에서 마지막 PHARMATOPES 행 다음부터 Solution 표 앞까지
'  내 파일이 맞다는 전제로, 아래 항목만 봅니다.
'   1) 선택 파일의 번호(A열)가 베트남 마지막 번호 다음부터 1씩 이어지는지
'   2) 병원명 / 요일 / 규격 / 대리점이 내 파일과 같은지 (행 순서는 무시, 같은 조합의 개수로 비교)
'   - 병원명이 빈 행(구분선 등)은 비교에서 제외
'  다른 행은 가장 비슷한 행끼리 짝지어 무엇이 다른지(규격/요일/대리점/병원명) 알려줍니다.
'=====================================================================

Private Type DItem
    r As Long
    num As String
    hosp As String
    dy As String
    spec As Double
    dist As String
    used As Boolean
End Type

' 결과 목록 (1 구분, 2 선택 파일 행, 3 내 파일 행, 4 위치, 5 내용)
Private rep() As String
Private repKey() As Long
Private repN As Long

' 선택 파일(오차장 파일)에 색칠할 목록
Private mkN As Long
Private mkLv() As String
Private mkR() As Long
Private mkC() As String
Private mkMsg() As String

Private gA As Workbook
Private gB As Workbook

Private Function SN() As String
    SN = ChrW(&HBCD1) & ChrW(&HC6D0)
End Function

Private Function NL() As String
    NL = ChrW(&HBCD1) & ChrW(&HC6D0) & ChrW(&HBA85)
End Function

Private Sub AddRep(ByVal lvl As String, ByVal rowB As Long, ByVal rowA As Long, ByVal pos As String, ByVal msg As String)
    repN = repN + 1
    If repN > UBound(rep, 2) Then
        ReDim Preserve rep(1 To 5, 1 To UBound(rep, 2) + 200)
        ReDim Preserve repKey(1 To UBound(repKey) + 200)
    End If
    rep(1, repN) = lvl
    If rowB > 0 Then rep(2, repN) = CStr(rowB) Else rep(2, repN) = ""
    If rowA > 0 Then rep(3, repN) = CStr(rowA) Else rep(3, repN) = ""
    rep(4, repN) = pos
    rep(5, repN) = msg
    If rowB > 0 Then
        repKey(repN) = rowB
    Else
        repKey(repN) = 100000 + rowA
    End If
End Sub

Private Sub AddMark(ByVal lvl As String, ByVal r As Long, ByVal col As String, ByVal msg As String)
    mkN = mkN + 1
    If mkN > UBound(mkLv) Then
        ReDim Preserve mkLv(1 To UBound(mkLv) + 100)
        ReDim Preserve mkR(1 To UBound(mkR) + 100)
        ReDim Preserve mkC(1 To UBound(mkC) + 100)
        ReDim Preserve mkMsg(1 To UBound(mkMsg) + 100)
    End If
    mkLv(mkN) = lvl
    mkR(mkN) = r
    mkC(mkN) = col
    mkMsg(mkN) = msg
End Sub

Private Sub SortRep()
    Dim i As Long, j As Long, k As Long, t As String, tk As Long
    For i = 2 To repN
        For j = i To 2 Step -1
            If repKey(j) < repKey(j - 1) Then
                For k = 1 To 5
                    t = rep(k, j)
                    rep(k, j) = rep(k, j - 1)
                    rep(k, j - 1) = t
                Next k
                tk = repKey(j)
                repKey(j) = repKey(j - 1)
                repKey(j - 1) = tk
            Else
                Exit For
            End If
        Next j
    Next i
End Sub

Private Function Tx(ws As Worksheet, ByVal r As Long, ByVal col As String) As String
    Dim V As Variant
    V = ws.Range(col & r).Value
    If IsError(V) Then
        Tx = "#ERR"
    ElseIf IsEmpty(V) Then
        Tx = ""
    Else
        Tx = Trim$(CStr(V))
    End If
End Function

Private Function SpecOf(ws As Worksheet, ByVal r As Long) As Double
    Dim V As Variant
    V = ws.Range("F" & r).Value2
    If IsError(V) Then
        SpecOf = -1
    ElseIf IsEmpty(V) Then
        SpecOf = -1
    ElseIf IsNumeric(V) Then
        SpecOf = CDbl(V)
    Else
        SpecOf = -1
    End If
End Function

Private Function SpecShow(ByVal d As Double) As String
    If d < 0 Then
        SpecShow = "(읽지 못함 CHECK)"
    Else
        SpecShow = CStr(d) & "mCi"
    End If
End Function

Private Function KeyTx(it As DItem) As String
    KeyTx = it.hosp & " / " & it.dy & " / " & SpecShow(it.spec) & " / " & it.dist
End Function

' 병원명 시트 목록에 이름이 있는지: 1 있음, 0 없음, -1 시트를 찾지 못함
Private Function InNameList(wb As Workbook, ByVal nm As String) As Long
    Dim ws As Worksheet, lastR As Long, r As Long
    InNameList = -1
    On Error Resume Next
    Set ws = wb.Worksheets(NL())
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    InNameList = 0
    lastR = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For r = 2 To lastR
        If Not IsError(ws.Cells(r, 1).Value) Then
            If Trim$(CStr(ws.Cells(r, 1).Value)) = nm Then
                InNameList = 1
                Exit Function
            End If
        End If
    Next r
End Function

Private Function YesNo(ByVal V As Long) As String
    Select Case V
        Case 1: YesNo = "있음"
        Case 0: YesNo = "없음"
        Case Else: YesNo = "확인불가"
    End Select
End Function

' 국내 구간 행을 읽음. lastViet = 마지막 PHARMATOPES 행, lastEnd = 국내 구간 끝(Solution 표 바로 앞)
Private Sub LoadDomestic(ws As Worksheet, it() As DItem, ByRef n As Long, ByRef lastViet As Long, ByRef lastEnd As Long)
    Dim lastR As Long, r As Long, h As String
    n = 0
    lastViet = 0
    ReDim it(1 To 300)
    lastR = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    For r = 7 To lastR
        If InStr(1, Tx(ws, r, "A"), "Solution", vbTextCompare) > 0 Then
            lastR = r - 1
            Exit For
        End If
    Next r
    lastEnd = lastR
    For r = 7 To lastR
        If StrComp(Tx(ws, r, "J"), "PHARMATOPES", vbTextCompare) = 0 Then lastViet = r
    Next r
    If lastViet = 0 Then Exit Sub
    For r = lastViet + 1 To lastR
        h = Tx(ws, r, "C")
        If h <> "" And h <> "#ERR" And StrComp(h, "Total", vbTextCompare) <> 0 Then
            n = n + 1
            If n > UBound(it) Then ReDim Preserve it(1 To UBound(it) + 300)
            it(n).r = r
            it(n).num = Tx(ws, r, "A")
            it(n).hosp = h
            it(n).dy = Tx(ws, r, "E")
            it(n).spec = SpecOf(ws, r)
            it(n).dist = Tx(ws, r, "J")
            it(n).used = False
        End If
    Next r
End Sub

' 선택 파일의 번호(A열)가 베트남 마지막 번호 다음부터 1씩 이어지는지 (Total 행 앞까지)
Private Sub CheckNumbers(ws As Worksheet, ByVal lastViet As Long, ByVal lastEnd As Long, ByRef firstNo As String, ByRef lastNo As String)
    Dim r As Long, a As String, h As String, prev As Double, cur As Double, msg As String
    prev = Val(Tx(ws, lastViet, "A"))
    firstNo = ""
    lastNo = ""
    For r = lastViet + 1 To lastEnd
        h = Tx(ws, r, "C")
        If StrComp(h, "Total", vbTextCompare) = 0 Then Exit For
        a = Tx(ws, r, "A")
        If a = "" Then
            If h <> "" Then
                msg = "번호가 비어 있음"
                AddRep "오류", r, 0, h, msg
                AddMark "오류", r, "A", msg
            End If
        ElseIf Not IsNumeric(a) Then
            msg = "번호가 숫자가 아님: " & a
            AddRep "오류", r, 0, h, msg
            AddMark "오류", r, "A", msg
        Else
            cur = Val(a)
            If firstNo = "" Then firstNo = a
            lastNo = a
            If cur <> prev + 1 Then
                msg = "번호 불연속: 앞 번호 " & prev & " 다음이 " & cur & " (기대 " & (prev + 1) & ")"
                AddRep "오류", r, 0, h, msg
                AddMark "오류", r, "A", msg
            End If
            prev = cur
        End If
    Next r
End Sub

' mask 글자(H 병원명, D 요일, S 규격, X 대리점)가 모두 같으면 True
Private Function Same(a As DItem, b As DItem, ByVal mask As String) As Boolean
    Same = False
    If InStr(mask, "H") > 0 Then
        If a.hosp <> b.hosp Then Exit Function
    End If
    If InStr(mask, "D") > 0 Then
        If a.dy <> b.dy Then Exit Function
    End If
    If InStr(mask, "S") > 0 Then
        If Abs(a.spec - b.spec) > 0.0001 Then Exit Function
    End If
    If InStr(mask, "X") > 0 Then
        If a.dist <> b.dist Then Exit Function
    End If
    Same = True
End Function

' 아직 짝이 없는 행끼리 mask가 같은 첫 행과 짝지음. kind가 ""이면 완전 일치라 보고하지 않음
Private Sub PairTier(ia() As DItem, ByVal na As Long, ib() As DItem, ByVal nb As Long, ByVal mask As String, ByVal kind As String)
    Dim i As Long, j As Long, msg As String, col As String
    For i = 1 To na
        If Not ia(i).used Then
            For j = 1 To nb
                If Not ib(j).used Then
                    If Same(ia(i), ib(j), mask) Then
                        ia(i).used = True
                        ib(j).used = True
                        If kind <> "" Then
                            Select Case kind
                                Case "SPEC"
                                    msg = "규격 다름: 내 파일 " & SpecShow(ia(i).spec) & " / 선택 파일 " & SpecShow(ib(j).spec)
                                    col = "F"
                                Case "DAY"
                                    msg = "요일 다름: 내 파일 " & ia(i).dy & " / 선택 파일 " & ib(j).dy
                                    col = "E"
                                Case "DIST"
                                    msg = "대리점 다름: 내 파일 " & ia(i).dist & " / 선택 파일 " & ib(j).dist
                                    col = "J"
                                Case Else
                                    msg = "병원명 다름: 내 파일 [" & ia(i).hosp & "] / 선택 파일 [" & ib(j).hosp & "]" & _
                                          " (병원명 시트에 있는지: 내 파일 표기 " & YesNo(InNameList(gA, ia(i).hosp)) & _
                                          " / 선택 파일 표기 " & YesNo(InNameList(gB, ib(j).hosp)) & ")"
                                    col = "C"
                            End Select
                            AddRep "다름", ib(j).r, ia(i).r, KeyTx(ib(j)), msg
                            AddMark "다름", ib(j).r, col, msg
                        End If
                        Exit For
                    End If
                End If
            Next j
        End If
    Next i
End Sub

' 표시 목록대로 색칠하고, 원본은 건드리지 않고 "_국내검증표시" 사본으로 저장. 실패하면 빈 문자열
Private Function SaveMarked(wsB As Worksheet, wbB As Workbook, ByVal srcPath As String) As String
    Dim pass As Variant, i As Long, rg As Range, c As Range, cm As Object
    Dim p As Long, base As String, ext As String, tryPath As String
    On Error Resume Next
    For Each pass In Array("다름", "오류")
        For i = 1 To mkN
            If mkLv(i) = CStr(pass) Then
                Set rg = wsB.Range("A" & mkR(i) & ":K" & mkR(i))
                If mkLv(i) = "다름" Then
                    rg.Interior.Color = RGB(255, 235, 156)
                Else
                    rg.Interior.Color = RGB(255, 199, 206)
                End If
            End If
        Next i
    Next pass
    For i = 1 To mkN
        If mkC(i) <> "" Then
            If mkLv(i) = "다름" Then
                wsB.Range(mkC(i) & mkR(i)).Interior.Color = RGB(255, 192, 0)
            Else
                wsB.Range(mkC(i) & mkR(i)).Interior.Color = RGB(255, 110, 110)
            End If
        End If
    Next i
    For i = 1 To mkN
        If mkMsg(i) <> "" Then
            Set c = wsB.Range("A" & mkR(i))
            If c.Comment Is Nothing Then
                Set cm = c.AddComment(mkMsg(i))
                If Not cm Is Nothing Then
                    cm.Shape.Width = 320
                    cm.Shape.Height = 60
                End If
            Else
                c.Comment.Text c.Comment.Text & vbLf & mkMsg(i)
            End If
        End If
    Next i
    Err.Clear
    p = InStrRev(srcPath, ".")
    base = Left$(srcPath, p - 1)
    ext = Mid$(srcPath, p)
    tryPath = base & "_국내검증표시" & ext
    wbB.SaveCopyAs tryPath
    If Err.Number <> 0 Then
        Err.Clear
        tryPath = ThisWorkbook.path & "\" & Mid$(base, InStrRev(base, "\") + 1) & "_국내검증표시" & ext
        wbB.SaveCopyAs tryPath
        If Err.Number <> 0 Then tryPath = ""
    End If
    SaveMarked = tryPath
    On Error GoTo 0
End Function

Private Function IsOpenFile(ByVal path As String) As Boolean
    Dim w As Workbook
    For Each w In Workbooks
        If StrComp(w.FullName, path, vbTextCompare) = 0 Then
            IsOpenFile = True
            Exit Function
        End If
    Next w
End Function

' 이미 열려 있으면 그 통합문서를 그대로 쓰고(openedHere=False), 아니면 읽기 전용으로 엶
Private Function OpenReadOnly(ByVal path As String, ByRef openedHere As Boolean) As Workbook
    Dim w As Workbook
    openedHere = False
    For Each w In Workbooks
        If StrComp(w.FullName, path, vbTextCompare) = 0 Then
            Set OpenReadOnly = w
            Exit Function
        End If
    Next w
    Set OpenReadOnly = Workbooks.Open(Filename:=path, ReadOnly:=True, UpdateLinks:=0)
    openedHere = True
End Function

Public Sub CheckDomestic()
    Dim fpA As Variant, fpB As Variant, wbA As Workbook, wbB As Workbook
    Dim wsA As Worksheet, wsB As Worksheet, wsOut As Worksheet
    Dim openedA As Boolean, openedB As Boolean, oldSec As Long
    Dim ia() As DItem, ib() As DItem
    Dim na As Long, nb As Long, lastVietA As Long, lastVietB As Long, lastEndA As Long, lastEndB As Long
    Dim firstNoB As String, lastNoB As String
    Dim i As Long, j As Long, nErr As Long, nDiff As Long
    Dim outPath As String, errMsg As String, msg As String, sInfoA As String, sInfoB As String
    Dim outArr() As Variant

    fpA = Application.GetOpenFilename("Excel (*.xlsx;*.xlsm),*.xlsx;*.xlsm", , "1) 내 파일(빠로 xlsm)을 선택하세요")
    If VarType(fpA) = vbBoolean Then Exit Sub
    fpB = Application.GetOpenFilename("Excel (*.xlsx;*.xlsm),*.xlsx;*.xlsm", , "2) 비교할 파일(오차장 파일)을 선택하세요")
    If VarType(fpB) = vbBoolean Then Exit Sub
    If StrComp(CStr(fpA), CStr(fpB), vbTextCompare) = 0 Then
        MsgBox "같은 파일을 두 번 선택했습니다.", vbExclamation
        Exit Sub
    End If
    If IsOpenFile(CStr(fpB)) Then
        MsgBox "오차장 파일이 이미 열려 있습니다. 닫고 다시 실행해 주세요.", vbExclamation
        Exit Sub
    End If

    oldSec = Application.AutomationSecurity
    On Error GoTo EH
    Application.ScreenUpdating = False
    Application.AutomationSecurity = 3
    Set wbA = OpenReadOnly(CStr(fpA), openedA)
    Set wbB = OpenReadOnly(CStr(fpB), openedB)
    Application.AutomationSecurity = oldSec
    Set gA = wbA
    Set gB = wbB
    Set wsA = wbA.Worksheets(SN())
    Set wsB = wbB.Worksheets(SN())

    ReDim rep(1 To 5, 1 To 200)
    ReDim repKey(1 To 200)
    repN = 0
    ReDim mkLv(1 To 100)
    ReDim mkR(1 To 100)
    ReDim mkC(1 To 100)
    ReDim mkMsg(1 To 100)
    mkN = 0

    LoadDomestic wsA, ia, na, lastVietA, lastEndA
    LoadDomestic wsB, ib, nb, lastVietB, lastEndB
    If lastVietA = 0 Or lastVietB = 0 Then
        MsgBox "PHARMATOPES 행을 찾지 못해 국내 구간을 구분할 수 없습니다. (J열 대리점 확인)", vbExclamation
        GoTo CleanUp
    End If

    ' 1) 번호 검사 (선택 파일)
    CheckNumbers wsB, lastVietB, lastEndB, firstNoB, lastNoB

    ' 2) 병원명 / 요일 / 규격 / 대리점 비교: 완전 일치 -> 규격 -> 요일 -> 대리점 -> 병원명 순으로 짝짓기
    PairTier ia, na, ib, nb, "HDSX", ""
    PairTier ia, na, ib, nb, "HDX", "SPEC"
    PairTier ia, na, ib, nb, "HSX", "DAY"
    PairTier ia, na, ib, nb, "HDS", "DIST"
    PairTier ia, na, ib, nb, "DSX", "HOSP"
    For j = 1 To nb
        If Not ib(j).used Then
            msg = "내 파일에 없는 행 (선택 파일에만 있음)"
            AddRep "다름", ib(j).r, 0, KeyTx(ib(j)), msg
            AddMark "다름", ib(j).r, "", msg
        End If
    Next j
    For i = 1 To na
        If Not ia(i).used Then
            AddRep "다름", 0, ia(i).r, KeyTx(ia(i)), "선택 파일에 없는 행 (내 파일에만 있음)"
        End If
    Next i
    SortRep
    outPath = SaveMarked(wsB, wbB, CStr(fpB))

    For i = 1 To repN
        If rep(1, i) = "오류" Then nErr = nErr + 1 Else nDiff = nDiff + 1
    Next i

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("국내검증결과").Delete
    On Error GoTo EH
    Application.DisplayAlerts = True
    Set wsOut = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsOut.Name = "국내검증결과"

    wsOut.Range("A1").Value = "국내 구간 검증 결과"
    wsOut.Range("A1").Font.Bold = True
    wsOut.Range("A1").Font.Size = 14
    sInfoA = "내 파일: " & wbA.Name & " / 국내 " & na & "행"
    If na > 0 Then sInfoA = sInfoA & " (시트 행 " & ia(1).r & "~" & ia(na).r & ", 번호 " & ia(1).num & "~" & ia(na).num & ")"
    sInfoB = "선택 파일: " & wbB.Name & " / 국내 " & nb & "행"
    If nb > 0 Then sInfoB = sInfoB & " (시트 행 " & ib(1).r & "~" & ib(nb).r & ", 번호 " & firstNoB & "~" & lastNoB & ", 병원명이 빈 구분선 행 제외)"
    wsOut.Range("A2").Value = sInfoA
    wsOut.Range("A3").Value = sInfoB
    wsOut.Range("A4").Value = "오류 " & nErr & "건 (번호) / 다름 " & nDiff & "건 (내 파일과 비교)"
    wsOut.Range("A4").Font.Bold = True
    If outPath <> "" Then
        wsOut.Range("A5").Value = "표시본 저장: " & outPath & " (연한 빨강=번호 오류, 노랑=내 파일과 다른 행, 진한 색 칸=다른 항목, 이유는 A열 셀 메모)"
    Else
        wsOut.Range("A5").Value = "표시본을 저장하지 못했습니다 (폴더 쓰기 권한을 확인하세요)"
    End If

    wsOut.Range("A7:E7").Value = Array("구분", "선택 파일 행", "내 파일 행", "위치 (병원명 / 요일 / 규격 / 대리점)", "내용")
    wsOut.Range("A7:E7").Font.Bold = True
    wsOut.Range("A7:E7").Interior.Color = RGB(217, 217, 217)

    If repN > 0 Then
        ReDim outArr(1 To repN, 1 To 5)
        For i = 1 To repN
            outArr(i, 1) = rep(1, i)
            If rep(2, i) = "" Then outArr(i, 2) = "" Else outArr(i, 2) = CLng(rep(2, i))
            If rep(3, i) = "" Then outArr(i, 3) = "" Else outArr(i, 3) = CLng(rep(3, i))
            outArr(i, 4) = rep(4, i)
            outArr(i, 5) = rep(5, i)
            If rep(1, i) = "오류" Then
                wsOut.Range("A" & (7 + i) & ":E" & (7 + i)).Font.Color = RGB(192, 0, 0)
            End If
        Next i
        wsOut.Range("A8").Resize(repN, 5).Value = outArr
        wsOut.Range("A7:E" & (7 + repN)).AutoFilter
    Else
        wsOut.Range("A8").Value = "발견된 문제가 없습니다."
    End If
    wsOut.Columns("A").ColumnWidth = 8
    wsOut.Columns("B").ColumnWidth = 13
    wsOut.Columns("C").ColumnWidth = 11
    wsOut.Columns("D").ColumnWidth = 42
    wsOut.Columns("E").ColumnWidth = 100

CleanUp:
    wbB.Close SaveChanges:=False
    If openedA Then wbA.Close SaveChanges:=False
    Application.ScreenUpdating = True
    If Not wsOut Is Nothing Then
        wsOut.Activate
        MsgBox "국내 검증 완료: 번호 오류 " & nErr & "건, 다름 " & nDiff & "건" & vbCrLf & "표시본: " & outPath, vbInformation
    End If
    Exit Sub

EH:
    errMsg = Err.Description
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    On Error Resume Next
    Application.AutomationSecurity = oldSec
    If Not wbB Is Nothing Then
        If openedB Then wbB.Close SaveChanges:=False
    End If
    If Not wbA Is Nothing Then
        If openedA Then wbA.Close SaveChanges:=False
    End If
    MsgBox "실행 중 오류가 발생했습니다: " & errMsg, vbExclamation
End Sub
