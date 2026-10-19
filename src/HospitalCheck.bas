Attribute VB_Name = "HospitalCheck"
Option Explicit

'=====================================================================
' 병원 시트 검증 매크로
'  사용법: 검증용 xlsm 하나에 Alt+F11 > 파일 > 가져오기 > 이 파일 선택 (처음 한 번)
'          Alt+F8 > CheckHospital 실행 > 1) 내 파일(빠로 xlsm) 선택 > 2) 오차장 xlsx 선택
'  결과는 검증용 xlsm의 "검증결과" 시트로 만들어집니다 (빠로 파일/오차장 원본은 수정하지 않음)
'  검사 항목
'   1) 바이알(No. 시작, 최대 4행) 안에서 병원/날짜/요일/규격이 같은지
'   2) 바이알 mCi 표기 = 행수 x 규격 인지
'   3) 패키지(Q열 Package) 당 바이알 2개 이하, 병원/요일/규격이 섞이지 않았는지
'   4) 패키지 병합 범위가 실제 묶인 행수와 같은지
'   5) 병원명이 비었거나 어느 바이알에도 묶이지 않은 행
'   6) 두 파일의 병원/요일/규격별 행수와 묶는 방식 비교
'=====================================================================

Private Type VialInfo
    r1 As Long
    r2 As Long
    nRows As Long
    hosp As String
    cal As String
    dy As String
    spec As Double
End Type

Private Type PkgInfo
    v1 As Long
    v2 As Long
    r1 As Long
    r2 As Long
End Type

Private rep() As String
Private repN As Long

' 선택 파일(오차장 파일)에 색칠할 영역 목록
Private mkN As Long
Private mkLv() As String
Private mkR1() As Long
Private mkR2() As Long
Private mkC1() As String
Private mkC2() As String
Private mkMsg() As String

Private Function WK() As String
    WK = ChrW(&HC6D4) & ChrW(&HD654) & ChrW(&HC218) & ChrW(&HBAA9) & ChrW(&HAE08) & ChrW(&HD1A0) & ChrW(&HC77C)
End Function

Private Function SN() As String
    SN = ChrW(&HBCD1) & ChrW(&HC6D0)
End Function

Private Sub AddMark(ByVal lv As String, ByVal r1 As Long, ByVal r2 As Long, ByVal c1 As String, ByVal c2 As String, ByVal msg As String)
    mkN = mkN + 1
    If mkN > UBound(mkLv) Then
        ReDim Preserve mkLv(1 To UBound(mkLv) + 200)
        ReDim Preserve mkR1(1 To UBound(mkR1) + 200)
        ReDim Preserve mkR2(1 To UBound(mkR2) + 200)
        ReDim Preserve mkC1(1 To UBound(mkC1) + 200)
        ReDim Preserve mkC2(1 To UBound(mkC2) + 200)
        ReDim Preserve mkMsg(1 To UBound(mkMsg) + 200)
    End If
    mkLv(mkN) = lv
    mkR1(mkN) = r1
    mkR2(mkN) = r2
    mkC1(mkN) = c1
    mkC2(mkN) = c2
    mkMsg(mkN) = msg
End Sub

Private Sub AddRep(ByVal scope As String, ByVal level As String, ByVal ref As String, ByVal msg As String, _
                   Optional ByVal r1 As Long = 0, Optional ByVal r2 As Long = 0)
    repN = repN + 1
    If repN > UBound(rep, 2) Then ReDim Preserve rep(1 To 4, 1 To UBound(rep, 2) + 500)
    rep(1, repN) = scope
    rep(2, repN) = level
    rep(3, repN) = ref
    rep(4, repN) = msg
    ' 선택 파일의 규칙 오류는 해당 행을 표시 목록에 추가
    If scope = "선택 파일" And r1 > 0 And level = "오류" Then AddMark "오류", r1, r2, "A", "Q", msg
End Sub

' 선택 파일에서 같은 병원/요일/규격 키를 가진 모든 바이알(패키지) 행을 노란색 표시 대상으로 추가
Private Sub MarkKey(vB() As VialInfo, ByVal nVB As Long, ByVal k As String, ByVal msg As String)
    Dim i As Long, firstOne As Boolean
    firstOne = True
    For i = 1 To nVB
        If KeyOf(vB(i)) = k Then
            If firstOne Then
                AddMark "다름", vB(i).r1, vB(i).r2, "A", "Q", msg
                firstOne = False
            Else
                AddMark "다름", vB(i).r1, vB(i).r2, "A", "Q", ""
            End If
        End If
    Next i
End Sub

' 표시 목록대로 색칠하고, 원본은 건드리지 않고 "_검증표시" 사본으로 저장. 실패하면 빈 문자열
Private Function SaveMarked(wsB As Worksheet, wbB As Workbook, ByVal srcPath As String) As String
    Dim pass As Variant, i As Long, rg As Range, cm As Object
    Dim p As Long, base As String, ext As String, tryPath As String
    On Error Resume Next
    For Each pass In Array("다름", "오류", "오류강조")
        For i = 1 To mkN
            If mkLv(i) = CStr(pass) Then
                Set rg = wsB.Range(mkC1(i) & mkR1(i) & ":" & mkC2(i) & mkR2(i))
                Select Case mkLv(i)
                    Case "다름": rg.Interior.Color = RGB(255, 235, 156)
                    Case "오류": rg.Interior.Color = RGB(255, 199, 206)
                    Case "오류강조": rg.Interior.Color = RGB(255, 110, 110)
                End Select
            End If
        Next i
    Next pass
    For i = 1 To mkN
        If mkMsg(i) <> "" Then
            Set cm = Nothing
            Set cm = wsB.Range("A" & mkR1(i)).AddComment(mkMsg(i))
            If Not cm Is Nothing Then
                cm.Shape.Width = 300
                cm.Shape.Height = 70
            End If
        End If
    Next i
    Err.Clear
    p = InStrRev(srcPath, ".")
    base = Left$(srcPath, p - 1)
    ext = Mid$(srcPath, p)
    tryPath = base & "_검증표시" & ext
    wbB.SaveCopyAs tryPath
    If Err.Number <> 0 Then
        Err.Clear
        tryPath = ThisWorkbook.path & "\" & Mid$(base, InStrRev(base, "\") + 1) & "_검증표시" & ext
        wbB.SaveCopyAs tryPath
        If Err.Number <> 0 Then tryPath = ""
    End If
    SaveMarked = tryPath
    On Error GoTo 0
End Function

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

Private Function DateTx(ws As Worksheet, ByVal addr As String) As String
    Dim V As Variant
    V = ws.Range(addr).Value
    If IsError(V) Then
        DateTx = "#ERR"
    ElseIf IsEmpty(V) Then
        DateTx = ""
    ElseIf IsDate(V) Then
        DateTx = Format$(V, "yyyy-mm-dd")
    ElseIf IsNumeric(V) Then
        DateTx = Format$(CDate(CDbl(V)), "yyyy-mm-dd")
    Else
        DateTx = Trim$(CStr(V))
    End If
End Function

Private Function SpecOf(ws As Worksheet, ByVal r As Long) As Double
    Dim V As Variant
    V = ws.Range("F" & r).Value2
    If IsError(V) Or IsEmpty(V) Then
        SpecOf = 0
    ElseIf IsNumeric(V) Then
        SpecOf = CDbl(V)
    Else
        SpecOf = 0
    End If
End Function

Private Function SameKey(ws As Worksheet, ByVal a As Long, ByVal b As Long) As Boolean
    SameKey = (Tx(ws, a, "C") = Tx(ws, b, "C")) And _
              (DateTx(ws, "D" & a) = DateTx(ws, "D" & b)) And _
              (Tx(ws, a, "E") = Tx(ws, b, "E")) And _
              (Tx(ws, a, "F") = Tx(ws, b, "F"))
End Function

Private Function KeyOf(vi As VialInfo) As String
    Dim h As String
    h = vi.hosp
    If h = "" Then h = "(병원명 없음)"
    KeyOf = h & " / " & vi.dy & " / " & vi.spec & "mCi"
End Function

' 묶음 첫 행의 비고(M~P열) 번호를 읽어 A열 번호와 대조. 문제 없으면 "" 반환
' 예) 다행: "13 14 15 16" / 단일행: "5 1 Package (100mCi)" (뒤의 "1 Package"는 개수라 제외)
Private Function NumCheck(ws As Worksheet, ByVal r As Long, ByVal e As Long) As String
    Dim t As String, parts() As String, nums() As Long
    Dim i As Long, cnt As Long, n As Long, tok As String, a As String
    Dim c As Variant

    n = e - r + 1
    t = ""
    For Each c In Array("M", "N", "O", "P")
        t = t & " " & Tx(ws, r, CStr(c))
    Next c
    t = Replace(t, vbLf, " ")
    t = Replace(t, vbCr, " ")
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    t = Trim$(t)
    If t = "" Then
        NumCheck = "비고에 번호가 없음"
        Exit Function
    End If

    parts = Split(t, " ")
    ReDim nums(1 To UBound(parts) + 1)
    cnt = 0
    For i = 0 To UBound(parts)
        tok = parts(i)
        If tok Like "*[!0-9]*" Then
            If cnt > 0 Then
                If LCase$(Left$(tok, 7)) = "package" Or LCase$(Left$(tok, 4)) = "vial" Then cnt = cnt - 1
            End If
            Exit For
        End If
        cnt = cnt + 1
        nums(cnt) = CLng(tok)
    Next i

    If cnt <> n Then
        NumCheck = "비고 번호가 " & cnt & "개인데 묶인 행은 " & n & "행"
        Exit Function
    End If
    For i = 1 To cnt
        a = Tx(ws, r + i - 1, "A")
        If Val(a) <> nums(i) Then
            NumCheck = "비고 " & nums(i) & " / A열 " & a & " (행 " & (r + i - 1) & ")"
            Exit Function
        End If
    Next i
    NumCheck = ""
End Function

' 바이알 라벨의 (NmCi) 값을 읽음. 못 찾으면 -1
Private Function LabelMci(ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long) As Double
    Dim r As Long, c As Variant, t As String, p As Long, q As Long, s As String
    t = ""
    For r = r1 To r2
        For Each c In Array("L", "M", "N", "O", "P")
            t = t & " " & Tx(ws, r, CStr(c))
        Next c
    Next r
    t = Replace(t, vbLf, " ")
    t = Replace(t, vbCr, " ")
    p = InStr(1, t, "mCi", vbTextCompare)
    If p = 0 Then
        LabelMci = -1
        Exit Function
    End If
    q = p - 1
    Do While q > 0
        If Mid$(t, q, 1) = " " Then
            q = q - 1
        Else
            Exit Do
        End If
    Loop
    s = ""
    Do While q > 0
        If Mid$(t, q, 1) Like "[0-9.]" Then
            s = Mid$(t, q, 1) & s
            q = q - 1
        Else
            Exit Do
        End If
    Loop
    If s = "" Then
        LabelMci = -1
    Else
        LabelMci = Val(s)
    End If
End Function

Private Sub ParseSheet(ws As Worksheet, ByVal tag As String, V() As VialInfo, ByRef nV As Long, Pk() As PkgInfo, ByRef nP As Long)
    Dim lastR As Long, r As Long, e As Long, i As Long, k As Long, p As Long
    Dim covered() As Boolean
    Dim lb As Double, cnt As Long, bad As Boolean, vv As Variant, cols As Variant
    Dim ref As String, startsPkg As Boolean, hasPkgCol As Boolean, msgNum As String
    Dim c As Range, nm As Long
    Dim dPart As Object, kk As Variant, sKey As String

    ReDim V(1 To 400)
    ReDim Pk(1 To 400)
    nV = 0
    nP = 0
    lastR = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If lastR < 7 Then Exit Sub
    For r = 7 To lastR
        If InStr(1, Tx(ws, r, "A"), "Solution", vbTextCompare) > 0 Then
            lastR = r - 1
            Exit For
        End If
    Next r
    If lastR < 7 Then Exit Sub
    ReDim covered(1 To lastR)
    cols = Array("M", "N", "O", "P")

    r = 7
    Do While r <= lastR
        If Tx(ws, r, "L") = "No." Then
            e = r
            Do While e < lastR And e < r + 3
                If Tx(ws, e + 1, "L") = "No." Then Exit Do
                If Not SameKey(ws, r, e + 1) Then Exit Do
                If Tx(ws, e + 1, "H") = "" Then Exit Do
                e = e + 1
            Loop
            nV = nV + 1
            If nV > UBound(V) Then ReDim Preserve V(1 To UBound(V) + 400)
            V(nV).r1 = r
            V(nV).r2 = e
            V(nV).nRows = e - r + 1
            V(nV).hosp = Tx(ws, r, "C")
            V(nV).cal = DateTx(ws, "D" & r)
            V(nV).dy = Tx(ws, r, "E")
            V(nV).spec = SpecOf(ws, r)
            For i = r To e
                covered(i) = True
            Next i
            ref = "행 " & r & "~" & e

            ' 병원명이 비어 있는 행은 오차장 파일의 의도된 행(샘플/예비분)이라 요일 검사만 생략
            If V(nV).hosp <> "" Then
                If Len(V(nV).dy) <> 1 Or InStr(WK(), V(nV).dy) = 0 Then
                    AddRep tag, "오류", ref, "요일 값이 이상함: " & V(nV).dy, r, e
                End If
            End If

            lb = LabelMci(ws, r, e)
            If lb < 0 Then
                AddRep tag, "오류", ref, "mCi 표기를 읽지 못함 (CHECK)", r, e
            ElseIf Abs(lb - V(nV).nRows * V(nV).spec) > 0.0001 Then
                AddRep tag, "오류", ref, "mCi 표기 불일치: 표기 " & lb & "mCi, 계산 " & V(nV).nRows & "행 x " & V(nV).spec & " = " & (V(nV).nRows * V(nV).spec) & "mCi", r, e
                If tag = "선택 파일" Then AddMark "오류강조", r, e, "L", "P", ""
            End If

            msgNum = NumCheck(ws, r, e)
            If msgNum <> "" Then
                AddRep tag, "오류", ref, "비고 번호 오류: " & msgNum, r, e
                If tag = "선택 파일" Then AddMark "오류강조", r, r, "L", "P", ""
            End If
            r = e + 1
        Else
            r = r + 1
        End If
    Loop

    ' 어느 바이알에도 묶이지 않은 행
    For r = 7 To lastR
        If Not covered(r) Then
            If Tx(ws, r, "C") <> "" And Tx(ws, r, "H") <> "" Then
                If Len(Tx(ws, r, "E")) = 1 Then
                    If InStr(WK(), Tx(ws, r, "E")) > 0 Then
                        AddRep tag, "오류", "행 " & r, "어느 바이알에도 묶이지 않은 행 (바이알 안에서 병원/날짜/요일/규격이 바뀌었거나 5행 이상 묶임)", r, r
                    End If
                End If
            End If
        End If
    Next r

    ' 패키지 구성 (Q열에 Package가 있는 바이알이 새 패키지의 시작)
    hasPkgCol = False
    For i = 1 To nV
        If InStr(1, Tx(ws, V(i).r1, "Q"), "Package", vbTextCompare) > 0 Then
            hasPkgCol = True
            Exit For
        End If
    Next i

    For i = 1 To nV
        If hasPkgCol Then
            startsPkg = (InStr(1, Tx(ws, V(i).r1, "Q"), "Package", vbTextCompare) > 0)
        Else
            startsPkg = True
        End If
        If nP = 0 Or startsPkg Then
            nP = nP + 1
            If nP > UBound(Pk) Then ReDim Preserve Pk(1 To UBound(Pk) + 400)
            Pk(nP).v1 = i
            Pk(nP).r1 = V(i).r1
        End If
        Pk(nP).v2 = i
        Pk(nP).r2 = V(i).r2
    Next i

    For p = 1 To nP
        ref = "행 " & Pk(p).r1 & "~" & Pk(p).r2
        If Pk(p).v2 - Pk(p).v1 + 1 > 2 Then
            AddRep tag, "오류", ref, "패키지에 바이알이 " & (Pk(p).v2 - Pk(p).v1 + 1) & "개 (2개 초과)", Pk(p).r1, Pk(p).r2
        End If
        For i = Pk(p).v1 + 1 To Pk(p).v2
            If V(i).hosp <> V(Pk(p).v1).hosp Or V(i).cal <> V(Pk(p).v1).cal Or V(i).dy <> V(Pk(p).v1).dy Or V(i).spec <> V(Pk(p).v1).spec Then
                AddRep tag, "오류", ref, "한 패키지에 병원/날짜/요일/규격이 다른 바이알이 섞임", Pk(p).r1, Pk(p).r2
                Exit For
            End If
        Next i
        If hasPkgCol Then
            Set c = ws.Range("Q" & Pk(p).r1)
            If c.MergeCells Then
                nm = c.MergeArea.Rows.Count
            Else
                nm = 1
            End If
            If nm <> Pk(p).r2 - Pk(p).r1 + 1 Then
                AddRep tag, "오류", ref, "패키지 셀 병합이 " & nm & "행인데 실제 묶인 행은 " & (Pk(p).r2 - Pk(p).r1 + 1) & "행", Pk(p).r1, Pk(p).r2
            End If
        End If
    Next p

    ' 같은 병원/요일/규격에 4행 미만 바이알이 2개 이상
    Set dPart = CreateObject("Scripting.Dictionary")
    For i = 1 To nV
        If V(i).nRows < 4 And hasPkgCol Then
            sKey = KeyOf(V(i))
            If dPart.Exists(sKey) Then
                dPart(sKey) = dPart(sKey) + 1
            Else
                dPart.Add sKey, 1
            End If
        End If
    Next i
    For Each kk In dPart.Keys
        If dPart(kk) > 1 Then AddRep tag, "참고", CStr(kk), "4행 미만으로 남은 바이알이 " & dPart(kk) & "개 있음"
    Next kk
End Sub

' 병원/요일/규격별 행수, 묶는 방식(예: 4+4 , 4+1), 첫 날짜 요약
Private Sub Summ(V() As VialInfo, ByVal nV As Long, Pk() As PkgInfo, ByVal nP As Long, dRows As Object, dPat As Object, dCal As Object)
    Dim i As Long, p As Long, k As String, pat As String
    For i = 1 To nV
        k = KeyOf(V(i))
        If dRows.Exists(k) Then
            dRows(k) = dRows(k) + V(i).nRows
        Else
            dRows.Add k, V(i).nRows
            dCal.Add k, V(i).cal
        End If
    Next i
    For p = 1 To nP
        k = KeyOf(V(Pk(p).v1))
        pat = ""
        For i = Pk(p).v1 To Pk(p).v2
            If pat <> "" Then pat = pat & "+"
            pat = pat & V(i).nRows
        Next i
        If dPat.Exists(k) Then
            dPat(k) = dPat(k) & " , " & pat
        Else
            dPat.Add k, pat
        End If
    Next p
End Sub

Private Sub CompareFiles(vA() As VialInfo, ByVal nVA As Long, pA() As PkgInfo, ByVal nPA As Long, _
                         vB() As VialInfo, ByVal nVB As Long, pB() As PkgInfo, ByVal nPB As Long, ByVal sameWeek As Boolean, _
                         ByVal hasShift As Boolean, ByVal shiftDays As Long)
    Dim rA As Object, tA As Object, cA As Object, rB As Object, tB As Object, cB As Object
    Dim k As Variant
    Set rA = CreateObject("Scripting.Dictionary")
    Set tA = CreateObject("Scripting.Dictionary")
    Set cA = CreateObject("Scripting.Dictionary")
    Set rB = CreateObject("Scripting.Dictionary")
    Set tB = CreateObject("Scripting.Dictionary")
    Set cB = CreateObject("Scripting.Dictionary")
    Summ vA, nVA, pA, nPA, rA, tA, cA
    Summ vB, nVB, pB, nPB, rB, tB, cB

    For Each k In rA.Keys
        If InStr(1, CStr(k), "병원명 없음") > 0 Then
            ' 병원명 없는 행은 비교 제외
        ElseIf Not rB.Exists(k) Then
            AddRep "비교", DiffLevel(CStr(k)), CStr(k), "내 파일에만 있음 (" & rA(k) & "행)"
        ElseIf rA(k) <> rB(k) Then
            AddRep "비교", DiffLevel(CStr(k)), CStr(k), "행수가 다름: 내 파일 " & rA(k) & "행 / 선택 파일 " & rB(k) & "행"
        Else
            If tA(k) <> tB(k) Then
                AddRep "비교", DiffLevel(CStr(k)), CStr(k), "행수는 같으나 묶는 방식이 다름 - 내 파일: " & tA(k) & " / 선택 파일: " & tB(k)
            End If
            If sameWeek Then
                If cA(k) <> cB(k) Then AddRep "비교", DiffLevel(CStr(k)), CStr(k), "Cal 날짜가 다름: 내 파일 " & cA(k) & " / 선택 파일 " & cB(k)
            ElseIf hasShift Then
                If IsDate(cA(k)) And IsDate(cB(k)) Then
                    If CLng(CDate(cB(k))) - CLng(CDate(cA(k))) <> shiftDays Then
                        AddRep "비교", DiffLevel(CStr(k)), CStr(k), "Cal 날짜 이동폭이 발주일자 차이(" & shiftDays & "일)와 다름: 내 파일 " & cA(k) & " / 선택 파일 " & cB(k)
                    End If
                End If
            End If
        End If
    Next k
    For Each k In rB.Keys
        If InStr(1, CStr(k), "병원명 없음") > 0 Then
            ' 병원명 없는 행은 비교 제외
        ElseIf Not rA.Exists(k) Then
            AddRep "비교", DiffLevel(CStr(k)), CStr(k), "선택 파일에만 있음 (" & rB(k) & "행)"
        End If
    Next k
End Sub

Private Function DiffLevel(ByVal k As String) As String
    If InStr(1, k, "병원명 없음") > 0 Then
        DiffLevel = "참고"
    Else
        DiffLevel = "다름"
    End If
End Function

Private Function SumRows(V() As VialInfo, ByVal nV As Long) As Long
    Dim i As Long
    For i = 1 To nV
        SumRows = SumRows + V(i).nRows
    Next i
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

Public Sub CheckHospital()
    Dim fp As Variant, fpA As Variant, wbA As Workbook, wbB As Workbook
    Dim wsA As Worksheet, wsB As Worksheet, wsOut As Worksheet
    Dim openedA As Boolean, openedB As Boolean, oldSec As Long, nameA As String
    Dim vA() As VialInfo, vB() As VialInfo, pA() As PkgInfo, pB() As PkgInfo
    Dim nVA As Long, nVB As Long, nPA As Long, nPB As Long
    Dim dA As String, dB As String, nameB As String, sameWeek As Boolean
    Dim hasShift As Boolean, shiftDays As Long, outPath As String
    Dim i As Long, nErr As Long, nDiff As Long, nRef As Long
    Dim errMsg As String
    Dim outArr() As Variant

    fpA = Application.GetOpenFilename("Excel (*.xlsx;*.xlsm),*.xlsx;*.xlsm", , "1) 내 파일(빠로 xlsm)을 선택하세요")
    If VarType(fpA) = vbBoolean Then Exit Sub
    fp = Application.GetOpenFilename("Excel (*.xlsx;*.xlsm),*.xlsx;*.xlsm", , "2) 비교할 파일(오차장 파일)을 선택하세요")
    If VarType(fp) = vbBoolean Then Exit Sub
    If StrComp(CStr(fpA), CStr(fp), vbTextCompare) = 0 Then
        MsgBox "같은 파일을 두 번 선택했습니다.", vbExclamation
        Exit Sub
    End If
    If IsOpenFile(CStr(fp)) Then
        MsgBox "오차장 파일이 이미 열려 있습니다. 닫고 다시 실행해 주세요.", vbExclamation
        Exit Sub
    End If

    oldSec = Application.AutomationSecurity
    On Error GoTo EH
    Application.ScreenUpdating = False
    Application.AutomationSecurity = 3   ' 파일을 열 때 매크로를 실행하지 않음
    Set wbA = OpenReadOnly(CStr(fpA), openedA)
    Set wbB = OpenReadOnly(CStr(fp), openedB)
    Application.AutomationSecurity = oldSec
    Set wsA = wbA.Worksheets(SN())
    Set wsB = wbB.Worksheets(SN())
    nameA = wbA.Name
    nameB = wbB.Name

    ReDim rep(1 To 4, 1 To 500)
    repN = 0
    ReDim mkLv(1 To 200)
    ReDim mkR1(1 To 200)
    ReDim mkR2(1 To 200)
    ReDim mkC1(1 To 200)
    ReDim mkC2(1 To 200)
    ReDim mkMsg(1 To 200)
    mkN = 0

    ParseSheet wsA, "내 파일", vA, nVA, pA, nPA
    ParseSheet wsB, "선택 파일", vB, nVB, pB, nPB
    dA = DateTx(wsA, "K3")
    dB = DateTx(wsB, "K3")
    sameWeek = (dA = dB)
    hasShift = False
    If Not sameWeek Then
        If Not IsEmpty(wsA.Range("K3").Value2) And Not IsEmpty(wsB.Range("K3").Value2) Then
            If IsNumeric(wsA.Range("K3").Value2) And IsNumeric(wsB.Range("K3").Value2) Then
                hasShift = True
                shiftDays = CLng(wsB.Range("K3").Value2) - CLng(wsA.Range("K3").Value2)
            End If
        End If
    End If
    CompareFiles vA, nVA, pA, nPA, vB, nVB, pB, nPB, sameWeek, hasShift, shiftDays
    For i = 1 To repN
        If rep(1, i) = "비교" And rep(2, i) = "다름" Then MarkKey vB, nVB, rep(3, i), rep(4, i)
    Next i
    outPath = SaveMarked(wsB, wbB, CStr(fp))

    For i = 1 To repN
        Select Case rep(2, i)
            Case "오류": nErr = nErr + 1
            Case "다름": nDiff = nDiff + 1
            Case Else: nRef = nRef + 1
        End Select
    Next i

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("검증결과").Delete
    On Error GoTo EH
    Application.DisplayAlerts = True
    Set wsOut = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsOut.Name = "검증결과"

    wsOut.Range("A1").Value = "병원 시트 검증 결과"
    wsOut.Range("A1").Font.Bold = True
    wsOut.Range("A1").Font.Size = 14
    wsOut.Range("A2").Value = "내 파일: " & nameA & " / 발주일자 " & dA & " / 바이알 " & nVA & " / 패키지 " & nPA & " / 행 " & SumRows(vA, nVA)
    wsOut.Range("A3").Value = "선택 파일: " & nameB & " / 발주일자 " & dB & " / 바이알 " & nVB & " / 패키지 " & nPB & " / 행 " & SumRows(vB, nVB)
    wsOut.Range("A4").Value = "오류 " & nErr & "건 / 다름 " & nDiff & "건 / 참고 " & nRef & "건"
    If Not sameWeek Then
        If hasShift Then
            wsOut.Range("A5").Value = "참고: 두 파일의 발주일자가 " & shiftDays & "일 차이라, Cal 날짜는 같은 일수만큼 이동했는지로 비교했습니다."
        Else
            wsOut.Range("A5").Value = "주의: 두 파일의 발주일자가 달라 Cal 날짜는 비교하지 않았습니다."
        End If
    End If
    wsOut.Range("A4").Font.Bold = True
    If outPath <> "" Then
        wsOut.Range("A6").Value = "표시본 저장: " & outPath & " (연한 빨강=규칙 오류, 진한 빨강=mCi 표기 칸, 노랑=내 파일과 다름, 이유는 셀 메모)"
    Else
        wsOut.Range("A6").Value = "표시본을 저장하지 못했습니다 (폴더 쓰기 권한을 확인하세요)"
    End If

    wsOut.Range("A7:D7").Value = Array("대상", "구분", "위치", "내용")
    wsOut.Range("A7:D7").Font.Bold = True
    wsOut.Range("A7:D7").Interior.Color = RGB(217, 217, 217)

    If repN > 0 Then
        ReDim outArr(1 To repN, 1 To 4)
        For i = 1 To repN
            outArr(i, 1) = rep(1, i)
            outArr(i, 2) = rep(2, i)
            outArr(i, 3) = rep(3, i)
            outArr(i, 4) = rep(4, i)
            If rep(2, i) = "오류" Then
                wsOut.Range("A" & (7 + i) & ":D" & (7 + i)).Font.Color = RGB(192, 0, 0)
            End If
        Next i
        wsOut.Range("A8").Resize(repN, 4).Value = outArr
        wsOut.Range("A7:D" & (7 + repN)).AutoFilter
    Else
        wsOut.Range("A8").Value = "발견된 문제가 없습니다."
    End If
    wsOut.Columns("A").ColumnWidth = 12
    wsOut.Columns("B").ColumnWidth = 8
    wsOut.Columns("C").ColumnWidth = 28
    wsOut.Columns("D").ColumnWidth = 90

    wbB.Close SaveChanges:=False
    If openedA Then wbA.Close SaveChanges:=False
    Application.ScreenUpdating = True
    wsOut.Activate
    MsgBox "검증 완료: 오류 " & nErr & "건, 다름 " & nDiff & "건, 참고 " & nRef & "건" & vbCrLf & "표시본: " & outPath, vbInformation
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

