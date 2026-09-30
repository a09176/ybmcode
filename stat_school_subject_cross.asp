<%@ LANGUAGE="VBScript" %>
<%
Option Explicit
Response.Expires = -1

Const iMenuIdex = 611
Const current_menu_num = 75
%>

<!-- #include virtual = "/teacher_inc/definition.asp" -->
<!-- #include virtual = "/teacher_inc/function.asp" -->
<!-- #include virtual = "/lms/inc/admin_config.asp" -->
<!-- #include virtual = "/lms/inc/chkAdmin.asp" -->
<!-- #include virtual = "/v15/include/fnDB.asp" -->
<!-- #include virtual = "/common/lib/FNC_NonInjection.asp" -->

<%
Call AdminLoginCheck()
Call fnAdminPermission(Session("admin_group"), current_menu_num, False)

'==========================================================
' 변수 선언
'==========================================================
Dim strSQL
Dim arrPayData
Dim arrTeacherData
Dim arrSchoolData

Dim i
Dim j
Dim t
Dim batchStart
Dim batchEnd

Dim sYear
Dim sMonth
Dim eYear
Dim eMonth
Dim sDay
Dim eDay

Dim startDate
Dim endDate
Dim downloadDate

Dim userid
Dim schoolDiv
Dim subjectName
Dim buyerCount
Dim actualAmount

Dim teacherMap
Dim schoolMap
Dim cellMap
Dim subjectCells
Dim schoolKeyMap
Dim subjectKeyMap

Dim userKeys
Dim schoolKeys
Dim subjectKeys

Dim cellData
Dim cellValue
Dim tableTitle
Dim rowSum
Dim colSum
Dim grandSum

'==========================================================
' 조회 기간
'==========================================================
sYear = Trim(Request("sYear"))
sMonth = Trim(Request("sMonth"))
eYear = Trim(Request("eYear"))
eMonth = Trim(Request("eMonth"))
sDay = Trim(Request("sDay"))
eDay = Trim(Request("eDay"))

If sYear = "" Or Not IsNumeric(sYear) Then sYear = "2025"
If sMonth = "" Or Not IsNumeric(sMonth) Then sMonth = "01"
If eYear = "" Or Not IsNumeric(eYear) Then eYear = CStr(Year(Now()))
If eMonth = "" Or Not IsNumeric(eMonth) Then eMonth = CStr(Month(Now()))
If sDay = "" Or Not IsNumeric(sDay) Then sDay = "01"
If eDay = "" Or Not IsNumeric(eDay) Then eDay = CStr(Day(Now()))

startDate = sYear & "-" & Right("0" & sMonth, 2) & "-" & Right("0" & sDay, 2)
endDate = eYear & "-" & Right("0" & eMonth, 2) & "-" & Right("0" & eDay, 2)

downloadDate = Year(Now()) & _
               Right("0" & Month(Now()), 2) & _
               Right("0" & Day(Now()), 2)

'==========================================================
' 1. 기간 전체의 회원별 결제액 조회
'
' 결과는 userid별 한 행이다.
' 결제건수는 구매자수로 사용하지 않는다.
'==========================================================
strSQL = ""
strSQL = strSQL & "SELECT "
strSQL = strSQL & "    p.userid, "
strSQL = strSQL & "    SUM(ISNULL(p.pay_price, 0)) AS actual_amount "
strSQL = strSQL & "FROM pay_info AS p WITH (READUNCOMMITTED) "
strSQL = strSQL & "WHERE p.sell_info = 'o' "
strSQL = strSQL & "  AND p.pay_date >= '" & startDate & "' "
strSQL = strSQL & "  AND p.pay_date <= '" & endDate & "' "
strSQL = strSQL & "GROUP BY p.userid "
strSQL = strSQL & "ORDER BY p.userid "

arrPayData = ExecSP(strSQL, strCyberTeacher)

'==========================================================
' 2. 조회 대상 회원 목록
'==========================================================
Set teacherMap = CreateObject("Scripting.Dictionary")
Set schoolMap = CreateObject("Scripting.Dictionary")

If IsArray(arrPayData) Then

    For i = 0 To UBound(arrPayData, 2)

        userid = Trim(Nz(arrPayData(0, i), "") & "")

        If userid <> "" Then
            If Not teacherMap.Exists(userid) Then
                teacherMap.Add userid, "미분류"
                schoolMap.Add userid, "미분류"
            End If
        End If

    Next

End If

'==========================================================
' 3. 담당과목과 학교급 조회
'
' IN 절이 지나치게 커지지 않도록 500명씩 나눈다.
' 담당과목: vTeacher_info.subject
' 학교급: vUser_info.school_num
'         -> school_info.school_division
'==========================================================
If teacherMap.Count > 0 Then

    userKeys = teacherMap.Keys

    For batchStart = 0 To UBound(userKeys) Step 500

        batchEnd = batchStart + 499

        If batchEnd > UBound(userKeys) Then
            batchEnd = UBound(userKeys)
        End If

        '--------------------------------------------------
        ' 담당과목
        '
        ' 한 회원에 과목 행이 여러 개라면 과목 하나를
        ' 결정적으로 선택하여 구매자 중복을 방지한다.
        '--------------------------------------------------
        strSQL = ""
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "    userid, "
        strSQL = strSQL & "    MIN(NULLIF(LTRIM(RTRIM(subject)), '')) AS subject "
        strSQL = strSQL & "FROM vTeacher_info WITH (READUNCOMMITTED) "
        strSQL = strSQL & "WHERE userid IN ("

        For j = batchStart To batchEnd
            If j > batchStart Then strSQL = strSQL & ","
            strSQL = strSQL & fnCdbValue(userKeys(j))
        Next

        strSQL = strSQL & ") "
        strSQL = strSQL & "GROUP BY userid "

        arrTeacherData = ExecSP(strSQL, strCyberTeacher)

        If IsArray(arrTeacherData) Then

            For i = 0 To UBound(arrTeacherData, 2)

                userid = Trim(Nz(arrTeacherData(0, i), "") & "")
                subjectName = Trim(Nz(arrTeacherData(1, i), "") & "")

                If subjectName = "" Then
                    subjectName = "미분류"
                End If

                If userid <> "" Then
                    If teacherMap.Exists(userid) Then
                        teacherMap(userid) = subjectName
                    End If
                End If

            Next

        End If

        '--------------------------------------------------
        ' 학교급
        '
        ' 한 회원의 학교 정보가 여러 행으로 조회되더라도
        ' 학교급 코드 하나만 선택한다.
        '--------------------------------------------------
        strSQL = ""
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "    u.userid, "
        strSQL = strSQL & "    CONVERT(VARCHAR(20), MIN(si.school_division)) AS school_division "
        strSQL = strSQL & "FROM vUser_info AS u WITH (READUNCOMMITTED) "
        strSQL = strSQL & "LEFT JOIN school_info AS si WITH (READUNCOMMITTED) "
        strSQL = strSQL & "    ON u.school_num = si.school_num "
        strSQL = strSQL & "WHERE u.userid IN ("

        For j = batchStart To batchEnd
            If j > batchStart Then strSQL = strSQL & ","
            strSQL = strSQL & fnCdbValue(userKeys(j))
        Next

        strSQL = strSQL & ") "
        strSQL = strSQL & "GROUP BY u.userid "

        arrSchoolData = ExecSP(strSQL, strCyberTeacher)

        If IsArray(arrSchoolData) Then

            For i = 0 To UBound(arrSchoolData, 2)

                userid = Trim(Nz(arrSchoolData(0, i), "") & "")
                schoolDiv = Trim(Nz(arrSchoolData(1, i), "") & "")

                If schoolDiv = "" Then
                    schoolDiv = "미분류"
                End If

                If userid <> "" Then
                    If schoolMap.Exists(userid) Then
                        schoolMap(userid) = schoolDiv
                    End If
                End If

            Next

        End If

    Next

End If

'==========================================================
' 4. 학교급 × 담당과목 교차표 구성
'
' cellMap(학교급코드)(담당과목)
'     = Array(구매자수, 실제결제액)
'
' arrPayData는 userid별 한 행이므로 유효한 userid마다
' 구매자수를 정확히 1씩 더한다.
'==========================================================
Set cellMap = CreateObject("Scripting.Dictionary")
Set schoolKeyMap = CreateObject("Scripting.Dictionary")
Set subjectKeyMap = CreateObject("Scripting.Dictionary")

If IsArray(arrPayData) Then

    For i = 0 To UBound(arrPayData, 2)

        userid = Trim(Nz(arrPayData(0, i), "") & "")
        actualAmount = CDbl(Nz(arrPayData(1, i), 0))

        schoolDiv = "미분류"
        subjectName = "미분류"
        buyerCount = 0

        If userid <> "" Then

            buyerCount = 1

            If schoolMap.Exists(userid) Then
                schoolDiv = Trim(Nz(schoolMap(userid), "") & "")
                If schoolDiv = "" Then schoolDiv = "미분류"
            End If

            If teacherMap.Exists(userid) Then
                subjectName = Trim(Nz(teacherMap(userid), "") & "")
                If subjectName = "" Then subjectName = "미분류"
            End If

        End If

        If Not schoolKeyMap.Exists(schoolDiv) Then
            schoolKeyMap.Add schoolDiv, 1
        End If

        If Not subjectKeyMap.Exists(subjectName) Then
            subjectKeyMap.Add subjectName, 1
        End If

        If Not cellMap.Exists(schoolDiv) Then
            Set subjectCells = CreateObject("Scripting.Dictionary")
            cellMap.Add schoolDiv, subjectCells
        End If

        Set subjectCells = cellMap(schoolDiv)

        If subjectCells.Exists(subjectName) Then

            cellData = subjectCells(subjectName)
            cellData(0) = CDbl(cellData(0)) + buyerCount
            cellData(1) = CDbl(cellData(1)) + actualAmount
            subjectCells(subjectName) = cellData

        Else

            subjectCells.Add subjectName, Array( _
                CDbl(buyerCount), _
                actualAmount _
            )

        End If

    Next

End If

'==========================================================
' 5. 행·열 정렬
'==========================================================
If schoolKeyMap.Count > 0 Then
    schoolKeys = schoolKeyMap.Keys
    Call SortCategoryKeys(schoolKeys, "school")
End If

If subjectKeyMap.Count > 0 Then
    subjectKeys = subjectKeyMap.Keys
    Call SortCategoryKeys(subjectKeys, "subject")
End If

'==========================================================
' 개인정보 파일 다운로드 로그
'==========================================================
If IsArray(arrPayData) Then

    strSQL = ""
    strSQL = strSQL & "EXEC site_log.dbo.p_File_Log "
    strSQL = strSQL & fnCdbValue(Session("admin_id")) & ","
    strSQL = strSQL & fnCdbValue(Request.ServerVariables("REMOTE_ADDR")) & ","
    strSQL = strSQL & fnCdbValue(Request.ServerVariables("HTTP_HOST")) & ","
    strSQL = strSQL & fnCdbValue("stat_school_subject_cross_" & downloadDate) & ","
    strSQL = strSQL & fnCdbValue(Request.ServerVariables("URL")) & ","
    strSQL = strSQL & fnCdbValue(Session("admin_name")) & ","
    strSQL = strSQL & "'',"
    strSQL = strSQL & fnCdbValue("") & ","
    strSQL = strSQL & "'' "

    Call ExecSQL(strSQL, e4u2006DBStr)

End If

'==========================================================
' Excel 다운로드 설정
'==========================================================
Response.ContentType = "application/vnd.ms-excel"
Response.CharSet = "euc-kr"

Response.AddHeader _
    "Content-Disposition", _
    "attachment; filename=stat_school_subject_cross_" & _
    downloadDate & ".xls"
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
</head>

<body>

<!-- ================= 조회 기간 정보 ================= -->
<div style="margin-bottom: 30px;">
    <h3>학교급 × 담당과목 교차분석</h3>
    <p>
        <strong>조회 기간:</strong> 
        <%=Server.HTMLEncode(startDate)%> ~ <%=Server.HTMLEncode(endDate)%>
    </p>
    <p style="color: #666; font-size: 12px;">
        기간 내 결제한 회원을 학교급과 담당과목으로 분류하여 집계합니다.<br>
        각 회원은 정확히 한 번만 계산되며, 구매자수와 실제결제액은 아래 두 테이블에서 각각 확인할 수 있습니다.
    </p>
</div>

<%
If schoolKeyMap.Count = 0 Or subjectKeyMap.Count = 0 Then

    Response.Write "<table border=""1"">" & _
                   "<tr><td>조회 결과가 없습니다.</td></tr>" & _
                   "</table>"

Else

    ' t = 0: 구매자수
    ' t = 1: 실제결제액
    For t = 0 To 1

        If t = 0 Then
            tableTitle = "학교급 × 담당과목 구매자수 (명)"
        Else
            tableTitle = "학교급 × 담당과목 실제결제액 (원)"
        End If

        ' 테이블 구분선
        If t > 0 Then
            Response.Write "<hr style=""margin: 40px 0; border: 1px solid #999;"">" & vbCrLf
        End If

        Response.Write "<table border=""1"" style=""margin-top: 20px;"">" & vbCrLf

        ' 제목 및 설명
        Response.Write "<tr><td colspan=""" & _
                       (UBound(subjectKeys) + 3) & _
                       """ style=""font-weight:bold; " & _
                       "font-size:14px; background-color:#D3D3D3;"">" & _
                       Server.HTMLEncode(tableTitle) & _
                       "</td></tr>" & vbCrLf

        ' 설명 행
        Response.Write "<tr><td colspan=""" & _
                       (UBound(subjectKeys) + 3) & _
                       """ style=""background-color:#F5F5F5; " & _
                       "font-size:11px; color:#666; padding:10px;"">"

        If t = 0 Then
            Response.Write "기간 내 결제한 회원 중 학교급별, 담당과목별 구매자 수입니다. " & _
                           "각 회원은 정확히 한 번만 계산됩니다."
        Else
            Response.Write "기간 내 결제한 회원들의 학교급별, 담당과목별 총 결제액입니다. " & _
                           "한 회원이 여러 번 결제한 경우 모든 결제액이 합산됩니다."
        End If

        Response.Write "</td></tr>" & vbCrLf

        ' 헤더
        Response.Write "<tr style=""background-color:#4472C4; " & _
                       "color:white; font-weight:bold;"">" & vbCrLf

        Response.Write "<th width=""120"">학교급 \ 담당과목</th>" & vbCrLf

        For j = 0 To UBound(subjectKeys)
            Response.Write "<th width=""100"">" & _
                           Server.HTMLEncode(subjectKeys(j) & "") & _
                           "</th>" & vbCrLf
        Next

        Response.Write "<th width=""120"">행합계</th>" & vbCrLf
        Response.Write "</tr>" & vbCrLf

        grandSum = 0

        For i = 0 To UBound(schoolKeys)

            rowSum = 0

            Response.Write "<tr>" & vbCrLf
            Response.Write "<td align=""center"" " & _
                           "style=""mso-number-format:'\@'; " & _
                           "font-weight:bold;"">" & _
                           Server.HTMLEncode( _
                               GetSchoolDivisionName(schoolKeys(i)) _
                           ) & _
                           "</td>" & vbCrLf

            Set subjectCells = cellMap(schoolKeys(i))

            For j = 0 To UBound(subjectKeys)

                cellValue = 0

                If subjectCells.Exists(subjectKeys(j)) Then
                    cellData = subjectCells(subjectKeys(j))
                    cellValue = CDbl(cellData(t))
                End If

                rowSum = rowSum + cellValue

                Response.Write "<td align=""right"" " & _
                               "style=""mso-number-format:'\#\,\#\#0';"">" & _
                               FormatNumber(cellValue, 0) & _
                               "</td>" & vbCrLf

            Next

            grandSum = grandSum + rowSum

            Response.Write "<td align=""right"" " & _
                           "style=""mso-number-format:'\#\,\#\#0'; " & _
                           "background-color:#E7E6E6; " & _
                           "font-weight:bold;"">" & _
                           FormatNumber(rowSum, 0) & _
                           "</td>" & vbCrLf

            Response.Write "</tr>" & vbCrLf

        Next

        Response.Write "<tr style=""background-color:#E7E6E6; " & _
                       "font-weight:bold;"">" & vbCrLf
        Response.Write "<td align=""center"">열합계</td>" & vbCrLf

        For j = 0 To UBound(subjectKeys)

            colSum = 0

            For i = 0 To UBound(schoolKeys)

                Set subjectCells = cellMap(schoolKeys(i))

                If subjectCells.Exists(subjectKeys(j)) Then
                    cellData = subjectCells(subjectKeys(j))
                    colSum = colSum + CDbl(cellData(t))
                End If

            Next

            Response.Write "<td align=""right"" " & _
                           "style=""mso-number-format:'\#\,\#\#0';"">" & _
                           FormatNumber(colSum, 0) & _
                           "</td>" & vbCrLf

        Next

        Response.Write "<td align=""right"" " & _
                       "style=""mso-number-format:'\#\,\#\#0';"">" & _
                       FormatNumber(grandSum, 0) & _
                       "</td>" & vbCrLf

        Response.Write "</tr>" & vbCrLf
        Response.Write "</table>" & vbCrLf

    Next

End If
%>

</body>
</html>

<%
'==========================================================
' 키 배열 정렬
'==========================================================
Sub SortCategoryKeys(ByRef arrKeys, ByVal mode)

    Dim a
    Dim b
    Dim tempKey

    If Not IsArray(arrKeys) Then Exit Sub
    If UBound(arrKeys) < 1 Then Exit Sub

    For a = 0 To UBound(arrKeys) - 1

        For b = a + 1 To UBound(arrKeys)

            If CompareCategoryKeys(arrKeys(a), arrKeys(b), mode) > 0 Then
                tempKey = arrKeys(a)
                arrKeys(a) = arrKeys(b)
                arrKeys(b) = tempKey
            End If

        Next

    Next

End Sub

'==========================================================
' 학교급: 코드 숫자 오름차순, 미분류 마지막
' 담당과목: 가나다순, 미분류 마지막
'==========================================================
Function CompareCategoryKeys(ByVal leftValue, ByVal rightValue, ByVal mode)

    Dim leftText
    Dim rightText

    leftText = Trim(leftValue & "")
    rightText = Trim(rightValue & "")

    If leftText = "미분류" Then

        If rightText = "미분류" Then
            CompareCategoryKeys = 0
        Else
            CompareCategoryKeys = 1
        End If

        Exit Function

    End If

    If rightText = "미분류" Then
        CompareCategoryKeys = -1
        Exit Function
    End If

    If mode = "school" Then

        If IsNumeric(leftText) And IsNumeric(rightText) Then

            If CDbl(leftText) > CDbl(rightText) Then
                CompareCategoryKeys = 1
            ElseIf CDbl(leftText) < CDbl(rightText) Then
                CompareCategoryKeys = -1
            Else
                CompareCategoryKeys = 0
            End If

            Exit Function

        End If

    End If

    CompareCategoryKeys = StrComp(leftText, rightText, 1)

End Function

'==========================================================
' 확인된 학교급 코드명
'==========================================================
Function GetSchoolDivisionName(ByVal code)

    Select Case Trim(code & "")

        Case "1":   GetSchoolDivisionName = "중학교"
        Case "2":   GetSchoolDivisionName = "기타학교"
        Case "3":   GetSchoolDivisionName = "유치원"
        Case "4":   GetSchoolDivisionName = "고등학교"
        Case "5":   GetSchoolDivisionName = "기타"
        Case "6":   GetSchoolDivisionName = "맹아학교"
        Case "7":   GetSchoolDivisionName = "초등학교"
        Case "8":   GetSchoolDivisionName = "지역교육청"
        Case "9":   GetSchoolDivisionName = "농아학교"
        Case "10":  GetSchoolDivisionName = "시도교육청"
        Case "241": GetSchoolDivisionName = "교과부"
        Case "248": GetSchoolDivisionName = "특수학교"
        Case "276": GetSchoolDivisionName = "대안학교"

        Case "미분류"
            GetSchoolDivisionName = "미분류"

        Case Else
            GetSchoolDivisionName = code & ""

    End Select

End Function

'==========================================================
' NULL·빈 값 처리
'==========================================================
Function Nz(ByVal value, ByVal defaultValue)

    If IsNull(value) Then
        Nz = defaultValue
    ElseIf IsEmpty(value) Then
        Nz = defaultValue
    ElseIf Trim(value & "") = "" Then
        Nz = defaultValue
    Else
        Nz = value
    End If

End Function
%>