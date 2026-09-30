
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

Dim strSQL
Dim arrPay
Dim arrTeacher
Dim arrSchool

Dim i
Dim j

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
Dim payMonth
Dim subjectName
Dim schoolDivision

Dim payCount
Dim buyerCount
Dim actualAmount
Dim averageAmount

Dim teacherMap
Dim schoolMap
Dim resultMap

Dim userKeys
Dim resultKeys
Dim resultKey

Dim keyPartsA
Dim keyPartsB
Dim keyTemp
Dim sortA
Dim sortB

Dim rowData

Dim totalPayCount
Dim totalBuyerCount
Dim totalActual

'==========================================================
' 조회 기간
'==========================================================
sYear = Trim(Request("sYear"))
sMonth = Trim(Request("sMonth"))
eYear = Trim(Request("eYear"))
eMonth = Trim(Request("eMonth"))
sDay = Trim(Request("sDay"))
eDay = Trim(Request("eDay"))

If sYear = "" Or Not IsNumeric(sYear) Then
    sYear = "2025"
End If

If sMonth = "" Or Not IsNumeric(sMonth) Then
    sMonth = "01"
End If

If eYear = "" Or Not IsNumeric(eYear) Then
    eYear = CStr(Year(Now()))
End If

If eMonth = "" Or Not IsNumeric(eMonth) Then
    eMonth = CStr(Month(Now()))
End If

If sDay = "" Or Not IsNumeric(sDay) Then
    sDay = "01"
End If

If eDay = "" Or Not IsNumeric(eDay) Then
    eDay = CStr(Day(Now()))
End If

startDate = sYear & "-" & Right("0" & sMonth, 2) & "-" & Right("0" & sDay, 2)
endDate = eYear & "-" & Right("0" & eMonth, 2) & "-" & Right("0" & eDay, 2)

downloadDate = Year(Now()) & _
               Right("0" & Month(Now()), 2) & _
               Right("0" & Day(Now()), 2)

'==========================================================
' 1. 월별·회원별 결제 기본 데이터 조회
'
' 회원별로 먼저 집계한 뒤 담당과목·학교급을 매핑한다.
' 이렇게 해야 한 회원이 여러 결제를 한 경우 구매자 수가
' 결제 건수만큼 중복되지 않는다.
'==========================================================
strSQL = ""

strSQL = strSQL & "SELECT "
strSQL = strSQL & "    LEFT(p.pay_date, 7) AS pay_month, "
strSQL = strSQL & "    p.userid, "
strSQL = strSQL & "    COUNT(DISTINCT p.pay_num) AS pay_count, "
strSQL = strSQL & "    SUM(ISNULL(p.pay_price, 0)) AS actual_amount "

strSQL = strSQL & "FROM pay_info AS p WITH (READUNCOMMITTED) "

strSQL = strSQL & "WHERE p.sell_info = 'o' "
strSQL = strSQL & "  AND p.pay_date >= '" & startDate & "' "
strSQL = strSQL & "  AND p.pay_date < DATEADD(DAY, 1, '" & endDate & "') "

strSQL = strSQL & "GROUP BY "
strSQL = strSQL & "    LEFT(p.pay_date, 7), "
strSQL = strSQL & "    p.userid "

strSQL = strSQL & "ORDER BY "
strSQL = strSQL & "    LEFT(p.pay_date, 7), "
strSQL = strSQL & "    p.userid "

arrPay = ExecSP(strSQL, strCyberTeacher)

'==========================================================
' 2. 담당과목 조회
'
' vTeacher_info는 페이지에서 조회
'==========================================================
Set teacherMap = CreateObject("Scripting.Dictionary")

If IsArray(arrPay) Then

    For i = 0 To UBound(arrPay, 2)

        userid = Nz(arrPay(1, i), "")

        If userid <> "" Then

            If Not teacherMap.Exists(userid) Then
                teacherMap.Add userid, "미분류"
            End If

        End If

    Next

End If

If teacherMap.Count > 0 Then

    strSQL = ""
    strSQL = strSQL & "SELECT userid, MIN(NULLIF(LTRIM(RTRIM(subject)), '')) AS subject "
    strSQL = strSQL & "FROM vTeacher_info WITH (READUNCOMMITTED) "
    strSQL = strSQL & "WHERE userid IN ("

    userKeys = teacherMap.Keys

    For i = 0 To UBound(userKeys)

        If i > 0 Then
            strSQL = strSQL & ","
        End If

        strSQL = strSQL & fnCdbValue(userKeys(i))

    Next

    strSQL = strSQL & ") "
    strSQL = strSQL & "GROUP BY userid "

    arrTeacher = ExecSP(strSQL, strCyberTeacher)

    If IsArray(arrTeacher) Then

        For i = 0 To UBound(arrTeacher, 2)

            userid = Nz(arrTeacher(0, i), "")
            subjectName = Nz(arrTeacher(1, i), "미분류")

            If userid <> "" Then

                If Trim(subjectName & "") = "" Then
                    subjectName = "미분류"
                End If

                If teacherMap.Exists(userid) Then
                    teacherMap(userid) = subjectName
                End If

            End If

        Next

    End If

End If

'==========================================================
' 3. 학교급 조회
'
' vUser_info.school_num
'      -> school_info.school_num
'      -> school_info.school_division
'
' school_division 값은 원본 코드값을 표시한다.
'==========================================================
Set schoolMap = CreateObject("Scripting.Dictionary")

If IsArray(arrPay) Then

    For i = 0 To UBound(arrPay, 2)

        userid = Nz(arrPay(1, i), "")

        If userid <> "" Then

            If Not schoolMap.Exists(userid) Then
                schoolMap.Add userid, "미분류"
            End If

        End If

    Next

End If

If schoolMap.Count > 0 Then

    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "    u.userid, "
    strSQL = strSQL & "    ISNULL(CONVERT(VARCHAR(20), MIN(si.school_division)), '미분류') AS school_division "
    strSQL = strSQL & "FROM vUser_info AS u WITH (READUNCOMMITTED) "
    strSQL = strSQL & "LEFT JOIN school_info AS si WITH (READUNCOMMITTED) "
    strSQL = strSQL & "    ON u.school_num = si.school_num "
    strSQL = strSQL & "WHERE u.userid IN ("

    userKeys = schoolMap.Keys

    For i = 0 To UBound(userKeys)

        If i > 0 Then
            strSQL = strSQL & ","
        End If

        strSQL = strSQL & fnCdbValue(userKeys(i))

    Next

    strSQL = strSQL & ") "
    strSQL = strSQL & "GROUP BY u.userid "

    arrSchool = ExecSP(strSQL, strCyberTeacher)

    If IsArray(arrSchool) Then

        For i = 0 To UBound(arrSchool, 2)

            userid = Nz(arrSchool(0, i), "")
            schoolDivision = Nz(arrSchool(1, i), "미분류")

            If Trim(schoolDivision & "") = "" Then
                schoolDivision = "미분류"
            End If

            If userid <> "" Then

                If schoolMap.Exists(userid) Then
                    schoolMap(userid) = schoolDivision
                End If

            End If

        Next

    End If

End If

'==========================================================
' 4. 담당과목·학교급별 결과 누적
'==========================================================
Set resultMap = CreateObject("Scripting.Dictionary")

If IsArray(arrPay) Then

    For i = 0 To UBound(arrPay, 2)

        payMonth = Nz(arrPay(0, i), "")
        userid = Nz(arrPay(1, i), "")

        payCount = CLng(Nz(arrPay(2, i), 0))
        actualAmount = CLng(Nz(arrPay(3, i), 0))

        subjectName = "미분류"
        schoolDivision = "미분류"

        If teacherMap.Exists(userid) Then
            subjectName = Nz(teacherMap(userid), "미분류")
        End If

        If schoolMap.Exists(userid) Then
            schoolDivision = Nz(schoolMap(userid), "미분류")
        End If

        If Trim(subjectName & "") = "" Then
            subjectName = "미분류"
        End If

        If Trim(schoolDivision & "") = "" Then
            schoolDivision = "미분류"
        End If

        resultKey = payMonth & "|" & subjectName & "|" & schoolDivision

        If Not resultMap.Exists(resultKey) Then

            resultMap.Add resultKey, Array( _
                payMonth, _
                subjectName, _
                schoolDivision, _
                payCount, _
                actualAmount, _
                1 _
            )

        Else

            rowData = resultMap(resultKey)

            rowData(3) = rowData(3) + payCount
            rowData(4) = rowData(4) + actualAmount
            rowData(5) = rowData(5) + 1

            resultMap(resultKey) = rowData

        End If

    Next

End If

'==========================================================
' 개인정보 파일 다운로드 로그
'==========================================================
If IsArray(arrPay) Then

    strSQL = ""
    strSQL = strSQL & "EXEC site_log.dbo.p_File_Log "
    strSQL = strSQL & fnCdbValue(Session("admin_id")) & ","
    strSQL = strSQL & fnCdbValue(Request.ServerVariables("REMOTE_ADDR")) & ","
    strSQL = strSQL & fnCdbValue(Request.ServerVariables("HTTP_HOST")) & ","
    strSQL = strSQL & fnCdbValue("stat_subject_analysis_" & downloadDate) & ","
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
    "attachment; filename=stat_subject_analysis_" & _
    downloadDate & ".xls"
%>

<html>
<head>
<meta http-equiv="Content-Type"
      content="text/html; charset=euc-kr">
</head>

<body>

<table border="1">
    <thead>
        <tr style="background-color: #4472C4; color: white; font-weight: bold;">
            <th width="100">연월</th>
            <th width="120">담당과목</th>
            <th width="120">학교급</th>
            <th width="120">결제건수</th>
            <th width="150">구매자수</th>
            <th width="150">실제결제액</th>
            <th width="150">1인당 평균</th>
        </tr>
    </thead>

    <tbody>
<%
totalPayCount = 0
totalBuyerCount = 0
totalActual = 0

If resultMap.Count > 0 Then

    resultKeys = resultMap.Keys

    '======================================================
    ' 정렬
    ' 1. 연월 오름차순
    ' 2. 담당과목 가나다순
    ' 3. 학교급 오름차순
    '======================================================
    If UBound(resultKeys) > 0 Then

        For i = 0 To UBound(resultKeys) - 1

            For j = i + 1 To UBound(resultKeys)

                keyPartsA = Split(resultKeys(i), "|")
                keyPartsB = Split(resultKeys(j), "|")

                sortA = GetSubjectSortValue(keyPartsA)
                sortB = GetSubjectSortValue(keyPartsB)

                If sortA > sortB Then

                    keyTemp = resultKeys(i)
                    resultKeys(i) = resultKeys(j)
                    resultKeys(j) = keyTemp

                End If

            Next

        Next

    End If

    For i = 0 To UBound(resultKeys)

        rowData = resultMap(resultKeys(i))

        payMonth = rowData(0)
        subjectName = rowData(1)
        schoolDivision = rowData(2)

        payCount = CLng(Nz(rowData(3), 0))
        actualAmount = CLng(Nz(rowData(4), 0))
        buyerCount = CLng(Nz(rowData(5), 0))

        If buyerCount > 0 Then
            averageAmount = actualAmount / buyerCount
        Else
            averageAmount = 0
        End If

        totalPayCount = totalPayCount + payCount
        totalBuyerCount = totalBuyerCount + buyerCount
        totalActual = totalActual + actualAmount
%>
        <tr>
            <td align="center"
                style="mso-number-format:'\@';">
                <%=Server.HTMLEncode(payMonth & "")%>
            </td>

            <td align="center">
                <%=Server.HTMLEncode(subjectName & "")%>
            </td>

            <td align="center">
                <%=Server.HTMLEncode(schoolDivision & "")%>
            </td>

            <td align="right"
                style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(payCount, 0)%>
            </td>

            <td align="right"
                style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(buyerCount, 0)%>
            </td>

            <td align="right"
                style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(actualAmount, 0)%>
            </td>

            <td align="right"
                style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(averageAmount, 0)%>
            </td>
        </tr>
<%
    Next

Else
%>
        <tr>
            <td colspan="7" align="center">
                조회 결과가 없습니다.
            </td>
        </tr>
<%
End If
%>
    </tbody>

    <tfoot>
        <tr style="background-color: #E7E6E6; font-weight: bold;">
            <td colspan="3" align="center">
                전체 합계
            </td>

            <td align="right"
                style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(totalPayCount, 0)%>
            </td>

            <td align="right"
                style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(totalBuyerCount, 0)%>
            </td>

            <td align="right"
                style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(totalActual, 0)%>
            </td>

            <td align="right"
                style="mso-number-format:'\#\,\#\#0';">
                <%
                If totalBuyerCount > 0 Then
                    Response.Write FormatNumber(totalActual / totalBuyerCount, 0)
                Else
                    Response.Write "0"
                End If
                %>
            </td>
        </tr>
    </tfoot>
</table>

</body>
</html>

<%
'==========================================================
' 정렬용 숫자 반환
' 연월 → 담당과목 → 학교급
'==========================================================
Function GetSubjectSortValue(ByVal keyParts)

    Dim monthValue
    Dim subjectValue
    Dim schoolValue

    monthValue = ""
    subjectValue = ""
    schoolValue = ""

    If IsArray(keyParts) Then

        If UBound(keyParts) >= 0 Then
            monthValue = Trim(keyParts(0))
        End If

        If UBound(keyParts) >= 1 Then
            subjectValue = Trim(keyParts(1))
        End If

        If UBound(keyParts) >= 2 Then
            schoolValue = Trim(keyParts(2))
        End If

    End If

    ' 미분류는 각 정렬 그룹의 아래쪽으로 이동
    If subjectValue = "미분류" Then
        subjectValue = "ZZZZ_미분류"
    End If

    If schoolValue = "미분류" Then
        schoolValue = "ZZZZ_미분류"
    End If

    GetSubjectSortValue = monthValue & "|" & subjectValue & "|" & schoolValue

End Function


'==========================================================
' NULL 및 빈 문자열 처리
'==========================================================
Function Nz(ByVal value, ByVal defaultValue)

    If IsNull(value) Or Trim(value & "") = "" Then
        Nz = defaultValue
    Else
        Nz = value
    End If

End Function
%>