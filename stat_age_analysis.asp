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
Dim arrResult
Dim arrAgeData
Dim arrTopProducts
Dim i
Dim j
Dim k

Dim sYear
Dim sMonth
Dim eYear
Dim eMonth
Dim sDay
Dim eDay

Dim startDate
Dim endDate
Dim downloadDate

Dim totalBuyerCount
Dim totalActual
Dim totalSmsCount

Dim payYear
Dim userid
Dim genderName
Dim ageRange
Dim birthYear
Dim calcAge

Dim buyerCount
Dim actualAmount
Dim smsCount
Dim averageAmount
Dim smsRate

Dim dictAge
Dim dictResult
Dim dictTopProducts

Dim userIdList
Dim resultKeys
Dim resultKey
Dim keyTemp

Dim rowData
Dim existData
Dim keyPartsA
Dim keyPartsB
Dim sortA
Dim sortB

Dim productName
Dim saleCount
Dim productAmount

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
' 1. 결제 기본 데이터 조회
'
' 기준: pay_info 기본 집계
' 특징: FamilyM 연령 데이터는 별도 조회
'      연도·회원·성별·결제금액·SMS만 조회
' 구성:
'   - 결제연도: LEFT(pay_date, 4)
'   - 구매자: 고유 userid
'   - 성별: vUser_info.gender 매핑 (1=남성, 0=여성, 그외=미분류)
'   - 결제액: pay_info.pay_price 합계
'   - SMS: info_receive_agree에서 agree_type = 'SMS' 확인
'==========================================================
strSQL = ""

strSQL = strSQL & "SELECT "
strSQL = strSQL & "    LEFT(p.pay_date, 4) AS pay_year, "
strSQL = strSQL & "    p.userid, "
strSQL = strSQL & "    1 AS buyer_count, "
strSQL = strSQL & "    SUM(ISNULL(p.pay_price, 0)) AS actual_amount, "
strSQL = strSQL & "    CASE "
strSQL = strSQL & "        WHEN sms.userid IS NOT NULL THEN 1 "
strSQL = strSQL & "        ELSE 0 "
strSQL = strSQL & "    END AS sms_count, "
strSQL = strSQL & "    CASE "
strSQL = strSQL & "        WHEN u.gender = 1 THEN '남성' "
strSQL = strSQL & "        WHEN u.gender = 0 THEN '여성' "
strSQL = strSQL & "        ELSE '미분류' "
strSQL = strSQL & "    END AS gender_name "

strSQL = strSQL & "FROM pay_info AS p WITH (READUNCOMMITTED) "

strSQL = strSQL & "LEFT JOIN vUser_info AS u WITH (READUNCOMMITTED) "
strSQL = strSQL & "    ON p.userid = u.userid "

strSQL = strSQL & "LEFT JOIN ( "
strSQL = strSQL & "    SELECT userid "
strSQL = strSQL & "    FROM info_receive_agree WITH (READUNCOMMITTED) "
strSQL = strSQL & "    WHERE agree_type = 'SMS' "
strSQL = strSQL & "    GROUP BY userid "
strSQL = strSQL & ") AS sms "
strSQL = strSQL & "    ON p.userid = sms.userid "

strSQL = strSQL & "WHERE p.sell_info = 'o' "
strSQL = strSQL & "  AND p.pay_date >= '" & startDate & "' "
strSQL = strSQL & "  AND p.pay_date < DATEADD(DAY, 1, '" & endDate & "') "

strSQL = strSQL & "GROUP BY "
strSQL = strSQL & "    LEFT(p.pay_date, 4), "
strSQL = strSQL & "    p.userid, "
strSQL = strSQL & "    CASE "
strSQL = strSQL & "        WHEN sms.userid IS NOT NULL THEN 1 "
strSQL = strSQL & "        ELSE 0 "
strSQL = strSQL & "    END, "
strSQL = strSQL & "    CASE "
strSQL = strSQL & "        WHEN u.gender = 1 THEN '남성' "
strSQL = strSQL & "        WHEN u.gender = 0 THEN '여성' "
strSQL = strSQL & "        ELSE '미분류' "
strSQL = strSQL & "    END "

strSQL = strSQL & "ORDER BY "
strSQL = strSQL & "    LEFT(p.pay_date, 4), "
strSQL = strSQL & "    p.userid "

arrResult = ExecSP(strSQL, strCyberTeacher)

'==========================================================
' 2. FamilyM 회원 생년 조회
'
' 기준: FamilyM.aes.f_user_add_site 함수
' 특징: usp_openKey 실행하여 암호화된 데이터 접근 활성화
'      기간 내 결제한 모든 회원의 출생연도 조회
' 목적: 연령대 계산
'   - 20세미만, 20대, 30대, 40대, 50대, 60대, 70대이상, 미분류
'==========================================================
Set dictAge = CreateObject("Scripting.Dictionary")

If IsArray(arrResult) Then

    For i = 0 To UBound(arrResult, 2)

        userid = Nz(arrResult(1, i), "")

        If userid <> "" Then
            If Not dictAge.Exists(userid) Then
                dictAge.Add userid, 0
            End If
        End If

    Next

    If dictAge.Count > 0 Then

        strSQL = ""
        strSQL = strSQL & "EXEC dbo.usp_openKey; "
        strSQL = strSQL & "SELECT userid, b_year "
        strSQL = strSQL & "FROM FamilyM.aes.f_user_add_site('www.ybmteachers.com') "
        strSQL = strSQL & "WHERE userid IN ("

        userIdList = dictAge.Keys

        For i = 0 To UBound(userIdList)

            If i > 0 Then
                strSQL = strSQL & ","
            End If

            strSQL = strSQL & fnCdbValue(userIdList(i))

        Next

        strSQL = strSQL & ")"

        arrAgeData = ExecSP(strSQL, strCyberTeacher)

        If IsArray(arrAgeData) Then

            For i = 0 To UBound(arrAgeData, 2)

                userid = Nz(arrAgeData(0, i), "")
                birthYear = Nz(arrAgeData(1, i), 0)

                If userid <> "" And dictAge.Exists(userid) Then

                    If IsNumeric(birthYear) Then
                        dictAge(userid) = CLng(birthYear)
                    Else
                        dictAge(userid) = 0
                    End If

                End If

            Next

        End If

    End If

End If

'==========================================================
' 3. 연령대·성별별 결과 누적
'
' 기준: 연도·연령대·성별 조합
' 특징:
'   - 같은 조합이 여러 회원에서 나타나면 누적
'   - 연령대: 생년에서 계산 (결제연도 - 생년 + 1)
'   - 성별: vUser_info.gender 기준
'   - 결제액: 해당 기간 회원의 모든 결제액
'==========================================================
Set dictResult = CreateObject("Scripting.Dictionary")

If IsArray(arrResult) Then

    For i = 0 To UBound(arrResult, 2)

        payYear = Nz(arrResult(0, i), "")
        userid = Nz(arrResult(1, i), "")

        buyerCount = CLng(Nz(arrResult(2, i), 0))
        actualAmount = CLng(Nz(arrResult(3, i), 0))
        smsCount = CLng(Nz(arrResult(4, i), 0))
        genderName = Nz(arrResult(5, i), "미분류")

        birthYear = 0

        If dictAge.Exists(userid) Then
            birthYear = CLng(Nz(dictAge(userid), 0))
        End If

        '==================================================
        ' 연령대 계산
        ' 결제연도 - 생년 + 1
        '==================================================
        If birthYear > 0 And IsNumeric(payYear) Then

            calcAge = CLng(payYear) - birthYear + 1

            If calcAge < 20 Then
                ageRange = "20세미만"
            ElseIf calcAge <= 29 Then
                ageRange = "20대"
            ElseIf calcAge <= 39 Then
                ageRange = "30대"
            ElseIf calcAge <= 49 Then
                ageRange = "40대"
            ElseIf calcAge <= 59 Then
                ageRange = "50대"
            ElseIf calcAge <= 69 Then
                ageRange = "60대"
            Else
                ageRange = "70대이상"
            End If

        Else

            ageRange = "미분류"

        End If

        resultKey = payYear & "|" & ageRange & "|" & genderName

        If Not dictResult.Exists(resultKey) Then

            dictResult.Add resultKey, Array( _
                payYear, _
                ageRange, _
                genderName, _
                buyerCount, _
                actualAmount, _
                smsCount _
            )

        Else

            existData = dictResult(resultKey)

            existData(3) = existData(3) + buyerCount
            existData(4) = existData(4) + actualAmount
            existData(5) = existData(5) + smsCount

            dictResult(resultKey) = existData

        End If

    Next

End If

'==========================================================
' 4. 연령대·성별별 인기 과정 TOP 20 조회
'
' 기준: 판매수 기준으로 정렬
' 특징:
'   - 각 연령대·성별 조합별로 인기 과정 TOP 20 추출
'   - 고유 pay_product_num 기준 (중복 제거)
'   - product_type = 'S' (과정 상품만 포함)
' 구성:
'   - 과정명: product_info.product_name
'   - 판매수: COUNT(DISTINCT pay_product_num)
'   - 결제액: SUM(product_price)
'==========================================================
strSQL = ""

strSQL = strSQL & "EXEC dbo.usp_openKey; "

strSQL = strSQL & "WITH AgeGenderPayment AS ( "
strSQL = strSQL & "    SELECT "
strSQL = strSQL & "        CASE "
strSQL = strSQL & "            WHEN fm.b_year IS NULL OR ISNUMERIC(fm.b_year) = 0 "
strSQL = strSQL & "                THEN '미분류' "
strSQL = strSQL & "            WHEN YEAR(p.pay_date) - CONVERT(INT, fm.b_year) + 1 < 20 "
strSQL = strSQL & "                THEN '20세미만' "
strSQL = strSQL & "            WHEN YEAR(p.pay_date) - CONVERT(INT, fm.b_year) + 1 <= 29 "
strSQL = strSQL & "                THEN '20대' "
strSQL = strSQL & "            WHEN YEAR(p.pay_date) - CONVERT(INT, fm.b_year) + 1 <= 39 "
strSQL = strSQL & "                THEN '30대' "
strSQL = strSQL & "            WHEN YEAR(p.pay_date) - CONVERT(INT, fm.b_year) + 1 <= 49 "
strSQL = strSQL & "                THEN '40대' "
strSQL = strSQL & "            WHEN YEAR(p.pay_date) - CONVERT(INT, fm.b_year) + 1 <= 59 "
strSQL = strSQL & "                THEN '50대' "
strSQL = strSQL & "            WHEN YEAR(p.pay_date) - CONVERT(INT, fm.b_year) + 1 <= 69 "
strSQL = strSQL & "                THEN '60대' "
strSQL = strSQL & "            ELSE '70대이상' "
strSQL = strSQL & "        END AS age_range, "
strSQL = strSQL & "        CASE "
strSQL = strSQL & "            WHEN u.gender = 1 THEN '남성' "
strSQL = strSQL & "            WHEN u.gender = 0 THEN '여성' "
strSQL = strSQL & "            ELSE '미분류' "
strSQL = strSQL & "        END AS gender_name, "
strSQL = strSQL & "        pi.product_num, "
strSQL = strSQL & "        pi.product_name, "
strSQL = strSQL & "        pip.pay_product_num, "
strSQL = strSQL & "        pip.product_price "
strSQL = strSQL & "    FROM pay_info AS p WITH (READUNCOMMITTED) "
strSQL = strSQL & "    LEFT JOIN vUser_info AS u WITH (READUNCOMMITTED) "
strSQL = strSQL & "        ON p.userid = u.userid "
strSQL = strSQL & "    LEFT JOIN FamilyM.aes.f_user_add_site('www.ybmteachers.com') AS fm "
strSQL = strSQL & "        ON p.userid = fm.userid "
strSQL = strSQL & "    INNER JOIN pay_info_product AS pip WITH (READUNCOMMITTED) "
strSQL = strSQL & "        ON p.pay_num = pip.pay_num "
strSQL = strSQL & "    INNER JOIN product_info AS pi WITH (READUNCOMMITTED) "
strSQL = strSQL & "        ON pip.product_num = pi.product_num "
strSQL = strSQL & "    WHERE p.sell_info = 'o' "
strSQL = strSQL & "        AND pi.product_type = 'S' "
strSQL = strSQL & "        AND p.pay_date >= '" & startDate & "' "
strSQL = strSQL & "        AND p.pay_date < DATEADD(DAY, 1, '" & endDate & "') "
strSQL = strSQL & ") "
strSQL = strSQL & "SELECT "
strSQL = strSQL & "    age_range, "
strSQL = strSQL & "    gender_name, "
strSQL = strSQL & "    ISNULL(product_name, '과정명 없음') AS product_name, "
strSQL = strSQL & "    COUNT(DISTINCT pay_product_num) AS sale_count, "
strSQL = strSQL & "    SUM(ISNULL(product_price, 0)) AS actual_amount "
strSQL = strSQL & "FROM AgeGenderPayment "
strSQL = strSQL & "GROUP BY age_range, gender_name, product_name "
strSQL = strSQL & "ORDER BY "
strSQL = strSQL & "    age_range, "
strSQL = strSQL & "    gender_name, "
strSQL = strSQL & "    COUNT(DISTINCT pay_product_num) DESC, "
strSQL = strSQL & "    SUM(ISNULL(product_price, 0)) DESC "

arrTopProducts = ExecSP(strSQL, strCyberTeacher)

'==========================================================
' Top Products를 Dictionary로 구성
' 키: age_range|gender_name
' 값: Dictionary of product items
'==========================================================
Set dictTopProducts = CreateObject("Scripting.Dictionary")

If IsArray(arrTopProducts) Then

    Dim currentKey
    Dim productItem
    Dim productIdx

    currentKey = ""
    productIdx = 0

    For i = 0 To UBound(arrTopProducts, 2)

        ageRange = Nz(arrTopProducts(0, i), "미분류")
        genderName = Nz(arrTopProducts(1, i), "미분류")
        productName = Nz(arrTopProducts(2, i), "과정명 없음")
        saleCount = CLng(Nz(arrTopProducts(3, i), 0))
        productAmount = CDbl(Nz(arrTopProducts(4, i), 0))

        resultKey = ageRange & "|" & genderName

        If resultKey <> currentKey Then

            currentKey = resultKey
            productIdx = 0

            If Not dictTopProducts.Exists(resultKey) Then
                Set dictTopProducts(resultKey) = CreateObject("Scripting.Dictionary")
            End If

        End If

        If productIdx < 20 Then

            productItem = Array(productName, saleCount, productAmount)
            dictTopProducts(resultKey).Add CStr(productIdx), productItem
            productIdx = productIdx + 1

        End If

    Next

End If

'==========================================================
' 개인정보 파일 다운로드 로그
'==========================================================
If IsArray(arrResult) Then

    strSQL = ""
    strSQL = strSQL & "EXEC site_log.dbo.p_File_Log "
    strSQL = strSQL & fnCdbValue(Session("admin_id")) & ","
    strSQL = strSQL & fnCdbValue(Request.ServerVariables("REMOTE_ADDR")) & ","
    strSQL = strSQL & fnCdbValue(Request.ServerVariables("HTTP_HOST")) & ","
    strSQL = strSQL & fnCdbValue("stat_age_analysis_" & downloadDate) & ","
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
    "attachment; filename=stat_age_analysis_" & _
    downloadDate & ".xls"
%>

<html>
<head>
<meta http-equiv="Content-Type"
      content="text/html; charset=euc-kr">
</head>

<body>

<!--======================================================
    Excel 상단 설명
=======================================================-->
<p style="color: #666; font-size: 12px;">
    조회 기간 동안 결제한 회원을 연령대와 성별로 나누어 구매 현황을 비교할 수 있는 자료입니다.<br>
    20세미만, 20대, 30대, 40대, 50대, 60대, 70대이상 회원의 연도별·성별 구매자 수, 총 결제액,<br>
    1인당 평균 결제액과 문자 수신동의 현황을 확인할 수 있습니다.<br>
    각 연령대·성별 그룹의 인기 과정 TOP 20도 함께 확인할 수 있습니다.<br>
    연령 정보가 없는 회원은 '미분류'로 표시됩니다.
</p>

<hr style="margin: 20px 0; border: 1px solid #999;">

<!--======================================================
    연령대·성별별 구매 현황 표
=======================================================-->
<table border="1">
    <thead>
        <tr style="background-color: #4472C4; color: white; font-weight: bold;">
            <th width="100">연도</th>
            <th width="150">연령대</th>
            <th width="100">성별</th>
            <th width="120">구매자수</th>
            <th width="150">실제결제액</th>
            <th width="150">1인당 평균</th>
            <th width="150">SMS 수신동의</th>
            <th width="150">SMS 동의율</th>
        </tr>
    </thead>

    <tbody>
<%
totalBuyerCount = 0
totalActual = 0
totalSmsCount = 0

'==========================================================
' 결과 출력
' 정렬 순서:
' 1. 연도 오름차순
' 2. 연령대 오름차순
' 3. 성별 순서 (남성 → 여성 → 미분류)
'
' 연령대:
' 20세미만 → 20대 → 30대 → 40대
' → 50대 → 60대 → 70대이상 → 미분류
'==========================================================
If dictResult.Count > 0 Then

    resultKeys = dictResult.Keys

    '------------------------------------------------------
    ' Dictionary 키 정렬
    '------------------------------------------------------
    If UBound(resultKeys) > 0 Then

        For i = 0 To UBound(resultKeys) - 1

            For j = i + 1 To UBound(resultKeys)

                keyPartsA = Split(resultKeys(i), "|")
                keyPartsB = Split(resultKeys(j), "|")

                sortA = GetAgeResultSortValue(keyPartsA)
                sortB = GetAgeResultSortValue(keyPartsB)

                If sortA > sortB Then

                    keyTemp = resultKeys(i)
                    resultKeys(i) = resultKeys(j)
                    resultKeys(j) = keyTemp

                End If

            Next

        Next

    End If

    '------------------------------------------------------
    ' 정렬된 결과 출력
    '------------------------------------------------------
    For i = 0 To UBound(resultKeys)

        rowData = dictResult(resultKeys(i))

        payYear = rowData(0)
        ageRange = rowData(1)
        genderName = rowData(2)

        buyerCount = CLng(Nz(rowData(3), 0))
        actualAmount = CLng(Nz(rowData(4), 0))
        smsCount = CLng(Nz(rowData(5), 0))

        If buyerCount > 0 Then
            averageAmount = actualAmount / buyerCount
            smsRate = (smsCount / buyerCount) * 100
        Else
            averageAmount = 0
            smsRate = 0
        End If

        totalBuyerCount = totalBuyerCount + buyerCount
        totalActual = totalActual + actualAmount
        totalSmsCount = totalSmsCount + smsCount
%>
        <tr>
            <td align="center"
                style="mso-number-format:'\@';">
                <%=Server.HTMLEncode(payYear & "")%>
            </td>

            <td align="center">
                <%=Server.HTMLEncode(ageRange & "")%>
            </td>

            <td align="center">
                <%=Server.HTMLEncode(genderName & "")%>
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

            <td align="right"
                style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(smsCount, 0)%>
            </td>

            <td align="right"
                style="mso-number-format:'0.00%';">
                <%=FormatNumber(smsRate / 100, 4)%>
            </td>
        </tr>
<%
    Next

Else
%>
        <tr>
            <td colspan="8" align="center">
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

            <td align="right"
                style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(totalSmsCount, 0)%>
            </td>

            <td align="right"
                style="mso-number-format:'0.00%';">
                <%
                If totalBuyerCount > 0 Then
                    Response.Write FormatNumber(totalSmsCount / totalBuyerCount, 4)
                Else
                    Response.Write "0"
                End If
                %>
            </td>
        </tr>
    </tfoot>
</table>

<hr style="margin: 30px 0; border: 1px solid #999;">

<!--======================================================
    연령대·성별별 인기 과정 TOP 20
=======================================================-->
<%
If dictTopProducts.Count > 0 Then

    Dim allTopKeys
    Dim topKey
    Dim topKeyParts
    Dim topAgeRange
    Dim topGenderName
    Dim productDictionary
    Dim productKeys
    Dim productRank

    allTopKeys = dictTopProducts.Keys

    '------------------------------------------------------
    ' Top Products 키 정렬
    '------------------------------------------------------
    If UBound(allTopKeys) > 0 Then

        For i = 0 To UBound(allTopKeys) - 1

            For j = i + 1 To UBound(allTopKeys)

                keyPartsA = Split("0|" & allTopKeys(i), "|")
                keyPartsB = Split("0|" & allTopKeys(j), "|")

                sortA = GetAgeResultSortValue(keyPartsA)
                sortB = GetAgeResultSortValue(keyPartsB)

                If sortA > sortB Then

                    keyTemp = allTopKeys(i)
                    allTopKeys(i) = allTopKeys(j)
                    allTopKeys(j) = keyTemp

                End If

            Next

        Next

    End If

    '------------------------------------------------------
    ' 각 연령대·성별 그룹의 인기 과정 출력
    '------------------------------------------------------
    For i = 0 To UBound(allTopKeys)

        topKey = allTopKeys(i)
        keyPartsA = Split(topKey, "|")

        topAgeRange = Nz(keyPartsA(0), "미분류")
        topGenderName = Nz(keyPartsA(1), "미분류")

        Response.Write "<table border=""1"" style=""margin-bottom:20px;"">" & vbCrLf

        ' 제목
        Response.Write "<tr><td colspan=""5"" " & _
                       "style=""font-weight:bold; font-size:14px; background-color:#D3D3D3; padding:8px;"">" & _
                       Server.HTMLEncode(topAgeRange & " · " & topGenderName & " - 선호 과정 TOP 20") & _
                       "</td></tr>" & vbCrLf

        ' 헤더
        Response.Write "<tr style=""background-color:#4472C4; color:white; font-weight:bold;"">" & vbCrLf
        Response.Write "<th width=""50"">순위</th>" & vbCrLf
        Response.Write "<th width=""280"">과정명</th>" & vbCrLf
        Response.Write "<th width=""100"">판매수</th>" & vbCrLf
        Response.Write "<th width=""150"">결제액</th>" & vbCrLf
        Response.Write "<th width=""150"">판매액/판매수</th>" & vbCrLf
        Response.Write "</tr>" & vbCrLf

        ' 본문
        Set productDictionary = dictTopProducts(topKey)
        productKeys = productDictionary.Keys

        Dim subTotalSaleCount
        Dim subTotalAmount

        subTotalSaleCount = 0
        subTotalAmount = 0

        For j = 0 To UBound(productKeys)

            productItem = productDictionary(productKeys(j))

            productName = Nz(productItem(0), "과정명 없음")
            saleCount = CLng(Nz(productItem(1), 0))
            productAmount = CDbl(Nz(productItem(2), 0))

            subTotalSaleCount = subTotalSaleCount + saleCount
            subTotalAmount = subTotalAmount + productAmount

            Response.Write "<tr>" & vbCrLf
            Response.Write "<td align=""center"">" & (j + 1) & "</td>" & vbCrLf
            Response.Write "<td align=""left"">" & Server.HTMLEncode(productName & "") & "</td>" & vbCrLf
            Response.Write "<td align=""right"" style=""mso-number-format:'\#\,\#\#0';"">" & _
                           FormatNumber(saleCount, 0) & "</td>" & vbCrLf
            Response.Write "<td align=""right"" style=""mso-number-format:'\#\,\#\#0';"">" & _
                           FormatNumber(productAmount, 0) & "</td>" & vbCrLf

            If saleCount > 0 Then
                Response.Write "<td align=""right"" style=""mso-number-format:'\#\,\#\#0';"">" & _
                               FormatNumber(productAmount / saleCount, 0) & "</td>" & vbCrLf
            Else
                Response.Write "<td align=""right"">0</td>" & vbCrLf
            End If

            Response.Write "</tr>" & vbCrLf

        Next

        ' 행 합계
        Response.Write "<tr style=""background-color:#E7E6E6; font-weight:bold;"">" & vbCrLf
        Response.Write "<td colspan=""2"" align=""center"">상위 " & (j) & " 합계</td>" & vbCrLf
        Response.Write "<td align=""right"" style=""mso-number-format:'\#\,\#\#0';"">" & _
                       FormatNumber(subTotalSaleCount, 0) & "</td>" & vbCrLf
        Response.Write "<td align=""right"" style=""mso-number-format:'\#\,\#\#0';"">" & _
                       FormatNumber(subTotalAmount, 0) & "</td>" & vbCrLf

        If subTotalSaleCount > 0 Then
            Response.Write "<td align=""right"" style=""mso-number-format:'\#\,\#\#0';"">" & _
                           FormatNumber(subTotalAmount / subTotalSaleCount, 0) & "</td>" & vbCrLf
        Else
            Response.Write "<td align=""right"">0</td>" & vbCrLf
        End If

        Response.Write "</tr>" & vbCrLf

        Response.Write "</table>" & vbCrLf

    Next

End If
%>

</body>
</html>

<%
'==========================================================
' 정렬용 숫자 반환
' 연도 → 연령대 → 성별 순서
'==========================================================
Function GetAgeResultSortValue(ByVal keyParts)

    Dim sortYear
    Dim sortAge
    Dim sortGender

    sortYear = 0
    sortAge = 99
    sortGender = 99

    If IsArray(keyParts) Then

        If UBound(keyParts) >= 0 Then
            If IsNumeric(keyParts(0)) Then
                sortYear = CLng(keyParts(0))
            End If
        End If

        If UBound(keyParts) >= 1 Then
            sortAge = GetAgeSortOrder(keyParts(1))
        End If

        If UBound(keyParts) >= 2 Then
            sortGender = GetGenderSortOrder(keyParts(2))
        End If

    End If

    GetAgeResultSortValue = _
        (sortYear * 100000) + _
        (sortAge * 100) + _
        sortGender

End Function

'==========================================================
' 연령대 정렬 순서
'==========================================================
Function GetAgeSortOrder(ByVal ageName)

    Select Case Trim(ageName)

        Case "20세미만"
            GetAgeSortOrder = 1

        Case "20대"
            GetAgeSortOrder = 2

        Case "30대"
            GetAgeSortOrder = 3

        Case "40대"
            GetAgeSortOrder = 4

        Case "50대"
            GetAgeSortOrder = 5

        Case "60대"
            GetAgeSortOrder = 6

        Case "70대이상"
            GetAgeSortOrder = 7

        Case "미분류"
            GetAgeSortOrder = 99

        Case Else
            GetAgeSortOrder = 98

    End Select

End Function

'==========================================================
' 성별 정렬 순서
'==========================================================
Function GetGenderSortOrder(ByVal genderName)

    Select Case Trim(genderName)

        Case "남성"
            GetGenderSortOrder = 1

        Case "여성"
            GetGenderSortOrder = 2

        Case "미분류"
            GetGenderSortOrder = 99

        Case Else
            GetGenderSortOrder = 98

    End Select

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
