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
Dim arrUserPay
Dim arrTopProduct

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
Dim payCount
Dim actualAmount
Dim couponUsed

Dim groupName
Dim groupMap
Dim groupKeys
Dim groupData

Dim productName
Dim buyerCount
Dim saleCount

Dim totalUserCount
Dim totalPayCount
Dim totalActual
Dim totalCouponUser

Dim repurchaseUserCount
Dim repurchasePayCount
Dim repurchaseActual

Dim avgPayCount
Dim avgAmount
Dim couponRate
Dim repurchaseRate
%>
<%
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
' 1. 회원별 구매횟수 / 결제액 / 쿠폰사용 여부
'
' 쿠폰 사용은 pay_coupon_use에 연결된 결제가
' 한 건이라도 있으면 사용자로 본다.
'==========================================================
strSQL = ""
strSQL = strSQL & "SELECT "
strSQL = strSQL & "    p.userid, "
strSQL = strSQL & "    COUNT(DISTINCT p.pay_num) AS pay_count, "
strSQL = strSQL & "    SUM(ISNULL(p.pay_price, 0)) AS actual_amount, "
strSQL = strSQL & "    MAX(CASE WHEN cu.pay_product_num IS NOT NULL THEN 1 ELSE 0 END) AS coupon_used "
strSQL = strSQL & "FROM pay_info AS p WITH (READUNCOMMITTED) "
strSQL = strSQL & "LEFT JOIN pay_info_product AS pip WITH (READUNCOMMITTED) "
strSQL = strSQL & "    ON p.pay_num = pip.pay_num "
strSQL = strSQL & "LEFT JOIN pay_coupon_use AS cu WITH (READUNCOMMITTED) "
strSQL = strSQL & "    ON pip.pay_product_num = cu.pay_product_num "
strSQL = strSQL & "WHERE p.sell_info = 'o' "
strSQL = strSQL & "  AND p.pay_date >= '" & startDate & "' "
strSQL = strSQL & "  AND p.pay_date <= '" & endDate & "' "
strSQL = strSQL & "  AND p.userid IS NOT NULL "
strSQL = strSQL & "  AND LTRIM(RTRIM(p.userid)) <> '' "
strSQL = strSQL & "GROUP BY p.userid "

arrUserPay = ExecSP(strSQL, strCyberTeacher)

'==========================================================
' 2. 구매횟수 구간별 집계
'
' groupMap(구간) = Array(회원수, 결제건수, 결제액, 쿠폰사용자수)
'==========================================================
Set groupMap = CreateObject("Scripting.Dictionary")

totalUserCount = 0
totalPayCount = 0
totalActual = 0
totalCouponUser = 0

repurchaseUserCount = 0
repurchasePayCount = 0
repurchaseActual = 0

If IsArray(arrUserPay) Then

    For i = 0 To UBound(arrUserPay, 2)

        userid = Trim(Nz(arrUserPay(0, i), "") & "")
        payCount = CDbl(Nz(arrUserPay(1, i), 0))
        actualAmount = CDbl(Nz(arrUserPay(2, i), 0))
        couponUsed = CDbl(Nz(arrUserPay(3, i), 0))

        If userid <> "" Then

            groupName = GetPayCountGroup(payCount)

            If Not groupMap.Exists(groupName) Then
                groupMap.Add groupName, Array(0, 0, 0, 0)
            End If

            groupData = groupMap(groupName)
            groupData(0) = CDbl(groupData(0)) + 1
            groupData(1) = CDbl(groupData(1)) + payCount
            groupData(2) = CDbl(groupData(2)) + actualAmount
            groupData(3) = CDbl(groupData(3)) + couponUsed
            groupMap(groupName) = groupData

            totalUserCount = totalUserCount + 1
            totalPayCount = totalPayCount + payCount
            totalActual = totalActual + actualAmount
            totalCouponUser = totalCouponUser + couponUsed

            If payCount >= 2 Then
                repurchaseUserCount = repurchaseUserCount + 1
                repurchasePayCount = repurchasePayCount + payCount
                repurchaseActual = repurchaseActual + actualAmount
            End If

        End If

    Next

End If

'==========================================================
' 3. 재구매자(2회 이상) 선호 과정 TOP 20
'
' 기간 내 결제건수가 2건 이상인 회원만 대상으로 한다.
' 교재는 제외하고 과정(product_type = 'S')만 집계한다.
'==========================================================
strSQL = ""
strSQL = strSQL & ";WITH RepurchaseUser AS ( "
strSQL = strSQL & "    SELECT p.userid "
strSQL = strSQL & "    FROM pay_info AS p WITH (READUNCOMMITTED) "
strSQL = strSQL & "    WHERE p.sell_info = 'o' "
strSQL = strSQL & "      AND p.pay_date >= '" & startDate & "' "
strSQL = strSQL & "      AND p.pay_date <= '" & endDate & "' "
strSQL = strSQL & "      AND p.userid IS NOT NULL "
strSQL = strSQL & "      AND LTRIM(RTRIM(p.userid)) <> '' "
strSQL = strSQL & "    GROUP BY p.userid "
strSQL = strSQL & "    HAVING COUNT(DISTINCT p.pay_num) >= 2 "
strSQL = strSQL & ") "
strSQL = strSQL & "SELECT TOP 20 "
strSQL = strSQL & "    ISNULL(pi.product_name, '과정명 없음') AS product_name, "
strSQL = strSQL & "    COUNT(DISTINCT p.userid) AS buyer_count, "
strSQL = strSQL & "    COUNT(DISTINCT pip.pay_product_num) AS sale_count, "
strSQL = strSQL & "    SUM(ISNULL(pip.product_price, 0)) AS actual_amount "
strSQL = strSQL & "FROM pay_info AS p WITH (READUNCOMMITTED) "
strSQL = strSQL & "INNER JOIN RepurchaseUser AS ru "
strSQL = strSQL & "    ON p.userid = ru.userid "
strSQL = strSQL & "INNER JOIN pay_info_product AS pip WITH (READUNCOMMITTED) "
strSQL = strSQL & "    ON p.pay_num = pip.pay_num "
strSQL = strSQL & "INNER JOIN product_info AS pi WITH (READUNCOMMITTED) "
strSQL = strSQL & "    ON pip.product_num = pi.product_num "
strSQL = strSQL & "WHERE p.sell_info = 'o' "
strSQL = strSQL & "  AND pi.product_type = 'S' "
strSQL = strSQL & "  AND p.pay_date >= '" & startDate & "' "
strSQL = strSQL & "  AND p.pay_date <= '" & endDate & "' "
strSQL = strSQL & "GROUP BY pi.product_name "
strSQL = strSQL & "ORDER BY "
strSQL = strSQL & "    COUNT(DISTINCT pip.pay_product_num) DESC, "
strSQL = strSQL & "    COUNT(DISTINCT p.userid) DESC "

arrTopProduct = ExecSP(strSQL, strCyberTeacher)

'==========================================================
' 개인정보 파일 다운로드 로그
'==========================================================
If IsArray(arrUserPay) Then

    strSQL = ""
    strSQL = strSQL & "EXEC site_log.dbo.p_File_Log "
    strSQL = strSQL & fnCdbValue(Session("admin_id")) & ","
    strSQL = strSQL & fnCdbValue(Request.ServerVariables("REMOTE_ADDR")) & ","
    strSQL = strSQL & fnCdbValue(Request.ServerVariables("HTTP_HOST")) & ","
    strSQL = strSQL & fnCdbValue("stat_repurchase_analysis_" & downloadDate) & ","
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
    "attachment; filename=stat_repurchase_analysis_" & _
    downloadDate & ".xls"
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
</head>

<body>

<!-- ================= 요약 ================= -->
<table border="1">
    <tr>
        <td colspan="2"
            style="font-weight:bold; font-size:14px; background-color:#D3D3D3;">
            재구매 요약 (<%=Server.HTMLEncode(startDate)%> ~ <%=Server.HTMLEncode(endDate)%>)
        </td>
    </tr>

    <tr>
        <td width="200">전체 구매자수</td>
        <td width="200" align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(totalUserCount, 0)%>
        </td>
    </tr>

    <tr>
        <td>재구매자수 (2회 이상)</td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(repurchaseUserCount, 0)%>
        </td>
    </tr>

    <tr>
        <td>재구매율</td>
        <td align="right" style="mso-number-format:'0.00%';">
            <%
            If totalUserCount > 0 Then
                repurchaseRate = repurchaseUserCount / totalUserCount
            Else
                repurchaseRate = 0
            End If
            Response.Write FormatNumber(repurchaseRate, 4)
            %>
        </td>
    </tr>

    <tr>
        <td>전체 평균 구매횟수</td>
        <td align="right" style="mso-number-format:'\#\,\#\#0\.00';">
            <%
            If totalUserCount > 0 Then
                Response.Write FormatNumber(totalPayCount / totalUserCount, 2)
            Else
                Response.Write "0"
            End If
            %>
        </td>
    </tr>

    <tr>
        <td>재구매자 평균 구매횟수</td>
        <td align="right" style="mso-number-format:'\#\,\#\#0\.00';">
            <%
            If repurchaseUserCount > 0 Then
                Response.Write FormatNumber(repurchasePayCount / repurchaseUserCount, 2)
            Else
                Response.Write "0"
            End If
            %>
        </td>
    </tr>

    <tr>
        <td>재구매자 결제액</td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(repurchaseActual, 0)%>
        </td>
    </tr>

    <tr>
        <td>재구매자 결제액 비중</td>
        <td align="right" style="mso-number-format:'0.00%';">
            <%
            If totalActual > 0 Then
                Response.Write FormatNumber(repurchaseActual / totalActual, 4)
            Else
                Response.Write "0"
            End If
            %>
        </td>
    </tr>
</table>

<br><br>

<!-- ================= 구매횟수 구간별 ================= -->
<table border="1">
    <tr>
        <td colspan="7"
            style="font-weight:bold; font-size:14px; background-color:#D3D3D3;">
            구매횟수 구간별 현황
        </td>
    </tr>

    <tr style="background-color:#4472C4; color:white; font-weight:bold;">
        <th width="100">구매횟수</th>
        <th width="120">구매자수</th>
        <th width="120">결제건수</th>
        <th width="150">실제결제액</th>
        <th width="130">1인당 평균결제액</th>
        <th width="130">쿠폰사용자수</th>
        <th width="120">쿠폰사용률</th>
    </tr>
<%
If groupMap.Count > 0 Then

    groupKeys = groupMap.Keys
    Call SortGroupKeys(groupKeys)

    For i = 0 To UBound(groupKeys)

        groupData = groupMap(groupKeys(i))

        buyerCount = CDbl(groupData(0))
        payCount = CDbl(groupData(1))
        actualAmount = CDbl(groupData(2))
        couponUsed = CDbl(groupData(3))

        If buyerCount > 0 Then
            avgAmount = actualAmount / buyerCount
            couponRate = couponUsed / buyerCount
        Else
            avgAmount = 0
            couponRate = 0
        End If
%>
    <tr>
        <td align="center"><%=Server.HTMLEncode(groupKeys(i) & "")%></td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(buyerCount, 0)%>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(payCount, 0)%>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(actualAmount, 0)%>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(avgAmount, 0)%>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(couponUsed, 0)%>
        </td>
        <td align="right" style="mso-number-format:'0.00%';">
            <%=FormatNumber(couponRate, 4)%>
        </td>
    </tr>
<%
    Next

Else
%>
    <tr>
        <td colspan="7" align="center">조회 결과가 없습니다.</td>
    </tr>
<%
End If
%>
    <tr style="background-color:#E7E6E6; font-weight:bold;">
        <td align="center">합계</td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(totalUserCount, 0)%>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(totalPayCount, 0)%>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(totalActual, 0)%>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%
            If totalUserCount > 0 Then
                Response.Write FormatNumber(totalActual / totalUserCount, 0)
            Else
                Response.Write "0"
            End If
            %>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(totalCouponUser, 0)%>
        </td>
        <td align="right" style="mso-number-format:'0.00%';">
            <%
            If totalUserCount > 0 Then
                Response.Write FormatNumber(totalCouponUser / totalUserCount, 4)
            Else
                Response.Write "0"
            End If
            %>
        </td>
    </tr>
</table>

<br><br>

<!-- ================= 재구매자 선호 과정 ================= -->
<table border="1">
    <tr>
        <td colspan="5"
            style="font-weight:bold; font-size:14px; background-color:#D3D3D3;">
            재구매자(2회 이상) 선호 과정 TOP 20
        </td>
    </tr>

    <tr style="background-color:#4472C4; color:white; font-weight:bold;">
        <th width="50">순위</th>
        <th width="300">과정명</th>
        <th width="120">구매자수</th>
        <th width="120">판매수</th>
        <th width="150">실제결제액</th>
    </tr>
<%
If IsArray(arrTopProduct) Then

    For i = 0 To UBound(arrTopProduct, 2)

        productName = Nz(arrTopProduct(0, i), "과정명 없음")
        buyerCount = CDbl(Nz(arrTopProduct(1, i), 0))
        saleCount = CDbl(Nz(arrTopProduct(2, i), 0))
        actualAmount = CDbl(Nz(arrTopProduct(3, i), 0))
%>
    <tr>
        <td align="center"><%=(i + 1)%></td>
        <td align="left"><%=Server.HTMLEncode(productName & "")%></td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(buyerCount, 0)%>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(saleCount, 0)%>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(actualAmount, 0)%>
        </td>
    </tr>
<%
    Next

Else
%>
    <tr>
        <td colspan="5" align="center">조회 결과가 없습니다.</td>
    </tr>
<%
End If
%>
</table>

</body>
</html>

<%
'==========================================================
' 구매횟수 구간명
'==========================================================
Function GetPayCountGroup(ByVal cnt)

    Dim n
    n = CDbl(cnt)

    If n <= 1 Then
        GetPayCountGroup = "1회"
    ElseIf n = 2 Then
        GetPayCountGroup = "2회"
    ElseIf n = 3 Then
        GetPayCountGroup = "3회"
    ElseIf n = 4 Then
        GetPayCountGroup = "4회"
    ElseIf n <= 9 Then
        GetPayCountGroup = "5~9회"
    Else
        GetPayCountGroup = "10회 이상"
    End If

End Function

'==========================================================
' 구간 정렬 순서
'==========================================================
Function GetPayCountGroupOrder(ByVal groupText)

    Select Case Trim(groupText & "")
        Case "1회":       GetPayCountGroupOrder = 1
        Case "2회":       GetPayCountGroupOrder = 2
        Case "3회":       GetPayCountGroupOrder = 3
        Case "4회":       GetPayCountGroupOrder = 4
        Case "5~9회":     GetPayCountGroupOrder = 5
        Case "10회 이상": GetPayCountGroupOrder = 6
        Case Else:        GetPayCountGroupOrder = 99
    End Select

End Function

Sub SortGroupKeys(ByRef arrKeys)

    Dim a
    Dim b
    Dim tempKey

    If Not IsArray(arrKeys) Then Exit Sub
    If UBound(arrKeys) < 1 Then Exit Sub

    For a = 0 To UBound(arrKeys) - 1
        For b = a + 1 To UBound(arrKeys)
            If GetPayCountGroupOrder(arrKeys(a)) > GetPayCountGroupOrder(arrKeys(b)) Then
                tempKey = arrKeys(a)
                arrKeys(a) = arrKeys(b)
                arrKeys(b) = tempKey
            End If
        Next
    Next

End Sub

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