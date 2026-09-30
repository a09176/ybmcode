
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
Dim arrUserCoupon
Dim arrTopProductWithCoupon
Dim arrTopProductWithoutCoupon

Dim i

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

Dim userCount
Dim totalPayCount
Dim totalActual
Dim couponUserCount
Dim noCouponUserCount
Dim couponTotalActual
Dim noCouponTotalActual
Dim couponPayCount
Dim noCouponPayCount

Dim avgCouponAmount
Dim avgNoCouponAmount
Dim couponRepurchaseRate
Dim noCouponRepurchaseRate
Dim couponRepurchaseCount
Dim noCouponRepurchaseCount

Dim productName
Dim buyerCount
Dim saleCount

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

arrUserCoupon = ExecSP(strSQL, strCyberTeacher)

'==========================================================
' 2. 쿠폰 사용/미사용 비교 분석
'==========================================================
userCount = 0
totalPayCount = 0
totalActual = 0
couponUserCount = 0
noCouponUserCount = 0
couponTotalActual = 0
noCouponTotalActual = 0
couponPayCount = 0
noCouponPayCount = 0
couponRepurchaseCount = 0
noCouponRepurchaseCount = 0

If IsArray(arrUserCoupon) Then

    For i = 0 To UBound(arrUserCoupon, 2)

        userid = Trim(Nz(arrUserCoupon(0, i), "") & "")
        payCount = CDbl(Nz(arrUserCoupon(1, i), 0))
        actualAmount = CDbl(Nz(arrUserCoupon(2, i), 0))
        couponUsed = CDbl(Nz(arrUserCoupon(3, i), 0))

        If userid <> "" Then

            userCount = userCount + 1
            totalPayCount = totalPayCount + payCount
            totalActual = totalActual + actualAmount

            If couponUsed > 0 Then

                couponUserCount = couponUserCount + 1
                couponTotalActual = couponTotalActual + actualAmount
                couponPayCount = couponPayCount + payCount

                If payCount >= 2 Then
                    couponRepurchaseCount = couponRepurchaseCount + 1
                End If

            Else

                noCouponUserCount = noCouponUserCount + 1
                noCouponTotalActual = noCouponTotalActual + actualAmount
                noCouponPayCount = noCouponPayCount + payCount

                If payCount >= 2 Then
                    noCouponRepurchaseCount = noCouponRepurchaseCount + 1
                End If

            End If

        End If

    Next

End If

'==========================================================
' 3. 쿠폰 사용자의 선호 과정 TOP 20
'
' 기간 내 쿠폰을 사용한 회원들이 구매한 과정
' 교재는 제외
'==========================================================
strSQL = ""
strSQL = strSQL & ";WITH CouponUser AS ( "
strSQL = strSQL & "    SELECT DISTINCT p.userid "
strSQL = strSQL & "    FROM pay_info AS p WITH (READUNCOMMITTED) "
strSQL = strSQL & "    INNER JOIN pay_info_product AS pip WITH (READUNCOMMITTED) "
strSQL = strSQL & "        ON p.pay_num = pip.pay_num "
strSQL = strSQL & "    INNER JOIN pay_coupon_use AS cu WITH (READUNCOMMITTED) "
strSQL = strSQL & "        ON pip.pay_product_num = cu.pay_product_num "
strSQL = strSQL & "    WHERE p.sell_info = 'o' "
strSQL = strSQL & "      AND p.pay_date >= '" & startDate & "' "
strSQL = strSQL & "      AND p.pay_date <= '" & endDate & "' "
strSQL = strSQL & ") "
strSQL = strSQL & "SELECT TOP 20 "
strSQL = strSQL & "    ISNULL(pi.product_name, '과정명 없음') AS product_name, "
strSQL = strSQL & "    COUNT(DISTINCT p.userid) AS buyer_count, "
strSQL = strSQL & "    COUNT(DISTINCT pip.pay_product_num) AS sale_count, "
strSQL = strSQL & "    SUM(ISNULL(pip.product_price, 0)) AS actual_amount "
strSQL = strSQL & "FROM pay_info AS p WITH (READUNCOMMITTED) "
strSQL = strSQL & "INNER JOIN CouponUser AS cu_user "
strSQL = strSQL & "    ON p.userid = cu_user.userid "
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

arrTopProductWithCoupon = ExecSP(strSQL, strCyberTeacher)

'==========================================================
' 4. 쿠폰 미사용자의 선호 과정 TOP 20
'
' 기간 내 쿠폰을 사용하지 않은 회원들이 구매한 과정
' 교재는 제외
'==========================================================
strSQL = ""
strSQL = strSQL & ";WITH NoCouponUser AS ( "
strSQL = strSQL & "    SELECT DISTINCT p.userid "
strSQL = strSQL & "    FROM pay_info AS p WITH (READUNCOMMITTED) "
strSQL = strSQL & "    INNER JOIN pay_info_product AS pip WITH (READUNCOMMITTED) "
strSQL = strSQL & "        ON p.pay_num = pip.pay_num "
strSQL = strSQL & "    LEFT JOIN pay_coupon_use AS cu WITH (READUNCOMMITTED) "
strSQL = strSQL & "        ON pip.pay_product_num = cu.pay_product_num "
strSQL = strSQL & "    WHERE p.sell_info = 'o' "
strSQL = strSQL & "      AND p.pay_date >= '" & startDate & "' "
strSQL = strSQL & "      AND p.pay_date <= '" & endDate & "' "
strSQL = strSQL & "      AND cu.pay_product_num IS NULL "
strSQL = strSQL & ") "
strSQL = strSQL & "SELECT TOP 20 "
strSQL = strSQL & "    ISNULL(pi.product_name, '과정명 없음') AS product_name, "
strSQL = strSQL & "    COUNT(DISTINCT p.userid) AS buyer_count, "
strSQL = strSQL & "    COUNT(DISTINCT pip.pay_product_num) AS sale_count, "
strSQL = strSQL & "    SUM(ISNULL(pip.product_price, 0)) AS actual_amount "
strSQL = strSQL & "FROM pay_info AS p WITH (READUNCOMMITTED) "
strSQL = strSQL & "INNER JOIN NoCouponUser AS ncu_user "
strSQL = strSQL & "    ON p.userid = ncu_user.userid "
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

arrTopProductWithoutCoupon = ExecSP(strSQL, strCyberTeacher)

'==========================================================
' 개인정보 파일 다운로드 로그
'==========================================================
If IsArray(arrUserCoupon) Then

    strSQL = ""
    strSQL = strSQL & "EXEC site_log.dbo.p_File_Log "
    strSQL = strSQL & fnCdbValue(Session("admin_id")) & ","
    strSQL = strSQL & fnCdbValue(Request.ServerVariables("REMOTE_ADDR")) & ","
    strSQL = strSQL & fnCdbValue(Request.ServerVariables("HTTP_HOST")) & ","
    strSQL = strSQL & fnCdbValue("stat_coupon_effect_" & downloadDate) & ","
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
    "attachment; filename=stat_coupon_effect_" & _
    downloadDate & ".xls"
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
</head>

<body>

<!-- ================= 조회 기간 정보 ================= -->
<div style="margin-bottom: 30px;">
    <h3>쿠폰 효과 분석</h3>
    <p>
        <strong>조회 기간:</strong> 
        <%=Server.HTMLEncode(startDate)%> ~ <%=Server.HTMLEncode(endDate)%>
    </p>
    <p style="color: #666; font-size: 12px;">
        기간 내 결제한 회원을 쿠폰 사용 여부로 분류하여 구매 패턴, 재구매율, 평균 구매액을 비교합니다.<br>
        쿠폰 사용 여부는 한 건의 결제라도 쿠폰이 적용되었으면 "사용"으로 분류합니다.
    </p>
</div>

<%
If userCount = 0 Then

    Response.Write "<table border=""1"">" & _
                   "<tr><td>조회 결과가 없습니다.</td></tr>" & _
                   "</table>"

Else
%>

<!-- ================= 쿠폰 사용/미사용 비교 ================= -->
<table border="1" style="margin-bottom: 30px;">
    <tr>
        <td colspan="8"
            style="font-weight:bold; font-size:14px; background-color:#D3D3D3;">
            쿠폰 사용 여부별 구매 현황 비교
        </td>
    </tr>

    <tr style="background-color:#4472C4; color:white; font-weight:bold;">
        <th width="120">구분</th>
        <th width="120">회원수</th>
        <th width="120">결제건수</th>
        <th width="150">총 결제액</th>
        <th width="150">1인당 평균 결제액</th>
        <th width="120">재구매자수</th>
        <th width="120">재구매율</th>
        <th width="120">비중</th>
    </tr>

    <tr>
        <td align="center" style="font-weight:bold;">쿠폰 사용</td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(couponUserCount, 0)%>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(couponPayCount, 0)%>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(couponTotalActual, 0)%>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%
            If couponUserCount > 0 Then
                avgCouponAmount = couponTotalActual / couponUserCount
                Response.Write FormatNumber(avgCouponAmount, 0)
            Else
                Response.Write "0"
            End If
            %>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(couponRepurchaseCount, 0)%>
        </td>
        <td align="right" style="mso-number-format:'0.00%';">
            <%
            If couponUserCount > 0 Then
                couponRepurchaseRate = couponRepurchaseCount / couponUserCount
                Response.Write FormatNumber(couponRepurchaseRate, 4)
            Else
                Response.Write "0"
            End If
            %>
        </td>
        <td align="right" style="mso-number-format:'0.00%';">
            <%
            If userCount > 0 Then
                Response.Write FormatNumber(couponUserCount / userCount, 4)
            Else
                Response.Write "0"
            End If
            %>
        </td>
    </tr>

    <tr>
        <td align="center" style="font-weight:bold;">쿠폰 미사용</td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(noCouponUserCount, 0)%>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(noCouponPayCount, 0)%>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(noCouponTotalActual, 0)%>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%
            If noCouponUserCount > 0 Then
                avgNoCouponAmount = noCouponTotalActual / noCouponUserCount
                Response.Write FormatNumber(avgNoCouponAmount, 0)
            Else
                Response.Write "0"
            End If
            %>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(noCouponRepurchaseCount, 0)%>
        </td>
        <td align="right" style="mso-number-format:'0.00%';">
            <%
            If noCouponUserCount > 0 Then
                noCouponRepurchaseRate = noCouponRepurchaseCount / noCouponUserCount
                Response.Write FormatNumber(noCouponRepurchaseRate, 4)
            Else
                Response.Write "0"
            End If
            %>
        </td>
        <td align="right" style="mso-number-format:'0.00%';">
            <%
            If userCount > 0 Then
                Response.Write FormatNumber(noCouponUserCount / userCount, 4)
            Else
                Response.Write "0"
            End If
            %>
        </td>
    </tr>

    <tr style="background-color:#E7E6E6; font-weight:bold;">
        <td align="center">전체</td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(userCount, 0)%>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(totalPayCount, 0)%>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(totalActual, 0)%>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%
            If userCount > 0 Then
                Response.Write FormatNumber(totalActual / userCount, 0)
            Else
                Response.Write "0"
            End If
            %>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(couponRepurchaseCount + noCouponRepurchaseCount, 0)%>
        </td>
        <td align="right" style="mso-number-format:'0.00%';">
            <%
            If userCount > 0 Then
                Response.Write FormatNumber( _
                    (couponRepurchaseCount + noCouponRepurchaseCount) / userCount, 4)
            Else
                Response.Write "0"
            End If
            %>
        </td>
        <td align="right" style="mso-number-format:'0.00%';">100%</td>
    </tr>
</table>

<hr style="margin: 40px 0; border: 1px solid #999;">

<!-- ================= 쿠폰 사용자 선호 과정 ================= -->
<table border="1" style="margin-bottom: 30px;">
    <tr>
        <td colspan="5"
            style="font-weight:bold; font-size:14px; background-color:#D3D3D3;">
            쿠폰 사용자의 선호 과정 TOP 20
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
    If IsArray(arrTopProductWithCoupon) Then

        For i = 0 To UBound(arrTopProductWithCoupon, 2)

            productName = Nz(arrTopProductWithCoupon(0, i), "과정명 없음")
            buyerCount = CDbl(Nz(arrTopProductWithCoupon(1, i), 0))
            saleCount = CDbl(Nz(arrTopProductWithCoupon(2, i), 0))
            actualAmount = CDbl(Nz(arrTopProductWithCoupon(3, i), 0))
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

<hr style="margin: 40px 0; border: 1px solid #999;">

<!-- ================= 쿠폰 미사용자 선호 과정 ================= -->
<table border="1">
    <tr>
        <td colspan="5"
            style="font-weight:bold; font-size:14px; background-color:#D3D3D3;">
            쿠폰 미사용자의 선호 과정 TOP 20
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
    If IsArray(arrTopProductWithoutCoupon) Then

        For i = 0 To UBound(arrTopProductWithoutCoupon, 2)

            productName = Nz(arrTopProductWithoutCoupon(0, i), "과정명 없음")
            buyerCount = CDbl(Nz(arrTopProductWithoutCoupon(1, i), 0))
            saleCount = CDbl(Nz(arrTopProductWithoutCoupon(2, i), 0))
            actualAmount = CDbl(Nz(arrTopProductWithoutCoupon(3, i), 0))
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

<%
End If
%>

</body>
</html>

<%
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