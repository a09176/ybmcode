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

Dim productNum
Dim productName
Dim buyerCount
Dim payCount
Dim saleCount
Dim productAmount
Dim discountAmount
Dim listAmount
Dim averageAmount

Dim rank

Dim totalBuyerCount
Dim totalPayCount
Dim totalSaleCount
Dim totalListAmount
Dim totalDiscount
Dim totalProductAmount

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
' 상위 과정 TOP 100
'
' 과정 기준: product_info.product_type = 'S' (교재 제외)
' 과정 연결: pay_info_product.product_num = product_info.product_num
' 상품금액: pay_info_product.product_price 합계 (주문 실결제액과 별도 지표)
' 할인액: 단체할인 + 쿠폰할인 + 포인트사용 (pay_product_num 기준)
'==========================================================
strSQL = ""

strSQL = strSQL & ";WITH ProductSales AS ( "
strSQL = strSQL & "    SELECT "
strSQL = strSQL & "        pip.product_num, "
strSQL = strSQL & "        ISNULL(pi.product_name, '과정명 없음') AS product_name, "
strSQL = strSQL & "        COUNT(DISTINCT p.userid) AS buyer_count, "
strSQL = strSQL & "        COUNT(DISTINCT p.pay_num) AS pay_count, "
strSQL = strSQL & "        COUNT(DISTINCT pip.pay_product_num) AS sale_count, "
strSQL = strSQL & "        SUM(ISNULL(pip.product_price, 0)) AS product_amount "
strSQL = strSQL & "    FROM pay_info AS p WITH (READUNCOMMITTED) "
strSQL = strSQL & "    INNER JOIN pay_info_product AS pip WITH (READUNCOMMITTED) "
strSQL = strSQL & "        ON p.pay_num = pip.pay_num "
strSQL = strSQL & "    INNER JOIN product_info AS pi WITH (READUNCOMMITTED) "
strSQL = strSQL & "        ON pip.product_num = pi.product_num "
strSQL = strSQL & "    WHERE p.sell_info = 'o' "
strSQL = strSQL & "      AND pi.product_type = 'S' "
strSQL = strSQL & "      AND p.pay_date >= '" & startDate & "' "
strSQL = strSQL & "      AND p.pay_date < DATEADD(DAY, 1, '" & endDate & "') "
strSQL = strSQL & "    GROUP BY "
strSQL = strSQL & "        pip.product_num, "
strSQL = strSQL & "        pi.product_name "
strSQL = strSQL & "), "
strSQL = strSQL & "PartyDiscount AS ( "
strSQL = strSQL & "    SELECT "
strSQL = strSQL & "        pip.product_num, "
strSQL = strSQL & "        SUM(ISNULL(pd.discount_price, 0)) AS party_discount "
strSQL = strSQL & "    FROM pay_info AS p WITH (READUNCOMMITTED) "
strSQL = strSQL & "    INNER JOIN pay_info_product AS pip WITH (READUNCOMMITTED) "
strSQL = strSQL & "        ON p.pay_num = pip.pay_num "
strSQL = strSQL & "    INNER JOIN product_info AS pi WITH (READUNCOMMITTED) "
strSQL = strSQL & "        ON pip.product_num = pi.product_num "
strSQL = strSQL & "    INNER JOIN pay_discount AS pd WITH (READUNCOMMITTED) "
strSQL = strSQL & "        ON pip.pay_product_num = pd.pay_product_num "
strSQL = strSQL & "    WHERE p.sell_info = 'o' "
strSQL = strSQL & "      AND pi.product_type = 'S' "
strSQL = strSQL & "      AND p.pay_date >= '" & startDate & "' "
strSQL = strSQL & "      AND p.pay_date < DATEADD(DAY, 1, '" & endDate & "') "
strSQL = strSQL & "    GROUP BY pip.product_num "
strSQL = strSQL & "), "
strSQL = strSQL & "CouponDiscount AS ( "
strSQL = strSQL & "    SELECT "
strSQL = strSQL & "        pip.product_num, "
strSQL = strSQL & "        SUM(ISNULL(pcu.coupon_discount, 0)) AS coupon_discount "
strSQL = strSQL & "    FROM pay_info AS p WITH (READUNCOMMITTED) "
strSQL = strSQL & "    INNER JOIN pay_info_product AS pip WITH (READUNCOMMITTED) "
strSQL = strSQL & "        ON p.pay_num = pip.pay_num "
strSQL = strSQL & "    INNER JOIN product_info AS pi WITH (READUNCOMMITTED) "
strSQL = strSQL & "        ON pip.product_num = pi.product_num "
strSQL = strSQL & "    INNER JOIN pay_coupon_use AS pcu WITH (READUNCOMMITTED) "
strSQL = strSQL & "        ON pip.pay_product_num = pcu.pay_product_num "
strSQL = strSQL & "    WHERE p.sell_info = 'o' "
strSQL = strSQL & "      AND pi.product_type = 'S' "
strSQL = strSQL & "      AND p.pay_date >= '" & startDate & "' "
strSQL = strSQL & "      AND p.pay_date < DATEADD(DAY, 1, '" & endDate & "') "
strSQL = strSQL & "    GROUP BY pip.product_num "
strSQL = strSQL & "), "
strSQL = strSQL & "PointDiscount AS ( "
strSQL = strSQL & "    SELECT "
strSQL = strSQL & "        pip.product_num, "
strSQL = strSQL & "        SUM(ISNULL(ppu.point_price, 0)) AS point_discount "
strSQL = strSQL & "    FROM pay_info AS p WITH (READUNCOMMITTED) "
strSQL = strSQL & "    INNER JOIN pay_info_product AS pip WITH (READUNCOMMITTED) "
strSQL = strSQL & "        ON p.pay_num = pip.pay_num "
strSQL = strSQL & "    INNER JOIN product_info AS pi WITH (READUNCOMMITTED) "
strSQL = strSQL & "        ON pip.product_num = pi.product_num "
strSQL = strSQL & "    INNER JOIN Pay_Point_Use AS ppu WITH (READUNCOMMITTED) "
strSQL = strSQL & "        ON pip.pay_product_num = ppu.pay_product_num "
strSQL = strSQL & "    WHERE p.sell_info = 'o' "
strSQL = strSQL & "      AND pi.product_type = 'S' "
strSQL = strSQL & "      AND p.pay_date >= '" & startDate & "' "
strSQL = strSQL & "      AND p.pay_date < DATEADD(DAY, 1, '" & endDate & "') "
strSQL = strSQL & "    GROUP BY pip.product_num "
strSQL = strSQL & ") "
strSQL = strSQL & "SELECT TOP 100 "
strSQL = strSQL & "    ps.product_num, "
strSQL = strSQL & "    ps.product_name, "
strSQL = strSQL & "    ps.buyer_count, "
strSQL = strSQL & "    ps.pay_count, "
strSQL = strSQL & "    ps.sale_count, "
strSQL = strSQL & "    ps.product_amount, "
strSQL = strSQL & "    ISNULL(pd.party_discount, 0) "
strSQL = strSQL & "        + ISNULL(cd.coupon_discount, 0) "
strSQL = strSQL & "        + ISNULL(pt.point_discount, 0) AS discount_amount, "
strSQL = strSQL & "    ps.product_amount "
strSQL = strSQL & "        + ISNULL(pd.party_discount, 0) "
strSQL = strSQL & "        + ISNULL(cd.coupon_discount, 0) "
strSQL = strSQL & "        + ISNULL(pt.point_discount, 0) AS original_amount "
strSQL = strSQL & "FROM ProductSales AS ps "
strSQL = strSQL & "LEFT JOIN PartyDiscount AS pd "
strSQL = strSQL & "    ON ps.product_num = pd.product_num "
strSQL = strSQL & "LEFT JOIN CouponDiscount AS cd "
strSQL = strSQL & "    ON ps.product_num = cd.product_num "
strSQL = strSQL & "LEFT JOIN PointDiscount AS pt "
strSQL = strSQL & "    ON ps.product_num = pt.product_num "
strSQL = strSQL & "ORDER BY "
strSQL = strSQL & "    ps.sale_count DESC, "
strSQL = strSQL & "    ps.buyer_count DESC, "
strSQL = strSQL & "    ps.product_amount DESC, "
strSQL = strSQL & "    ps.product_num ASC "

arrResult = ExecSP(strSQL, strCyberTeacher)

'==========================================================
' 개인정보 파일 다운로드 로그
'==========================================================
If IsArray(arrResult) Then

    strSQL = ""
    strSQL = strSQL & "EXEC site_log.dbo.p_File_Log "
    strSQL = strSQL & fnCdbValue(Session("admin_id")) & ","
    strSQL = strSQL & fnCdbValue(Request.ServerVariables("REMOTE_ADDR")) & ","
    strSQL = strSQL & fnCdbValue(Request.ServerVariables("HTTP_HOST")) & ","
    strSQL = strSQL & fnCdbValue("stat_top_products_" & downloadDate) & ","
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
    "attachment; filename=stat_top_products_" & _
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
            <th width="50">순위</th>
            <th width="250">과정명</th>
            <th width="100">구매자수</th>
            <th width="100">결제건수</th>
            <th width="100">판매수</th>
            <th width="150">할인 전 상품금액</th>
            <th width="150">할인액</th>
            <th width="150">상품금액(product_price)</th>
            <th width="150">1인당 상품금액</th>
        </tr>
    </thead>

    <tbody>
<%
totalBuyerCount = 0
totalPayCount = 0
totalSaleCount = 0
totalListAmount = 0
totalDiscount = 0
totalProductAmount = 0
rank = 0

If IsArray(arrResult) Then

    For i = 0 To UBound(arrResult, 2)

        rank = i + 1

        productNum = Nz(arrResult(0, i), 0)
        productName = Nz(arrResult(1, i), "과정명 없음")

        buyerCount = CLng(Nz(arrResult(2, i), 0))
        payCount = CLng(Nz(arrResult(3, i), 0))
        saleCount = CLng(Nz(arrResult(4, i), 0))
        productAmount = CDbl(Nz(arrResult(5, i), 0))
        discountAmount = CDbl(Nz(arrResult(6, i), 0))
        listAmount = CDbl(Nz(arrResult(7, i), 0))

        If buyerCount > 0 Then
            averageAmount = productAmount / buyerCount
        Else
            averageAmount = 0
        End If

        totalBuyerCount = totalBuyerCount + buyerCount
        totalPayCount = totalPayCount + payCount
        totalSaleCount = totalSaleCount + saleCount
        totalListAmount = totalListAmount + listAmount
        totalDiscount = totalDiscount + discountAmount
        totalProductAmount = totalProductAmount + productAmount
%>
        <tr>
            <td align="center"
                style="mso-number-format:'\#\,\#\#0';">
                <%=rank%>
            </td>

            <td align="left">
                <%=Server.HTMLEncode(productName & "")%>
            </td>

            <td align="right"
                style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(buyerCount, 0)%>
            </td>

            <td align="right"
                style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(payCount, 0)%>
            </td>

            <td align="right"
                style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(saleCount, 0)%>
            </td>

            <td align="right"
                style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(listAmount, 0)%>
            </td>

            <td align="right"
                style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(discountAmount, 0)%>
            </td>

            <td align="right"
                style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(productAmount, 0)%>
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
            <td colspan="9" align="center">
                조회 결과가 없습니다.
            </td>
        </tr>
<%
End If
%>
    </tbody>

    <tfoot>
        <tr style="background-color: #E7E6E6; font-weight: bold;">
            <td colspan="2" align="center">
                전체 합계
            </td>

            <td align="right"
                style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(totalBuyerCount, 0)%>
            </td>

            <td align="right"
                style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(totalPayCount, 0)%>
            </td>

            <td align="right"
                style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(totalSaleCount, 0)%>
            </td>

            <td align="right"
                style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(totalListAmount, 0)%>
            </td>

            <td align="right"
                style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(totalDiscount, 0)%>
            </td>

            <td align="right"
                style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(totalProductAmount, 0)%>
            </td>

            <td align="right"
                style="mso-number-format:'\#\,\#\#0';">
                <%
                If totalBuyerCount > 0 Then
                    Response.Write FormatNumber(totalProductAmount / totalBuyerCount, 0)
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