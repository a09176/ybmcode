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

Dim strSQL, arrResult, i
Dim sYear, sMonth, eYear, eMonth, sDay, eDay
Dim startDate, endDate, downloadDate
Dim totalOrigB2C, totalOrigB2B, totalDiscF, totalDiscC, totalDiscP
Dim totalActualB2C, totalActualB2B, totalPayCount, totalBuyerCount

sYear = Trim(Request("sYear"))
sMonth = Trim(Request("sMonth"))
eYear = Trim(Request("eYear"))
eMonth = Trim(Request("eMonth"))
sDay = Trim(Request("sDay"))
eDay = Trim(Request("eDay"))

If sYear = "" Then sYear = "2025"
If sMonth = "" Then sMonth = "01"
If eYear = "" Then eYear = Year(Now())
If eMonth = "" Then eMonth = Month(Now())
If sDay = "" Then sDay = "01"
If eDay = "" Then eDay = Day(Now())

startDate = sYear & "-" & Right("0" & sMonth, 2) & "-" & Right("0" & sDay, 2)
endDate = eYear & "-" & Right("0" & eMonth, 2) & "-" & Right("0" & eDay, 2)
downloadDate = Year(Now()) & Right("0" & Month(Now()), 2) & Right("0" & Day(Now()), 2)

'==========================================================
' 월별 매출 통계 쿼리
' 할인 테이블은 pay_info_product를 통해 pay_num으로 집계
'==========================================================
strSQL = ""
strSQL = strSQL & ";WITH OrderBiz AS ( "
strSQL = strSQL & "    SELECT "
strSQL = strSQL & "        pip.pay_num, "
strSQL = strSQL & "        CASE "
strSQL = strSQL & "            WHEN MAX(CASE WHEN sd.comp_yn = 'y' THEN 1 ELSE 0 END) = 1 "
strSQL = strSQL & "                THEN 'B2B' "
strSQL = strSQL & "            ELSE 'B2C' "
strSQL = strSQL & "        END AS biz_type "
strSQL = strSQL & "    FROM pay_info_product AS pip WITH (READUNCOMMITTED) "
strSQL = strSQL & "    LEFT JOIN schedule_degree AS sd WITH (READUNCOMMITTED) "
strSQL = strSQL & "        ON pip.schedule_num = sd.schedule_num "
strSQL = strSQL & "    GROUP BY pip.pay_num "
strSQL = strSQL & "), "
strSQL = strSQL & "PartyDiscount AS ( "
strSQL = strSQL & "    SELECT "
strSQL = strSQL & "        pip.pay_num, "
strSQL = strSQL & "        SUM(ISNULL(pd.discount_price, 0)) AS party_discount "
strSQL = strSQL & "    FROM pay_info_product AS pip WITH (READUNCOMMITTED) "
strSQL = strSQL & "    INNER JOIN pay_discount AS pd WITH (READUNCOMMITTED) "
strSQL = strSQL & "        ON pip.pay_product_num = pd.pay_product_num "
strSQL = strSQL & "    GROUP BY pip.pay_num "
strSQL = strSQL & "), "
strSQL = strSQL & "CouponDiscount AS ( "
strSQL = strSQL & "    SELECT "
strSQL = strSQL & "        pip.pay_num, "
strSQL = strSQL & "        SUM(ISNULL(pcu.coupon_discount, 0)) AS coupon_discount "
strSQL = strSQL & "    FROM pay_info_product AS pip WITH (READUNCOMMITTED) "
strSQL = strSQL & "    INNER JOIN pay_coupon_use AS pcu WITH (READUNCOMMITTED) "
strSQL = strSQL & "        ON pip.pay_product_num = pcu.pay_product_num "
strSQL = strSQL & "    GROUP BY pip.pay_num "
strSQL = strSQL & "), "
strSQL = strSQL & "PointDiscount AS ( "
strSQL = strSQL & "    SELECT "
strSQL = strSQL & "        pip.pay_num, "
strSQL = strSQL & "        SUM(ISNULL(ppu.point_price, 0)) AS point_discount "
strSQL = strSQL & "    FROM pay_info_product AS pip WITH (READUNCOMMITTED) "
strSQL = strSQL & "    INNER JOIN Pay_Point_Use AS ppu WITH (READUNCOMMITTED) "
strSQL = strSQL & "        ON pip.pay_product_num = ppu.pay_product_num "
strSQL = strSQL & "    GROUP BY pip.pay_num "
strSQL = strSQL & "), "
strSQL = strSQL & "OrderBase AS ( "
strSQL = strSQL & "    SELECT "
strSQL = strSQL & "        p.pay_num, "
strSQL = strSQL & "        p.userid, "
strSQL = strSQL & "        p.pay_date, "
strSQL = strSQL & "        ISNULL(ob.biz_type, "
strSQL = strSQL & "            CASE WHEN ISNULL(p.site_code, '') = '' "
strSQL = strSQL & "                 THEN 'B2C' ELSE 'B2B' END "
strSQL = strSQL & "        ) AS biz_type, "
strSQL = strSQL & "        ISNULL(p.pay_price, 0) AS actual_amount, "
strSQL = strSQL & "        ISNULL(pd.party_discount, 0) AS party_discount, "
strSQL = strSQL & "        ISNULL(cd.coupon_discount, 0) AS coupon_discount, "
strSQL = strSQL & "        ISNULL(pt.point_discount, 0) AS point_discount "
strSQL = strSQL & "    FROM pay_info AS p WITH (READUNCOMMITTED) "
strSQL = strSQL & "    LEFT JOIN OrderBiz AS ob "
strSQL = strSQL & "        ON p.pay_num = ob.pay_num "
strSQL = strSQL & "    LEFT JOIN PartyDiscount AS pd "
strSQL = strSQL & "        ON p.pay_num = pd.pay_num "
strSQL = strSQL & "    LEFT JOIN CouponDiscount AS cd "
strSQL = strSQL & "        ON p.pay_num = cd.pay_num "
strSQL = strSQL & "    LEFT JOIN PointDiscount AS pt "
strSQL = strSQL & "        ON p.pay_num = pt.pay_num "
strSQL = strSQL & "    WHERE p.sell_info = 'o' "
strSQL = strSQL & "      AND p.pay_date >= '" & startDate & "' "
strSQL = strSQL & "      AND p.pay_date < DATEADD(DAY, 1, '" & endDate & "') "
strSQL = strSQL & ") "
strSQL = strSQL & "SELECT "
strSQL = strSQL & "    LEFT(pay_date, 7) AS ym, "
strSQL = strSQL & "    SUM(CASE WHEN biz_type = 'B2C' "
strSQL = strSQL & "        THEN actual_amount + party_discount "
strSQL = strSQL & "             + coupon_discount + point_discount "
strSQL = strSQL & "        ELSE 0 END) AS original_b2c, "
strSQL = strSQL & "    SUM(CASE WHEN biz_type = 'B2B' "
strSQL = strSQL & "        THEN actual_amount + party_discount "
strSQL = strSQL & "             + coupon_discount + point_discount "
strSQL = strSQL & "        ELSE 0 END) AS original_b2b, "
strSQL = strSQL & "    SUM(party_discount) AS party_discount, "
strSQL = strSQL & "    SUM(coupon_discount) AS coupon_discount, "
strSQL = strSQL & "    SUM(point_discount) AS point_discount, "
strSQL = strSQL & "    SUM(CASE WHEN biz_type = 'B2C' "
strSQL = strSQL & "        THEN actual_amount ELSE 0 END) AS actual_b2c, "
strSQL = strSQL & "    SUM(CASE WHEN biz_type = 'B2B' "
strSQL = strSQL & "        THEN actual_amount ELSE 0 END) AS actual_b2b, "
strSQL = strSQL & "    COUNT(DISTINCT pay_num) AS pay_count, "
strSQL = strSQL & "    COUNT(DISTINCT userid) AS buyer_count "
strSQL = strSQL & "FROM OrderBase "
strSQL = strSQL & "GROUP BY LEFT(pay_date, 7) "
strSQL = strSQL & "ORDER BY LEFT(pay_date, 7) "

arrResult = ExecSP(strSQL, strCyberTeacher)

If IsArray(arrResult) Then
    strSQL = ""
    strSQL = strSQL & "EXEC site_log.dbo.p_File_Log "
    strSQL = strSQL & fnCdbValue(Session("admin_id")) & ","
    strSQL = strSQL & fnCdbValue(Request.ServerVariables("REMOTE_ADDR")) & ","
    strSQL = strSQL & fnCdbValue(Request.ServerVariables("HTTP_HOST")) & ","
    strSQL = strSQL & fnCdbValue("stat_monthly_sales_" & downloadDate) & ","
    strSQL = strSQL & fnCdbValue(Request.ServerVariables("URL")) & ","
    strSQL = strSQL & fnCdbValue(Session("admin_name")) & ","
    strSQL = strSQL & "'','','' "
    Call ExecSQL(strSQL, e4u2006DBStr)
End If

Response.ContentType = "application/vnd.ms-excel"
Response.CharSet = "euc-kr"
Response.AddHeader "Content-Disposition", "attachment; filename=stat_monthly_sales_" & downloadDate & ".xls"
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
</head>
<body>

<table border="1" style="border-collapse: collapse;">
    <thead>
        <tr style="background-color: #4472C4; color: white; font-weight: bold;">
            <th rowspan="2">연월</th>
            <th colspan="2">원금액</th>
            <th colspan="3">할인액</th>
            <th colspan="2">실결제액</th>
            <th rowspan="2">결제건수</th>
            <th rowspan="2">월별 고유 구매자수<br>(합계는 월별 합산)</th>
            <th rowspan="2">평균구매액</th>
        </tr>
        <tr style="background-color: #4472C4; color: white; font-weight: bold;">
            <th>B2C</th>
            <th>B2B</th>
            <th>단체</th>
            <th>쿠폰</th>
            <th>포인트</th>
            <th>B2C</th>
            <th>B2B</th>
        </tr>
    </thead>
    <tbody>
<%
Dim ymKey, b2cOrig, b2bOrig, partyDiscount, couponDiscount, pointDiscount
Dim b2cActual, b2bActual, payCount, buyerCount, actualTotal, avgPrice

totalOrigB2C = 0
totalOrigB2B = 0
totalDiscF = 0
totalDiscC = 0
totalDiscP = 0
totalActualB2C = 0
totalActualB2B = 0
totalPayCount = 0
totalBuyerCount = 0

If IsArray(arrResult) Then
    For i = 0 To UBound(arrResult, 2)
        ymKey = Nz(arrResult(0, i), "")
        b2cOrig = CLng(Nz(arrResult(1, i), 0))
        b2bOrig = CLng(Nz(arrResult(2, i), 0))
        partyDiscount = CLng(Nz(arrResult(3, i), 0))
        couponDiscount = CLng(Nz(arrResult(4, i), 0))
        pointDiscount = CLng(Nz(arrResult(5, i), 0))
        b2cActual = CLng(Nz(arrResult(6, i), 0))
        b2bActual = CLng(Nz(arrResult(7, i), 0))
        payCount = CLng(Nz(arrResult(8, i), 0))
        buyerCount = CLng(Nz(arrResult(9, i), 0))
        
        actualTotal = b2cActual + b2bActual
        If buyerCount > 0 Then
            avgPrice = actualTotal / buyerCount
        Else
            avgPrice = 0
        End If
        
        totalOrigB2C = totalOrigB2C + b2cOrig
        totalOrigB2B = totalOrigB2B + b2bOrig
        totalDiscF = totalDiscF + partyDiscount
        totalDiscC = totalDiscC + couponDiscount
        totalDiscP = totalDiscP + pointDiscount
        totalActualB2C = totalActualB2C + b2cActual
        totalActualB2B = totalActualB2B + b2bActual
        totalPayCount = totalPayCount + payCount
        totalBuyerCount = totalBuyerCount + buyerCount
%>
        <tr>
            <td align="center" style="mso-number-format:'\@';">
                <%=Server.HTMLEncode(ymKey & "")%>
            </td>
            <td align="right" style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(b2cOrig, 0)%>
            </td>
            <td align="right" style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(b2bOrig, 0)%>
            </td>
            <td align="right" style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(partyDiscount, 0)%>
            </td>
            <td align="right" style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(couponDiscount, 0)%>
            </td>
            <td align="right" style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(pointDiscount, 0)%>
            </td>
            <td align="right" style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(b2cActual, 0)%>
            </td>
            <td align="right" style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(b2bActual, 0)%>
            </td>
            <td align="right" style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(payCount, 0)%>
            </td>
            <td align="right" style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(buyerCount, 0)%>
            </td>
            <td align="right" style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(avgPrice, 0)%>
            </td>
        </tr>
<%
    Next
Else
%>
        <tr>
            <td colspan="11" align="center">
                조회 결과가 없습니다.
            </td>
        </tr>
<%
End If
%>
    </tbody>
    <tfoot>
        <tr style="background-color: #E7E6E6; font-weight: bold;">
            <td align="center">합계</td>
            <td align="right" style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(totalOrigB2C, 0)%>
            </td>
            <td align="right" style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(totalOrigB2B, 0)%>
            </td>
            <td align="right" style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(totalDiscF, 0)%>
            </td>
            <td align="right" style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(totalDiscC, 0)%>
            </td>
            <td align="right" style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(totalDiscP, 0)%>
            </td>
            <td align="right" style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(totalActualB2C, 0)%>
            </td>
            <td align="right" style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(totalActualB2B, 0)%>
            </td>
            <td align="right" style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(totalPayCount, 0)%>
            </td>
            <td align="right" style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(totalBuyerCount, 0)%>
            </td>
            <td align="right" style="mso-number-format:'\#\,\#\#0';">
                <%
                If totalBuyerCount > 0 Then
                    Response.Write FormatNumber((totalActualB2C + totalActualB2B) / totalBuyerCount, 0)
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
Function Nz(ByVal value, ByVal defaultValue)
    If IsNull(value) Or Trim(value & "") = "" Then
        Nz = defaultValue
    Else
        Nz = value
    End If
End Function
%>