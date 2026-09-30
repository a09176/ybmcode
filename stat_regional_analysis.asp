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

Dim totalBuyerCount
Dim totalActual
Dim totalSmsCount

Dim payYear
Dim regionName
Dim buyerCount
Dim actualAmount
Dim smsCount
Dim averageAmount
Dim smsRate

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
' 연도별·지역별 구매자 통계
'
' 지역 연결:
' pay_info.userid
'   -> vUser_info.userid
'   -> school_info.school_num
'   -> school_category.scate_num
'
' 결제금액:
' pay_info.pay_price를 주문 단위로 집계
' pay_info_product를 직접 조인하지 않아 결제금액 중복 방지
'
' SMS:
' info_receive_agree에서 agree_type = 'SMS'인 회원
'==========================================================
strSQL = ""

strSQL = strSQL & ";WITH UserRegion AS ( "
strSQL = strSQL & "    SELECT u.userid, MIN(sc.scate_name) AS region_name "
strSQL = strSQL & "    FROM vUser_info AS u WITH (READUNCOMMITTED) "
strSQL = strSQL & "    LEFT JOIN school_info AS si WITH (READUNCOMMITTED) ON u.school_num = si.school_num "
strSQL = strSQL & "    LEFT JOIN school_category AS sc WITH (READUNCOMMITTED) ON si.sido_position = sc.scate_num "
strSQL = strSQL & "    GROUP BY u.userid "
strSQL = strSQL & ") "
strSQL = strSQL & "SELECT "
strSQL = strSQL & "    LEFT(p.pay_date, 4) AS pay_year, "
strSQL = strSQL & "    ISNULL(ur.region_name, '미분류') AS region_name, "
strSQL = strSQL & "    COUNT(DISTINCT p.userid) AS buyer_count, "
strSQL = strSQL & "    SUM(ISNULL(p.pay_price, 0)) AS actual_amount, "
strSQL = strSQL & "    COUNT(DISTINCT CASE "
strSQL = strSQL & "        WHEN sms.userid IS NOT NULL THEN p.userid "
strSQL = strSQL & "    END) AS sms_count "

strSQL = strSQL & "FROM pay_info AS p WITH (READUNCOMMITTED) "

strSQL = strSQL & "LEFT JOIN UserRegion AS ur "
strSQL = strSQL & "    ON p.userid = ur.userid "

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
strSQL = strSQL & "    ISNULL(ur.region_name, '미분류') "

strSQL = strSQL & "ORDER BY "
strSQL = strSQL & "    LEFT(p.pay_date, 4) ASC, "
strSQL = strSQL & "    COUNT(DISTINCT p.userid) DESC, "
strSQL = strSQL & "    ISNULL(ur.region_name, '미분류') ASC "

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
    strSQL = strSQL & fnCdbValue("stat_regional_analysis_" & downloadDate) & ","
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
    "attachment; filename=stat_regional_analysis_" & _
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
            <th width="100">연도</th>
            <th width="150">지역</th>
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

If IsArray(arrResult) Then

    For i = 0 To UBound(arrResult, 2)

        payYear = Nz(arrResult(0, i), "")
        regionName = Nz(arrResult(1, i), "미분류")

        buyerCount = CLng(Nz(arrResult(2, i), 0))
        actualAmount = CLng(Nz(arrResult(3, i), 0))
        smsCount = CLng(Nz(arrResult(4, i), 0))

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

            <td align="left">
                <%=Server.HTMLEncode(regionName & "")%>
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
            <td colspan="2" align="center">
                연도별 합계
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