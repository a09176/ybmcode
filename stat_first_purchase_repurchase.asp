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
Dim j

Dim startDate
Dim endDate
Dim downloadDate

Dim firstYear
Dim firstMonth
Dim firstMonthName

Dim firstBuyerCount
Dim repurchaseBuyerCount
Dim repurchaseRate
Dim repurchasePaymentCount
Dim repurchaseAmount

Dim currentYear
Dim previousYear

Dim totalFirstBuyerCount
Dim totalRepurchaseBuyerCount
Dim totalRepurchasePaymentCount
Dim totalRepurchaseAmount

Dim totalRepurchaseRate

Dim hasResult
Dim rowCount

Dim resultYear
Dim resultMonth

'==========================================================
' 분석 기간
'
' 첫 구매 여부는 전체 결제 이력을 기준으로 확인합니다.
' 따라서 2025년 이전 결제 이력도 조회 대상에 포함됩니다.
'
' 결과 출력 대상은 다음 기간입니다.
' 2025-01-01 ~ 2026-12-31
'==========================================================
startDate = "2025-01-01"
endDate = "2027-01-01"

downloadDate = Year(Now()) & _
               Right("0" & Month(Now()), 2) & _
               Right("0" & Day(Now()), 2)

'==========================================================
' 첫 구매 고객 및 재구매 현황 조회
'
' 기준
' 1. sell_info = 'o'인 결제완료 건만 분석
' 2. 회원별 최초 결제 건을 첫 구매로 판단
' 3. 최초 결제가 2025년 또는 2026년에 발생한 회원만 출력
' 4. 최초 결제 이후 다시 결제한 회원을 재구매 고객으로 판단
' 5. 첫 구매월별로 고객 수와 재구매 현황 집계
'
' 재구매 판정
' - pay_date가 최초 결제일보다 이후인 경우
' - 결제일이 같으면 pay_num이 더 큰 경우
'
' pay_num을 함께 비교하는 이유
' - 동일한 시간에 발생한 결제 건도 순서를 구분하기 위함
'==========================================================
strSQL = ""

strSQL = strSQL & "WITH CompletedPayments AS ( "
strSQL = strSQL & "    SELECT "
strSQL = strSQL & "        p.pay_num, "
strSQL = strSQL & "        p.userid, "
strSQL = strSQL & "        p.pay_date, "
strSQL = strSQL & "        ISNULL(p.pay_price, 0) AS pay_price "
strSQL = strSQL & "    FROM pay_info AS p WITH (READUNCOMMITTED) "
strSQL = strSQL & "    WHERE p.sell_info = 'o' "
strSQL = strSQL & "      AND ISNULL(p.userid, '') <> '' "
strSQL = strSQL & "      AND p.pay_date < '" & endDate & "' "
strSQL = strSQL & "), "

strSQL = strSQL & "NumberedPayments AS ( "
strSQL = strSQL & "    SELECT "
strSQL = strSQL & "        cp.pay_num, "
strSQL = strSQL & "        cp.userid, "
strSQL = strSQL & "        cp.pay_date, "
strSQL = strSQL & "        cp.pay_price, "
strSQL = strSQL & "        ROW_NUMBER() OVER ( "
strSQL = strSQL & "            PARTITION BY cp.userid "
strSQL = strSQL & "            ORDER BY cp.pay_date, cp.pay_num "
strSQL = strSQL & "        ) AS payment_order "
strSQL = strSQL & "    FROM CompletedPayments AS cp "
strSQL = strSQL & "), "

strSQL = strSQL & "FirstPayments AS ( "
strSQL = strSQL & "    SELECT "
strSQL = strSQL & "        np.userid, "
strSQL = strSQL & "        np.pay_num AS first_pay_num, "
strSQL = strSQL & "        np.pay_date AS first_pay_date, "
strSQL = strSQL & "        np.pay_price AS first_pay_price "
strSQL = strSQL & "    FROM NumberedPayments AS np "
strSQL = strSQL & "    WHERE np.payment_order = 1 "
strSQL = strSQL & "), "

strSQL = strSQL & "UserRepurchase AS ( "
strSQL = strSQL & "    SELECT "
strSQL = strSQL & "        fp.userid, "
strSQL = strSQL & "        LEFT(fp.first_pay_date, 4) AS first_year, "
strSQL = strSQL & "        SUBSTRING(fp.first_pay_date, 6, 2) AS first_month, "
strSQL = strSQL & "        CASE "
strSQL = strSQL & "            WHEN COUNT(cp.pay_num) > 0 THEN 1 "
strSQL = strSQL & "            ELSE 0 "
strSQL = strSQL & "        END AS is_repurchase, "
strSQL = strSQL & "        COUNT(cp.pay_num) AS repurchase_payment_count, "
strSQL = strSQL & "        SUM(ISNULL(cp.pay_price, 0)) AS repurchase_amount "
strSQL = strSQL & "    FROM FirstPayments AS fp "
strSQL = strSQL & "    LEFT JOIN CompletedPayments AS cp "
strSQL = strSQL & "        ON cp.userid = fp.userid "
strSQL = strSQL & "       AND ( "
strSQL = strSQL & "            cp.pay_date > fp.first_pay_date "
strSQL = strSQL & "            OR ( "
strSQL = strSQL & "                cp.pay_date = fp.first_pay_date "
strSQL = strSQL & "                AND cp.pay_num > fp.first_pay_num "
strSQL = strSQL & "            ) "
strSQL = strSQL & "       ) "
strSQL = strSQL & "    WHERE fp.first_pay_date >= '" & startDate & "' "
strSQL = strSQL & "      AND fp.first_pay_date < '" & endDate & "' "
strSQL = strSQL & "    GROUP BY "
strSQL = strSQL & "        fp.userid, "
strSQL = strSQL & "        fp.first_pay_num, "
strSQL = strSQL & "        fp.first_pay_date "
strSQL = strSQL & ") "

strSQL = strSQL & "SELECT "
strSQL = strSQL & "    first_year, "
strSQL = strSQL & "    first_month, "
strSQL = strSQL & "    COUNT(*) AS first_buyer_count, "
strSQL = strSQL & "    SUM(is_repurchase) AS repurchase_buyer_count, "
strSQL = strSQL & "    SUM(repurchase_payment_count) AS repurchase_payment_count, "
strSQL = strSQL & "    SUM(ISNULL(repurchase_amount, 0)) AS repurchase_amount "
strSQL = strSQL & "FROM UserRepurchase "
strSQL = strSQL & "GROUP BY "
strSQL = strSQL & "    first_year, "
strSQL = strSQL & "    first_month "
strSQL = strSQL & "ORDER BY "
strSQL = strSQL & "    first_year, "
strSQL = strSQL & "    first_month "

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
    strSQL = strSQL & fnCdbValue("stat_first_purchase_repurchase_" & downloadDate) & ","
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
    "attachment; filename=stat_first_purchase_repurchase_" & _
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
    2025년과 2026년에 처음 구매한 고객의 재구매 현황을 월별로 확인할 수 있는 자료입니다.<br>
    회원별 전체 결제 이력을 기준으로 최초 결제 시점을 확인하고, 첫 구매 이후 다시 결제한 고객을 재구매 고객으로 집계합니다.<br>
    첫 구매 고객 수, 재구매 고객 수, 재구매율, 재구매 결제 건수와 재구매 매출액을 확인할 수 있습니다.<br>
    재구매율은 해당 월의 첫 구매 고객 중 이후 다시 구매한 고객의 비율입니다.<br>
    2025년 이전에 이미 결제한 회원은 2025년 또는 2026년의 첫 구매 고객에 포함되지 않습니다.<br>
    결제완료된 데이터만 분석 대상에 포함됩니다.
</p>

<hr style="margin: 20px 0; border: 1px solid #999;">

<%
hasResult = False
rowCount = 0

If IsArray(arrResult) Then

    On Error Resume Next
    rowCount = UBound(arrResult, 2)

    If Err.Number <> 0 Then
        rowCount = -1
        Err.Clear
    End If

    On Error GoTo 0

    If rowCount >= 0 Then
        hasResult = True
    End If

End If

If Not hasResult Then
%>

<table border="1">
    <tr>
        <td align="center">
            조회 결과가 없습니다.
        </td>
    </tr>
</table>

<%
Else

    currentYear = ""

    totalFirstBuyerCount = 0
    totalRepurchaseBuyerCount = 0
    totalRepurchasePaymentCount = 0
    totalRepurchaseAmount = 0

    For i = 0 To rowCount

        resultYear = Nz(arrResult(0, i), "")
        resultMonth = Nz(arrResult(1, i), "")

        firstBuyerCount = CLng(Nz(arrResult(2, i), 0))
        repurchaseBuyerCount = CLng(Nz(arrResult(3, i), 0))
        repurchasePaymentCount = CLng(Nz(arrResult(4, i), 0))
        repurchaseAmount = CDbl(Nz(arrResult(5, i), 0))

        If firstBuyerCount > 0 Then
            repurchaseRate = (repurchaseBuyerCount / firstBuyerCount) * 100
        Else
            repurchaseRate = 0
        End If

        '--------------------------------------------------
        ' 연도가 변경되면 이전 연도 표 종료
        '--------------------------------------------------
        If currentYear <> resultYear Then

            If currentYear <> "" Then
%>
                    <tr style="background-color: #E7E6E6; font-weight: bold;">
                        <td colspan="2" align="center">
                            <%=Server.HTMLEncode(currentYear & "년 합계")%>
                        </td>

                        <td align="right"
                            style="mso-number-format:'\#\,\#\#0';">
                            <%=FormatNumber(totalFirstBuyerCount, 0)%>
                        </td>

                        <td align="right"
                            style="mso-number-format:'\#\,\#\#0';">
                            <%=FormatNumber(totalRepurchaseBuyerCount, 0)%>
                        </td>

                        <td align="right"
                            style="mso-number-format:'0.00%';">
                            <%
                            If totalFirstBuyerCount > 0 Then
                                totalRepurchaseRate = _
                                    (totalRepurchaseBuyerCount / totalFirstBuyerCount) * 100
                                Response.Write FormatNumber(totalRepurchaseRate / 100, 4)
                            Else
                                Response.Write "0.00%"
                            End If
                            %>
                        </td>

                        <td align="right"
                            style="mso-number-format:'\#\,\#\#0';">
                            <%=FormatNumber(totalRepurchasePaymentCount, 0)%>
                        </td>

                        <td align="right"
                            style="mso-number-format:'\#\,\#\#0';">
                            <%=FormatNumber(totalRepurchaseAmount, 0)%>
                        </td>
                    </tr>
                </tbody>
            </table>

            <br>
<%
            End If

            currentYear = resultYear

            totalFirstBuyerCount = 0
            totalRepurchaseBuyerCount = 0
            totalRepurchasePaymentCount = 0
            totalRepurchaseAmount = 0
%>

            <!--==================================================
                연도별 첫 구매 고객 재구매 분석
            ===================================================-->
            <table border="1" style="margin-bottom: 20px;">
                <thead>
                    <tr>
                        <td colspan="7"
                            style="font-weight: bold; font-size: 14px; background-color: #D3D3D3; padding: 8px;">
                            <%=Server.HTMLEncode(resultYear & "년 첫 구매 고객 재구매 분석")%>
                        </td>
                    </tr>

                    <tr style="background-color: #4472C4; color: white; font-weight: bold;">
                        <th width="100">첫 구매월</th>
                        <th width="130">첫 구매 고객수</th>
                        <th width="130">재구매 고객수</th>
                        <th width="120">재구매율</th>
                        <th width="150">재구매 결제건수</th>
                        <th width="150">재구매 매출액</th>
                        <th width="250">분석 기준</th>
                    </tr>
                </thead>

                <tbody>
<%
        End If

        Select Case resultMonth
            Case "01"
                firstMonthName = "1월"
            Case "02"
                firstMonthName = "2월"
            Case "03"
                firstMonthName = "3월"
            Case "04"
                firstMonthName = "4월"
            Case "05"
                firstMonthName = "5월"
            Case "06"
                firstMonthName = "6월"
            Case "07"
                firstMonthName = "7월"
            Case "08"
                firstMonthName = "8월"
            Case "09"
                firstMonthName = "9월"
            Case "10"
                firstMonthName = "10월"
            Case "11"
                firstMonthName = "11월"
            Case "12"
                firstMonthName = "12월"
            Case Else
                firstMonthName = resultMonth & "월"
        End Select

        totalFirstBuyerCount = totalFirstBuyerCount + firstBuyerCount
        totalRepurchaseBuyerCount = totalRepurchaseBuyerCount + repurchaseBuyerCount
        totalRepurchasePaymentCount = _
            totalRepurchasePaymentCount + repurchasePaymentCount
        totalRepurchaseAmount = totalRepurchaseAmount + repurchaseAmount
%>

                    <tr>
                        <td align="center">
                            <%=Server.HTMLEncode(firstMonthName)%>
                        </td>

                        <td align="right"
                            style="mso-number-format:'\#\,\#\#0';">
                            <%=FormatNumber(firstBuyerCount, 0)%>
                        </td>

                        <td align="right"
                            style="mso-number-format:'\#\,\#\#0';">
                            <%=FormatNumber(repurchaseBuyerCount, 0)%>
                        </td>

                        <td align="right"
                            style="mso-number-format:'0.00%';">
                            <%=FormatNumber(repurchaseRate / 100, 4)%>
                        </td>

                        <td align="right"
                            style="mso-number-format:'\#\,\#\#0';">
                            <%=FormatNumber(repurchasePaymentCount, 0)%>
                        </td>

                        <td align="right"
                            style="mso-number-format:'\#\,\#\#0';">
                            <%=FormatNumber(repurchaseAmount, 0)%>
                        </td>

                        <td align="left">
                            첫 구매 이후 재결제 기준
                        </td>
                    </tr>

<%
        '------------------------------------------------------
        ' 현재 연도의 마지막 월이면 합계 출력
        '------------------------------------------------------
        If i = rowCount Then
%>
                    <tr style="background-color: #E7E6E6; font-weight: bold;">
                        <td align="center">
                            <%=Server.HTMLEncode(currentYear & "년 합계")%>
                        </td>

                        <td align="right"
                            style="mso-number-format:'\#\,\#\#0';">
                            <%=FormatNumber(totalFirstBuyerCount, 0)%>
                        </td>

                        <td align="right"
                            style="mso-number-format:'\#\,\#\#0';">
                            <%=FormatNumber(totalRepurchaseBuyerCount, 0)%>
                        </td>

                        <td align="right"
                            style="mso-number-format:'0.00%';">
                            <%
                            If totalFirstBuyerCount > 0 Then
                                totalRepurchaseRate = _
                                    (totalRepurchaseBuyerCount / totalFirstBuyerCount) * 100
                                Response.Write FormatNumber(totalRepurchaseRate / 100, 4)
                            Else
                                Response.Write "0.00%"
                            End If
                            %>
                        </td>

                        <td align="right"
                            style="mso-number-format:'\#\,\#\#0';">
                            <%=FormatNumber(totalRepurchasePaymentCount, 0)%>
                        </td>

                        <td align="right"
                            style="mso-number-format:'\#\,\#\#0';">
                            <%=FormatNumber(totalRepurchaseAmount, 0)%>
                        </td>

                        <td align="left">
                            해당 연도 첫 구매 고객 합계
                        </td>
                    </tr>

                </tbody>
            </table>
<%
        Else

            previousYear = Nz(arrResult(0, i + 1), "")

            If previousYear <> currentYear Then
%>
                    <tr style="background-color: #E7E6E6; font-weight: bold;">
                        <td align="center">
                            <%=Server.HTMLEncode(currentYear & "년 합계")%>
                        </td>

                        <td align="right"
                            style="mso-number-format:'\#\,\#\#0';">
                            <%=FormatNumber(totalFirstBuyerCount, 0)%>
                        </td>

                        <td align="right"
                            style="mso-number-format:'\#\,\#\#0';">
                            <%=FormatNumber(totalRepurchaseBuyerCount, 0)%>
                        </td>

                        <td align="right"
                            style="mso-number-format:'0.00%';">
                            <%
                            If totalFirstBuyerCount > 0 Then
                                totalRepurchaseRate = _
                                    (totalRepurchaseBuyerCount / totalFirstBuyerCount) * 100
                                Response.Write FormatNumber(totalRepurchaseRate / 100, 4)
                            Else
                                Response.Write "0.00%"
                            End If
                            %>
                        </td>

                        <td align="right"
                            style="mso-number-format:'\#\,\#\#0';">
                            <%=FormatNumber(totalRepurchasePaymentCount, 0)%>
                        </td>

                        <td align="right"
                            style="mso-number-format:'\#\,\#\#0';">
                            <%=FormatNumber(totalRepurchaseAmount, 0)%>
                        </td>

                        <td align="left">
                            해당 연도 첫 구매 고객 합계
                        </td>
                    </tr>

                </tbody>
            </table>

            <br>
<%
            End If

        End If

    Next

End If
%>

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