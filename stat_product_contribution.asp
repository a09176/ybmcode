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

Dim saleYear
Dim productNum
Dim tcourseNum
Dim productName

Dim saleCount
Dim productAmount
Dim yearTotalAmount

Dim contributionRate
Dim cumulativeRate

Dim currentYear
Dim previousYear

Dim totalSaleCount
Dim totalProductAmount

Dim hasResult
Dim rowCount

Dim yearText
Dim safeProductName

'==========================================================
' 조회 기간
'
' 2025년과 2026년을 분리해서 분석
' 2026년은 미래 데이터가 없으므로 실제 판매 데이터가 있는
' 날짜까지만 결과에 반영됨
'==========================================================
startDate = "2025-01-01"
endDate = "2027-01-01"

downloadDate = Year(Now()) & _
               Right("0" & Month(Now()), 2) & _
               Right("0" & Day(Now()), 2)

'==========================================================
' 과정별 상품금액 기여도 조회
'
' 집계 기준
' - 과정 식별 기준: product_num
' - 과정 상품만 포함: product_type = 'S'
' - 판매수: pay_product_num 고유 개수
' - 상품금액: pay_info_product.product_price 합계 (주문 실결제액과 별도 지표)
' - 매출 비중: 해당 연도 전체 과정 매출 대비 비율
'
' 주의
' - 같은 과정명이더라도 product_num이 다르면 별도 과정으로 표시
' - 할인상품, 패키지상품 등은 product_num 기준으로 별도 집계
'==========================================================
strSQL = ""

strSQL = strSQL & "WITH ProductSales AS ( "
strSQL = strSQL & "    SELECT "
strSQL = strSQL & "        LEFT(p.pay_date, 4) AS sale_year, "
strSQL = strSQL & "        pi.product_num, "
strSQL = strSQL & "        pi.Tcourse_num, "
strSQL = strSQL & "        ISNULL(pi.product_name, '과정명 없음') AS product_name, "
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
strSQL = strSQL & "      AND p.pay_date < '" & endDate & "' "
strSQL = strSQL & "    GROUP BY "
strSQL = strSQL & "        LEFT(p.pay_date, 4), "
strSQL = strSQL & "        pi.product_num, "
strSQL = strSQL & "        pi.Tcourse_num, "
strSQL = strSQL & "        pi.product_name "
strSQL = strSQL & ") "
strSQL = strSQL & "SELECT "
strSQL = strSQL & "    sale_year, "
strSQL = strSQL & "    product_num, "
strSQL = strSQL & "    Tcourse_num, "
strSQL = strSQL & "    product_name, "
strSQL = strSQL & "    sale_count, "
strSQL = strSQL & "    product_amount, "
strSQL = strSQL & "    SUM(product_amount) OVER (PARTITION BY sale_year) AS year_total_amount "
strSQL = strSQL & "FROM ProductSales "
strSQL = strSQL & "ORDER BY "
strSQL = strSQL & "    sale_year, "
strSQL = strSQL & "    product_amount DESC, "
strSQL = strSQL & "    sale_count DESC, "
strSQL = strSQL & "    product_num "

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
    strSQL = strSQL & fnCdbValue("stat_product_contribution_" & downloadDate) & ","
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
    "attachment; filename=stat_product_contribution_" & _
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
    2025년과 2026년의 과정별 판매 현황과 매출 기여도를 비교할 수 있는 자료입니다.<br>
    과정별 판매수와 상품금액(product_price)을 기준으로 해당 연도 전체 과정 상품금액에서 차지하는 비중을 확인할 수 있습니다.<br>
    상품금액이 높은 과정부터 순서대로 표시하며, 주문 실결제액(pay_info.pay_price)과는 별도 지표입니다.<br>
    같은 과정명이더라도 상품번호가 다른 경우에는 별도의 상품으로 집계됩니다.<br>
    과정 상품만 분석 대상에 포함하며, 결제완료된 판매 데이터만 집계합니다.
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
    cumulativeRate = 0
    totalSaleCount = 0
    totalProductAmount = 0

    For i = 0 To rowCount

        saleYear = Nz(arrResult(0, i), "")
        productNum = CLng(Nz(arrResult(1, i), 0))
        tcourseNum = CLng(Nz(arrResult(2, i), 0))
        productName = Nz(arrResult(3, i), "과정명 없음")
        saleCount = CLng(Nz(arrResult(4, i), 0))
        productAmount = CDbl(Nz(arrResult(5, i), 0))
        yearTotalAmount = CDbl(Nz(arrResult(6, i), 0))

        '--------------------------------------------------
        ' 연도가 변경되면 이전 연도 표 종료
        '--------------------------------------------------
        If currentYear <> saleYear Then

            If currentYear <> "" Then
%>
                <tr style="background-color: #E7E6E6; font-weight: bold;">
                    <td colspan="3" align="center">
                        <%=Server.HTMLEncode(currentYear & "년 합계")%>
                    </td>

                    <td align="right"
                        style="mso-number-format:'\#\,\#\#0';">
                        <%=FormatNumber(totalSaleCount, 0)%>
                    </td>

                    <td align="right"
                        style="mso-number-format:'\#\,\#\#0';">
                        <%=FormatNumber(totalProductAmount, 0)%>
                    </td>

                    <td align="right"
                        style="mso-number-format:'0.00%';">
                        100.00%
                    </td>

                    <td align="right"
                        style="mso-number-format:'0.00%';">
                        100.00%
                    </td>
                </tr>

                </tbody>
                </table>

                <br>
<%
            End If

            currentYear = saleYear
            cumulativeRate = 0
            totalSaleCount = 0
            totalProductAmount = 0
%>

            <!--==================================================
                <%=Server.HTMLEncode(saleYear & "년")%> 과정별 상품금액 기여도
            ===================================================-->
            <table border="1" style="margin-bottom: 20px;">
                <thead>
                    <tr>
                        <td colspan="7"
                            style="font-weight: bold; font-size: 14px; background-color: #D3D3D3; padding: 8px;">
                            <%=Server.HTMLEncode(saleYear & "년 과정별 상품금액 기여도")%>
                        </td>
                    </tr>

                    <tr style="background-color: #4472C4; color: white; font-weight: bold;">
                        <th width="70">순위</th>
                        <th width="100">상품번호</th>
                        <th width="100">과정번호</th>
                        <th width="350">과정명</th>
                        <th width="100">판매수</th>
                        <th width="150">상품금액(product_price)</th>
                        <th width="120">상품금액 비중</th>
                        <th width="120">누적 비중</th>
                    </tr>
                </thead>

                <tbody>
<%
        End If

        If yearTotalAmount > 0 Then
            contributionRate = (productAmount / yearTotalAmount) * 100
        Else
            contributionRate = 0
        End If

        cumulativeRate = cumulativeRate + contributionRate

        totalSaleCount = totalSaleCount + saleCount
        totalProductAmount = totalProductAmount + productAmount
%>

                    <tr>
                        <td align="center">
                            <%=i + 1%>
                        </td>

                        <td align="center"
                            style="mso-number-format:'\@';">
                            <%=productNum%>
                        </td>

                        <td align="center"
                            style="mso-number-format:'\@';">
                            <%=tcourseNum%>
                        </td>

                        <td align="left">
                            <%=Server.HTMLEncode(productName & "")%>
                        </td>

                        <td align="right"
                            style="mso-number-format:'\#\,\#\#0';">
                            <%=FormatNumber(saleCount, 0)%>
                        </td>

                        <td align="right"
                            style="mso-number-format:'\#\,\#\#0';">
                            <%=FormatNumber(productAmount, 0)%>
                        </td>

                        <td align="right"
                            style="mso-number-format:'0.00%';">
                            <%=FormatNumber(contributionRate / 100, 4)%>
                        </td>

                        <td align="right"
                            style="mso-number-format:'0.00%';">
                            <%=FormatNumber(cumulativeRate / 100, 4)%>
                        </td>
                    </tr>

<%
        '------------------------------------------------------
        ' 다음 행이 마지막이면 현재 연도의 합계 출력
        '------------------------------------------------------
        If i = rowCount Then
%>
                    <tr style="background-color: #E7E6E6; font-weight: bold;">
                        <td colspan="3" align="center">
                            <%=Server.HTMLEncode(currentYear & "년 합계")%>
                        </td>

                        <td align="center">
                            전체 과정
                        </td>

                        <td align="right"
                            style="mso-number-format:'\#\,\#\#0';">
                            <%=FormatNumber(totalSaleCount, 0)%>
                        </td>

                        <td align="right"
                            style="mso-number-format:'\#\,\#\#0';">
                            <%=FormatNumber(totalProductAmount, 0)%>
                        </td>

                        <td align="right"
                            style="mso-number-format:'0.00%';">
                            100.00%
                        </td>

                        <td align="right"
                            style="mso-number-format:'0.00%';">
                            100.00%
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
                        <td colspan="3" align="center">
                            <%=Server.HTMLEncode(currentYear & "년 합계")%>
                        </td>

                        <td align="center">
                            전체 과정
                        </td>

                        <td align="right"
                            style="mso-number-format:'\#\,\#\#0';">
                            <%=FormatNumber(totalSaleCount, 0)%>
                        </td>

                        <td align="right"
                            style="mso-number-format:'\#\,\#\#0';">
                            <%=FormatNumber(totalProductAmount, 0)%>
                        </td>

                        <td align="right"
                            style="mso-number-format:'0.00%';">
                            100.00%
                        </td>

                        <td align="right"
                            style="mso-number-format:'0.00%';">
                            100.00%
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