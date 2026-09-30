<%@ LANGUAGE="VBScript" %>
<%
Option Explicit
Response.Expires = -1

Const iMenuIdex = 611
Const current_menu_num = 75
%>

<!--#include virtual="/teacher_comm/session_keep.asp" -->
<!--#include virtual="/teacher_inc/definition.asp" -->
<!--#include virtual="/teacher_inc/function.asp" -->
<!--#include virtual="/lms/inc/admin_config.asp" -->
<!--#include virtual="/lms/inc/chkAdmin.asp" -->
<!--#include virtual="/v15/include/fnDB.asp" -->
<!--#include virtual="/common/lib/FNC_NonInjection.asp" -->

<%
Call AdminLoginCheck()
Call fnAdminPermission(Session("admin_group"), current_menu_num, False)

'==========================================================
' 변수 선언
'==========================================================
Dim strSQL
Dim arrOverview
Dim arrCoupon
Dim arrSms
Dim arrTopProduct
Dim arrTopRegion
Dim arrTopAge
Dim arrTopSubject

Dim i

Dim startDay
Dim endDay
Dim startDate
Dim endDate
Dim downloadDate

Dim totalUser
Dim totalPayCount
Dim totalActual
Dim couponUser
Dim smsUser
Dim avgPayAmount

Dim itemName
Dim itemCount
Dim itemAmount

Dim hasOverview
Dim hasCoupon
Dim hasSms
Dim hasTopProduct
Dim hasTopRegion
Dim hasTopAge
Dim hasTopSubject

'==========================================================
' 조회 기간
'==========================================================
startDay = ReadReportDate( _
    "sYear", _
    "sMonth", _
    "sDay", _
    DateSerial(2025, 1, 1) _
)

endDay = ReadReportDate( _
    "eYear", _
    "eMonth", _
    "eDay", _
    Date() _
)

If startDay > endDay Then
    Response.Status = "400 Bad Request"
    Response.Write "조회 시작일은 종료일보다 늦을 수 없습니다."
    Response.End
End If

startDate = SqlDateText(startDay)
endDate = SqlDateText(endDay)

downloadDate = Year(Date()) & _
               Right("0" & Month(Date()), 2) & _
               Right("0" & Day(Date()), 2)

'==========================================================
' 초기화
'==========================================================
totalUser = 0
totalPayCount = 0
totalActual = 0
couponUser = 0
smsUser = 0
avgPayAmount = 0

hasOverview = False
hasCoupon = False
hasSms = False
hasTopProduct = False
hasTopRegion = False
hasTopAge = False
hasTopSubject = False

'==========================================================
' 1. 기본 KPI
'
' pay_info만 집계하여 상품·쿠폰·SMS 조인으로 인한
' 결제액 중복을 방지한다.
'==========================================================
strSQL = ""
strSQL = strSQL & "SELECT "
strSQL = strSQL & "    COUNT(DISTINCT p.userid) AS buyer_count, "
strSQL = strSQL & "    COUNT(DISTINCT p.pay_num) AS pay_count, "
strSQL = strSQL & "    SUM(ISNULL(p.pay_price, 0)) AS paid_amount "
strSQL = strSQL & "FROM pay_info AS p WITH (READUNCOMMITTED) "
strSQL = strSQL & "WHERE p.sell_info = 'o' "
strSQL = strSQL & "  AND p.pay_date >= '" & startDate & "' "
strSQL = strSQL & "  AND p.pay_date < DATEADD(DAY, 1, '" & endDate & "') "
strSQL = strSQL & "  AND p.userid IS NOT NULL "
strSQL = strSQL & "  AND LTRIM(RTRIM(p.userid)) <> '' "

arrOverview = ExecSP(strSQL, strCyberTeacher)

If IsArray(arrOverview) Then
    hasOverview = True

    totalUser = ToNumber(arrOverview(0, 0))
    totalPayCount = ToNumber(arrOverview(1, 0))
    totalActual = ToNumber(arrOverview(2, 0))
End If

If totalUser > 0 Then
    avgPayAmount = totalActual / totalUser
End If

'==========================================================
' 2. 쿠폰 사용 구매자
'
' 기간 내 쿠폰이 적용된 결제가 한 건이라도 있는
' 고유 userid 수를 집계한다.
'==========================================================
strSQL = ""
strSQL = strSQL & "SELECT COUNT(DISTINCT p.userid) AS coupon_user "
strSQL = strSQL & "FROM pay_info AS p WITH (READUNCOMMITTED) "
strSQL = strSQL & "WHERE p.sell_info = 'o' "
strSQL = strSQL & "  AND p.pay_date >= '" & startDate & "' "
strSQL = strSQL & "  AND p.pay_date < DATEADD(DAY, 1, '" & endDate & "') "
strSQL = strSQL & "  AND p.userid IS NOT NULL "
strSQL = strSQL & "  AND LTRIM(RTRIM(p.userid)) <> '' "
strSQL = strSQL & "  AND EXISTS ( "
strSQL = strSQL & "      SELECT 1 "
strSQL = strSQL & "      FROM pay_info_product AS pip WITH (READUNCOMMITTED) "
strSQL = strSQL & "      INNER JOIN pay_coupon_use AS cu WITH (READUNCOMMITTED) "
strSQL = strSQL & "          ON pip.pay_product_num = cu.pay_product_num "
strSQL = strSQL & "      WHERE pip.pay_num = p.pay_num "
strSQL = strSQL & "  ) "

arrCoupon = ExecSP(strSQL, strCyberTeacher)

If IsArray(arrCoupon) Then
    hasCoupon = True
    couponUser = ToNumber(arrCoupon(0, 0))
End If

'==========================================================
' 3. SMS 수신동의자
'
' 기간 내 결제 회원 중 현재 SMS 수신동의자가 있는
' 고유 userid 수를 집계한다.
'==========================================================
strSQL = ""
strSQL = strSQL & "SELECT COUNT(DISTINCT p.userid) AS sms_user "
strSQL = strSQL & "FROM pay_info AS p WITH (READUNCOMMITTED) "
strSQL = strSQL & "WHERE p.sell_info = 'o' "
strSQL = strSQL & "  AND p.pay_date >= '" & startDate & "' "
strSQL = strSQL & "  AND p.pay_date < DATEADD(DAY, 1, '" & endDate & "') "
strSQL = strSQL & "  AND p.userid IS NOT NULL "
strSQL = strSQL & "  AND LTRIM(RTRIM(p.userid)) <> '' "
strSQL = strSQL & "  AND EXISTS ( "
strSQL = strSQL & "      SELECT 1 "
strSQL = strSQL & "      FROM info_receive_agree AS ira WITH (READUNCOMMITTED) "
strSQL = strSQL & "      WHERE ira.userid = p.userid "
strSQL = strSQL & "        AND ira.agree_type = 'SMS' "
strSQL = strSQL & "  ) "

arrSms = ExecSP(strSQL, strCyberTeacher)

If IsArray(arrSms) Then
    hasSms = True
    smsUser = ToNumber(arrSms(0, 0))
End If

'==========================================================
' 4. 상위 판매 과정 TOP 5
'
' 판매수 기준으로 정렬하고 판매수와 결제액을 모두 출력한다.
'==========================================================
strSQL = ""
strSQL = strSQL & "SELECT TOP 5 "
strSQL = strSQL & "    ISNULL(pi.product_name, '과정명 없음') AS product_name, "
strSQL = strSQL & "    COUNT(DISTINCT pip.pay_product_num) AS sale_count, "
strSQL = strSQL & "    SUM(ISNULL(pip.product_price, 0)) AS actual_amount "
strSQL = strSQL & "FROM pay_info AS p WITH (READUNCOMMITTED) "
strSQL = strSQL & "INNER JOIN pay_info_product AS pip WITH (READUNCOMMITTED) "
strSQL = strSQL & "    ON p.pay_num = pip.pay_num "
strSQL = strSQL & "INNER JOIN product_info AS pi WITH (READUNCOMMITTED) "
strSQL = strSQL & "    ON pip.product_num = pi.product_num "
strSQL = strSQL & "WHERE p.sell_info = 'o' "
strSQL = strSQL & "  AND pi.product_type = 'S' "
strSQL = strSQL & "  AND p.pay_date >= '" & startDate & "' "
strSQL = strSQL & "  AND p.pay_date < DATEADD(DAY, 1, '" & endDate & "') "
strSQL = strSQL & "GROUP BY pi.product_name "
strSQL = strSQL & "ORDER BY "
strSQL = strSQL & "    COUNT(DISTINCT pip.pay_product_num) DESC, "
strSQL = strSQL & "    SUM(ISNULL(pip.product_price, 0)) DESC "

arrTopProduct = ExecSP(strSQL, strCyberTeacher)

If IsArray(arrTopProduct) Then
    hasTopProduct = True
End If

'==========================================================
' 5. 지역별 상위 판매 TOP 5
'
' 지역 연결 기준 (stat_regional_analysis.asp와 동일):
'   pay_info.userid
'     -> vUser_info.userid
'     -> school_info.school_num
'     -> school_info.sido_position
'     -> school_category.scate_num
'==========================================================
strSQL = ""
strSQL = strSQL & "SELECT TOP 5 "
strSQL = strSQL & "    ISNULL(sc.scate_name, '미분류') AS region_name, "
strSQL = strSQL & "    COUNT(DISTINCT p.userid) AS user_count, "
strSQL = strSQL & "    SUM(ISNULL(p.pay_price, 0)) AS actual_amount "
strSQL = strSQL & "FROM pay_info AS p WITH (READUNCOMMITTED) "
strSQL = strSQL & "LEFT JOIN vUser_info AS u WITH (READUNCOMMITTED) "
strSQL = strSQL & "    ON p.userid = u.userid "
strSQL = strSQL & "LEFT JOIN school_info AS si WITH (READUNCOMMITTED) "
strSQL = strSQL & "    ON u.school_num = si.school_num "
strSQL = strSQL & "LEFT JOIN school_category AS sc WITH (READUNCOMMITTED) "
strSQL = strSQL & "    ON si.sido_position = sc.scate_num "
strSQL = strSQL & "WHERE p.sell_info = 'o' "
strSQL = strSQL & "  AND p.pay_date >= '" & startDate & "' "
strSQL = strSQL & "  AND p.pay_date < DATEADD(DAY, 1, '" & endDate & "') "
strSQL = strSQL & "  AND p.userid IS NOT NULL "
strSQL = strSQL & "  AND LTRIM(RTRIM(p.userid)) <> '' "
strSQL = strSQL & "GROUP BY sc.scate_name "
strSQL = strSQL & "ORDER BY "
strSQL = strSQL & "    SUM(ISNULL(p.pay_price, 0)) DESC "

arrTopRegion = ExecSP(strSQL, strCyberTeacher)

If IsArray(arrTopRegion) Then
    hasTopRegion = True
End If

'==========================================================
' 6. 연령대별 상위 판매 TOP 5
'
' FamilyM 함수 조회를 위해 usp_openKey를 같은 SQL 실행에
' 포함한다.
'
' fm.b_year는 출생연도 숫자이므로 YEAR(fm.b_year)를
' 사용하지 않는다.
'==========================================================
strSQL = ""
strSQL = strSQL & "EXEC dbo.usp_openKey; "

strSQL = strSQL & "WITH AgeData AS ( "
strSQL = strSQL & "    SELECT "
strSQL = strSQL & "        p.userid, "
strSQL = strSQL & "        p.pay_price, "
strSQL = strSQL & "        CASE "
strSQL = strSQL & "            WHEN fm.b_year IS NULL "
strSQL = strSQL & "                THEN '미분류' "
strSQL = strSQL & "            WHEN ISNUMERIC(fm.b_year) = 0 "
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
strSQL = strSQL & "        END AS age_range "
strSQL = strSQL & "    FROM pay_info AS p WITH (READUNCOMMITTED) "
strSQL = strSQL & "    LEFT JOIN FamilyM.aes.f_user_add_site('www.ybmteachers.com') AS fm "
strSQL = strSQL & "        ON p.userid = fm.userid "
strSQL = strSQL & "    WHERE p.sell_info = 'o' "
strSQL = strSQL & "      AND p.pay_date >= '" & startDate & "' "
strSQL = strSQL & "      AND p.pay_date < DATEADD(DAY, 1, '" & endDate & "') "
strSQL = strSQL & ") "
strSQL = strSQL & "SELECT TOP 5 "
strSQL = strSQL & "    age_range, "
strSQL = strSQL & "    COUNT(DISTINCT userid) AS user_count, "
strSQL = strSQL & "    SUM(ISNULL(pay_price, 0)) AS actual_amount "
strSQL = strSQL & "FROM AgeData "
strSQL = strSQL & "GROUP BY age_range "
strSQL = strSQL & "ORDER BY "
strSQL = strSQL & "    SUM(ISNULL(pay_price, 0)) DESC "

arrTopAge = ExecSP(strSQL, strCyberTeacher)

If IsArray(arrTopAge) Then
    hasTopAge = True
End If

'==========================================================
' 7. 담당과목별 상위 판매 TOP 5
'
' 회원별 담당과목을 먼저 하나로 정리한다.
' 한 회원이 여러 과목을 가지고 있어도 결제액이
' 여러 과목에 중복되지 않도록 한다.
'==========================================================
strSQL = ""
strSQL = strSQL & "SELECT TOP 5 "
strSQL = strSQL & "    ISNULL(t.subject, '미분류') AS subject_name, "
strSQL = strSQL & "    COUNT(DISTINCT p.userid) AS user_count, "
strSQL = strSQL & "    SUM(ISNULL(p.pay_price, 0)) AS actual_amount "
strSQL = strSQL & "FROM pay_info AS p WITH (READUNCOMMITTED) "
strSQL = strSQL & "INNER JOIN ( "
strSQL = strSQL & "    SELECT "
strSQL = strSQL & "        userid, "
strSQL = strSQL & "        MIN(NULLIF(LTRIM(RTRIM(subject)), '')) AS subject "
strSQL = strSQL & "    FROM vTeacher_info WITH (READUNCOMMITTED) "
strSQL = strSQL & "    GROUP BY userid "
strSQL = strSQL & ") AS t "
strSQL = strSQL & "    ON p.userid = t.userid "
strSQL = strSQL & "WHERE p.sell_info = 'o' "
strSQL = strSQL & "  AND p.pay_date >= '" & startDate & "' "
strSQL = strSQL & "  AND p.pay_date < DATEADD(DAY, 1, '" & endDate & "') "
strSQL = strSQL & "GROUP BY t.subject "
strSQL = strSQL & "ORDER BY "
strSQL = strSQL & "    SUM(ISNULL(p.pay_price, 0)) DESC "

arrTopSubject = ExecSP(strSQL, strCyberTeacher)

If IsArray(arrTopSubject) Then
    hasTopSubject = True
End If

'==========================================================
' 개인정보 파일 다운로드 로그
'==========================================================
If hasOverview Then

    strSQL = ""
    strSQL = strSQL & "EXEC site_log.dbo.p_File_Log "
    strSQL = strSQL & fnCdbValue(Session("admin_id")) & ","
    strSQL = strSQL & fnCdbValue(Request.ServerVariables("REMOTE_ADDR")) & ","
    strSQL = strSQL & fnCdbValue(Request.ServerVariables("HTTP_HOST")) & ","
    strSQL = strSQL & fnCdbValue("stat_dashboard_" & downloadDate) & ","
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
    "attachment; filename=stat_dashboard_" & _
    downloadDate & ".xls"
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">

<style>
body {
    font-family: gulim;
    font-size: 12px;
}

table {
    border-collapse: collapse;
    margin-bottom: 25px;
}

td,
th {
    border: 1px solid #999;
    padding: 7px 10px;
}

.section-title {
    font-weight: bold;
    font-size: 14px;
    background-color: #D3D3D3;
}

.section-header {
    background-color: #4472C4;
    color: white;
    font-weight: bold;
}

.kpi-box {
    background-color: #F0F8FF;
    border: 1px solid #4472C4;
    padding: 12px;
    text-align: center;
}

.kpi-label {
    font-size: 11px;
    color: #666;
    margin-bottom: 5px;
}

.kpi-value {
    font-size: 17px;
    font-weight: bold;
    color: #4472C4;
}

.kpi-unit {
    font-size: 11px;
    color: #777;
}

.notice {
    color: #666;
    font-size: 11px;
}
</style>
</head>

<body>

<h2>통합 대시보드</h2>

<p>
    <strong>조회 기간:</strong>
    <%=Server.HTMLEncode(SqlDisplayDate(startDay))%>
    ~
    <%=Server.HTMLEncode(SqlDisplayDate(endDay))%>
</p>

<p class="notice">
    결제 기준: 결제완료 상태(sell_info = 'o')의 pay_info 데이터입니다.
    구매자수는 기간 내 고유 회원수이며, 결제건수는 고유 결제번호 기준입니다.<br>
    지역은 school_info.sido_position → school_category.scate_num 기준으로 분류합니다.
</p>

<!--======================================================
    KPI
=======================================================-->
<table width="1000" border="0" cellspacing="8" cellpadding="0">
    <tr>
        <td colspan="4" class="section-title">
            핵심 지표 (KPI)
        </td>
    </tr>

    <tr>
        <td width="25%">
            <div class="kpi-box">
                <div class="kpi-label">총 구매자수</div>
                <div class="kpi-value">
                    <%=FormatNumber(totalUser, 0)%>
                    <span class="kpi-unit">명</span>
                </div>
            </div>
        </td>

        <td width="25%">
            <div class="kpi-box">
                <div class="kpi-label">총 결제건수</div>
                <div class="kpi-value">
                    <%=FormatNumber(totalPayCount, 0)%>
                    <span class="kpi-unit">건</span>
                </div>
            </div>
        </td>

        <td width="25%">
            <div class="kpi-box">
                <div class="kpi-label">총 결제액</div>
                <div class="kpi-value">
                    <%=FormatNumber(totalActual, 0)%>
                    <span class="kpi-unit">원</span>
                </div>
            </div>
        </td>

        <td width="25%">
            <div class="kpi-box">
                <div class="kpi-label">1인당 평균 결제액</div>
                <div class="kpi-value">
                    <%=FormatNumber(avgPayAmount, 0)%>
                    <span class="kpi-unit">원</span>
                </div>
            </div>
        </td>
    </tr>

    <tr>
        <td width="25%">
            <div class="kpi-box">
                <div class="kpi-label">쿠폰 사용자</div>
                <div class="kpi-value">
                    <%=FormatNumber(couponUser, 0)%>
                    <span class="kpi-unit">명</span>
                </div>
            </div>
        </td>

        <td width="25%">
            <div class="kpi-box">
                <div class="kpi-label">SMS 동의자</div>
                <div class="kpi-value">
                    <%=FormatNumber(smsUser, 0)%>
                    <span class="kpi-unit">명</span>
                </div>
            </div>
        </td>

        <td width="25%">
            <div class="kpi-box">
                <div class="kpi-label">쿠폰 사용자 비율</div>
                <div class="kpi-value">
                    <%
                    If totalUser > 0 Then
                        Response.Write FormatNumber(couponUser / totalUser, 4)
                    Else
                        Response.Write "0.00%"
                    End If
                    %>
                </div>
            </div>
        </td>

        <td width="25%">
            <div class="kpi-box">
                <div class="kpi-label">SMS 동의율</div>
                <div class="kpi-value">
                    <%
                    If totalUser > 0 Then
                        Response.Write FormatNumber(smsUser / totalUser, 4)
                    Else
                        Response.Write "0.00%"
                    End If
                    %>
                </div>
            </div>
        </td>
    </tr>
</table>

<!--======================================================
    상위 판매 과정
=======================================================-->
<table width="800" border="1">
    <tr>
        <td colspan="5" class="section-title">
            상위 판매 과정 TOP 5
        </td>
    </tr>

    <tr class="section-header">
        <th width="50">순위</th>
        <th width="350">과정명</th>
        <th width="120">판매수</th>
        <th width="150">결제액</th>
        <th width="120">비고</th>
    </tr>

<%
If hasTopProduct Then

    For i = 0 To UBound(arrTopProduct, 2)

        itemName = Nz(arrTopProduct(0, i), "과정명 없음")
        itemCount = ToNumber(arrTopProduct(1, i))
        itemAmount = ToNumber(arrTopProduct(2, i))
%>
    <tr>
        <td align="center"><%=(i + 1)%></td>
        <td><%=Server.HTMLEncode(itemName & "")%></td>
        <td align="right">
            <%=FormatNumber(itemCount, 0)%>
        </td>
        <td align="right">
            <%=FormatNumber(itemAmount, 0)%>
        </td>
        <td align="center">판매수 기준</td>
    </tr>
<%
    Next

Else
%>
    <tr>
        <td colspan="5" align="center">
            조회 결과가 없습니다.
        </td>
    </tr>
<%
End If
%>
</table>

<!--======================================================
    지역별 상위 판매
=======================================================-->
<table width="650" border="1">
    <tr>
        <td colspan="3" class="section-title">
            지역별 상위 판매 TOP 5
        </td>
    </tr>

    <tr class="section-header">
        <th width="50">순위</th>
        <th width="250">지역명</th>
        <th width="180">결제액</th>
    </tr>

<%
If hasTopRegion Then

    For i = 0 To UBound(arrTopRegion, 2)

        itemName = Nz(arrTopRegion(0, i), "미분류")
        itemAmount = ToNumber(arrTopRegion(2, i))
%>
    <tr>
        <td align="center"><%=(i + 1)%></td>
        <td><%=Server.HTMLEncode(itemName & "")%></td>
        <td align="right">
            <%=FormatNumber(itemAmount, 0)%>
        </td>
    </tr>
<%
    Next

Else
%>
    <tr>
        <td colspan="3" align="center">
            조회 결과가 없습니다.
        </td>
    </tr>
<%
End If
%>
</table>

<!--======================================================
    연령대별 상위 판매
=======================================================-->
<table width="650" border="1">
    <tr>
        <td colspan="3" class="section-title">
            연령대별 상위 판매 TOP 5
        </td>
    </tr>

    <tr class="section-header">
        <th width="50">순위</th>
        <th width="250">연령대</th>
        <th width="180">결제액</th>
    </tr>

<%
If hasTopAge Then

    For i = 0 To UBound(arrTopAge, 2)

        itemName = Nz(arrTopAge(0, i), "미분류")
        itemAmount = ToNumber(arrTopAge(2, i))
%>
    <tr>
        <td align="center"><%=(i + 1)%></td>
        <td><%=Server.HTMLEncode(itemName & "")%></td>
        <td align="right">
            <%=FormatNumber(itemAmount, 0)%>
        </td>
    </tr>
<%
    Next

Else
%>
    <tr>
        <td colspan="3" align="center">
            조회 결과가 없습니다.
        </td>
    </tr>
<%
End If
%>
</table>

<!--======================================================
    담당과목별 상위 판매
=======================================================-->
<table width="650" border="1">
    <tr>
        <td colspan="3" class="section-title">
            담당과목별 상위 판매 TOP 5
        </td>
    </tr>

    <tr class="section-header">
        <th width="50">순위</th>
        <th width="250">담당과목</th>
        <th width="180">결제액</th>
    </tr>

<%
If hasTopSubject Then

    For i = 0 To UBound(arrTopSubject, 2)

        itemName = Nz(arrTopSubject(0, i), "미분류")
        itemAmount = ToNumber(arrTopSubject(2, i))
%>
    <tr>
        <td align="center"><%=(i + 1)%></td>
        <td><%=Server.HTMLEncode(itemName & "")%></td>
        <td align="right">
            <%=FormatNumber(itemAmount, 0)%>
        </td>
    </tr>
<%
    Next

Else
%>
    <tr>
        <td colspan="3" align="center">
            조회 결과가 없습니다.
        </td>
    </tr>
<%
End If
%>
</table>

</body>
</html>

<%
'==========================================================
' 함수
'==========================================================

' 요청 날짜를 읽는다.
Function ReadReportDate(ByVal yearKey, ByVal monthKey, ByVal dayKey, ByVal defaultDate)

    Dim y
    Dim m
    Dim d
    Dim candidate

    y = Trim(Request(yearKey))
    m = Trim(Request(monthKey))
    d = Trim(Request(dayKey))

    If y = "" Or Not IsNumeric(y) Then
        y = Year(defaultDate)
    End If

    If m = "" Or Not IsNumeric(m) Then
        m = Month(defaultDate)
    End If

    If d = "" Or Not IsNumeric(d) Then
        d = Day(defaultDate)
    End If

    On Error Resume Next
    candidate = DateSerial(CLng(y), CLng(m), CLng(d))

    If Err.Number <> 0 Then
        Err.Clear
        candidate = defaultDate
    End If

    On Error GoTo 0

    ReadReportDate = candidate

End Function

' SQL 문자열에 사용할 날짜 형식
Function SqlDateText(ByVal value)

    SqlDateText = Year(value) & "-" & _
                  Right("0" & Month(value), 2) & "-" & _
                  Right("0" & Day(value), 2)

End Function

' 화면 표시용 날짜 형식
Function SqlDisplayDate(ByVal value)

    SqlDisplayDate = Year(value) & "-" & _
                     Right("0" & Month(value), 2) & "-" & _
                     Right("0" & Day(value), 2)

End Function

' 숫자 변환
Function ToNumber(ByVal value)

    If IsNull(value) Then
        ToNumber = 0
    ElseIf IsEmpty(value) Then
        ToNumber = 0
    ElseIf IsNumeric(value) Then
        ToNumber = CDbl(value)
    Else
        ToNumber = 0
    End If

End Function

' NULL·빈 값 처리
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