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
Dim reportYear

Dim downloadDate
Dim rowCount
Dim hasResult

Dim monthKey
Dim nextMonthKey
Dim monthStartDate
Dim nextMonthStartDate

Dim buyerCount
Dim retainedCount
Dim retentionRate

Dim yearBuyerTotal
Dim yearRetainedTotal
Dim yearRetentionRate
Dim yearRowCount

'==========================================================
' 다운로드 파일명 날짜
'==========================================================
downloadDate = Year(Now()) & _
               Right("0" & Month(Now()), 2) & _
               Right("0" & Day(Now()), 2)

'==========================================================
' 월간 고객 유지율 분석
'
' [정의]
' 기준 월 구매자 중 바로 다음 달에도 구매한 회원의 비율
'
' [계산식]
' 다음 달 유지자 수 / 기준 월 구매자 수
'
' [집계 기준]
' 1. sell_info = 'o'인 결제완료 건만 포함
' 2. userid가 있는 결제만 포함
' 3. 같은 회원이 같은 달에 여러 번 결제해도 1명으로 집계
' 4. 기준 월과 다음 달에 모두 존재하는 userid만 유지자로 집계
' 5. 12월에서 다음 해 1월로 넘어가는 비교도 포함
' 6. 다음 달이 완전히 종료된 비교만 표시
'
' [예시]
' 2026년 7월 구매자 중 2026년 8월에도 구매한 회원을
' 2026년 7월의 다음 달 유지자로 계산
'
' [현재 날짜 기준]
' 2026년 9월 30일에는 9월이 아직 끝나지 않았으므로
' 2026년 8월 → 9월 유지율은 표시하지 않음
'==========================================================
strSQL = ""

strSQL = strSQL & "WITH MonthlyBuyers AS ( "

strSQL = strSQL & "    SELECT "
strSQL = strSQL & "        LEFT(p.pay_date, 7) AS month_key, "
strSQL = strSQL & "        p.userid "

strSQL = strSQL & "    FROM pay_info AS p WITH (READUNCOMMITTED) "

strSQL = strSQL & "    WHERE p.sell_info = 'o' "
strSQL = strSQL & "      AND p.userid IS NOT NULL "
strSQL = strSQL & "      AND LTRIM(RTRIM(p.userid)) <> '' "

'----------------------------------------------------------
' 분석에 필요한 전체 월 범위
'
' 2025년 1월 유지율부터 계산
' 2026년 12월 → 2027년 1월 비교 가능성을 고려하여
' 2027년 2월 이전까지 범위를 설정
'----------------------------------------------------------
strSQL = strSQL & "      AND p.pay_date >= '2025-01-01' "
strSQL = strSQL & "      AND p.pay_date < '2027-02-01' "

'----------------------------------------------------------
' 현재 진행 중인 달은 제외
'
' 예:
' 2026-09-30 실행 시 2026년 9월 결제를 MonthlyBuyers에서
' 제외하므로 2026년 8월 → 9월 비교가 나오지 않음
'----------------------------------------------------------
strSQL = strSQL & "      AND p.pay_date < "
strSQL = strSQL & "          DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0) "

strSQL = strSQL & "    GROUP BY "
strSQL = strSQL & "        LEFT(p.pay_date, 7), "
strSQL = strSQL & "        p.userid "

strSQL = strSQL & "), "

'==========================================================
' 월별 유지자 계산
'
' cur: 기준 월 구매자
' nxt: 다음 달 구매자
'
' 반드시 userid까지 연결해야 다음 달 유지자 수가
' 기준 월 구매자 수를 초과하지 않음
'==========================================================
strSQL = strSQL & "RetentionByMonth AS ( "

strSQL = strSQL & "    SELECT "
strSQL = strSQL & "        cur.month_key, "
strSQL = strSQL & "        COUNT(cur.userid) AS buyer_count, "
strSQL = strSQL & "        COUNT(nxt.userid) AS retained_count "

strSQL = strSQL & "    FROM MonthlyBuyers AS cur "

strSQL = strSQL & "    LEFT JOIN MonthlyBuyers AS nxt "
strSQL = strSQL & "        ON nxt.userid = cur.userid "
strSQL = strSQL & "       AND nxt.month_key = CONVERT(CHAR(7), "
strSQL = strSQL & "           DATEADD(MONTH, 1, "
strSQL = strSQL & "               CONVERT(DATETIME, cur.month_key + '-01', 120) "
strSQL = strSQL & "           ), "
strSQL = strSQL & "           120 "
strSQL = strSQL & "       ) "

strSQL = strSQL & "    WHERE cur.month_key >= '2025-01' "
strSQL = strSQL & "      AND cur.month_key < '2027-01' "

'----------------------------------------------------------
' 다음 달이 완전히 종료된 기준 월만 출력
'
' 다음 달 시작일이 현재 달 시작일보다 이전이어야 함
'
' 2026-09-30 기준:
' - 2026-07 → 2026-08: 출력
' - 2026-08 → 2026-09: 제외
'----------------------------------------------------------
strSQL = strSQL & "      AND DATEADD(MONTH, 1, "
strSQL = strSQL & "          CONVERT(DATETIME, cur.month_key + '-01', 120) "
strSQL = strSQL & "      ) < DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0) "

strSQL = strSQL & "    GROUP BY "
strSQL = strSQL & "        cur.month_key "

strSQL = strSQL & ") "

'==========================================================
' 최종 결과
'==========================================================
strSQL = strSQL & "SELECT "
strSQL = strSQL & "    month_key, "
strSQL = strSQL & "    buyer_count, "
strSQL = strSQL & "    retained_count "

strSQL = strSQL & "FROM RetentionByMonth "

strSQL = strSQL & "ORDER BY "
strSQL = strSQL & "    month_key "

arrResult = ExecSP(strSQL, strCyberTeacher)

'==========================================================
' 결과 배열 확인
'==========================================================
hasResult = False
rowCount = -1

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

'==========================================================
' 개인정보 파일 다운로드 로그
'==========================================================
If hasResult Then

    strSQL = ""

    strSQL = strSQL & "EXEC site_log.dbo.p_File_Log "
    strSQL = strSQL & fnCdbValue(Session("admin_id")) & ","
    strSQL = strSQL & fnCdbValue(Request.ServerVariables("REMOTE_ADDR")) & ","
    strSQL = strSQL & fnCdbValue(Request.ServerVariables("HTTP_HOST")) & ","
    strSQL = strSQL & fnCdbValue("stat_monthly_retention_rate_" & downloadDate) & ","
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
    "attachment; filename=stat_monthly_retention_rate_" & _
    downloadDate & ".xls"
%>

<html>
<head>
<meta http-equiv="Content-Type"
      content="text/html; charset=euc-kr">

<style type="text/css">
    .description {
        color: #666666;
        font-size: 12px;
        line-height: 1.6;
    }

    .year-title {
        background-color: #D3D3D3;
        font-weight: bold;
        font-size: 14px;
    }

    .table-header {
        background-color: #4472C4;
        color: #FFFFFF;
        font-weight: bold;
    }

    .total-row {
        background-color: #E7E6E6;
        font-weight: bold;
    }
</style>
</head>

<body>

<!--======================================================
    Excel 상단 설명
=======================================================-->
<p class="description">
    <b>2025년·2026년 월간 고객 유지율 분석</b><br>
    이번 달 구매자 중 바로 다음 달에도 구매한 고객의 비율을 분석한 자료입니다.<br>
    동일한 회원이 한 달에 여러 번 결제한 경우에도 해당 월에는 1명으로 계산합니다.<br>
    다음 달 유지자 수는 기준 월과 바로 다음 달에 모두 결제한 고유 회원 수입니다.<br>
    유지율은 다음 달 유지자 수를 기준 월 구매자 수로 나누어 계산하며,
    0%에서 100% 사이의 값으로 표시됩니다.<br>
    12월 구매자가 다음 해 1월에도 구매한 경우도 유지 고객으로 포함됩니다.<br>
    다음 달이 완전히 종료된 비교만 표시하므로 진행 중인 달이 포함된 유지율은 표시하지 않습니다.<br>
    연도별 합계는 고유 회원 합계가 아니라 표시된 각 월의 구매자 수와 유지자 수를 합산한 값입니다.<br>
    연도별 유지율은 월별 인원 규모를 반영한 가중 유지율입니다.<br>
    결제완료 상태인 데이터만 분석 대상에 포함됩니다.
</p>

<hr style="margin: 20px 0; border: 1px solid #999999;">

<%
'==========================================================
' 2025년과 2026년을 각각 별도 표로 출력
'==========================================================
For reportYear = 2025 To 2026

    yearBuyerTotal = 0
    yearRetainedTotal = 0
    yearRetentionRate = 0
    yearRowCount = 0
%>

<table border="1"
       cellspacing="0"
       cellpadding="4"
       style="margin-bottom: 25px;">

    <thead>
        <tr class="year-title">
            <td colspan="5" style="padding: 8px;">
                <%=reportYear%>년 월간 고객 유지율
            </td>
        </tr>

        <tr class="table-header">
            <th width="110">기준 월</th>
            <th width="150">기준 월 구매자 수</th>
            <th width="150">다음 달 유지자 수</th>
            <th width="120">유지율</th>
            <th width="200">비교 기간</th>
        </tr>
    </thead>

    <tbody>

<%
    If hasResult Then

        For i = 0 To rowCount

            monthKey = Nz(arrResult(0, i), "")

            If Left(monthKey, 4) = CStr(reportYear) Then

                buyerCount = CLng(Nz(arrResult(1, i), 0))
                retainedCount = CLng(Nz(arrResult(2, i), 0))

                '------------------------------------------------
                ' 안전장치
                '
                ' 정상 SQL이라면 유지자 수가 구매자 수를
                ' 초과할 수 없음
                '
                ' 예상하지 못한 데이터 중복이 있더라도
                ' 잘못된 100% 초과 값이 출력되지 않도록 방어
                '------------------------------------------------
                If retainedCount > buyerCount Then
                    retainedCount = buyerCount
                End If

                If buyerCount > 0 Then
                    retentionRate = retainedCount / buyerCount
                Else
                    retentionRate = 0
                End If

                '------------------------------------------------
                ' 다음 달 표시값 계산
                '------------------------------------------------
                monthStartDate = DateSerial( _
                    CInt(Left(monthKey, 4)), _
                    CInt(Right(monthKey, 2)), _
                    1 _
                )

                nextMonthStartDate = DateAdd("m", 1, monthStartDate)

                nextMonthKey = Year(nextMonthStartDate) & "-" & _
                               Right("0" & Month(nextMonthStartDate), 2)

                yearBuyerTotal = yearBuyerTotal + buyerCount
                yearRetainedTotal = yearRetainedTotal + retainedCount
                yearRowCount = yearRowCount + 1
%>

        <tr>
            <td align="center"
                style="mso-number-format:'\@';">
                <%=Server.HTMLEncode(monthKey)%>
            </td>

            <td align="right"
                style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(buyerCount, 0)%>
            </td>

            <td align="right"
                style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(retainedCount, 0)%>
            </td>

            <td align="right"
                style="mso-number-format:'0.00%';">
                <%=FormatNumber(retentionRate, 4)%>
            </td>

            <td align="center"
                style="mso-number-format:'\@';">
                <%=Server.HTMLEncode(monthKey & " → " & nextMonthKey)%>
            </td>
        </tr>

<%
            End If

        Next

    End If

    '======================================================
    ' 연도별 데이터가 없는 경우
    '======================================================
    If yearRowCount = 0 Then
%>

        <tr>
            <td colspan="5" align="center">
                다음 달까지 완료된 비교 데이터가 없습니다.
            </td>
        </tr>

<%
    Else

        '==================================================
        ' 연도별 가중 유지율
        '
        ' 월별 유지율을 단순 평균하지 않고
        ' 전체 유지자 수 / 전체 기준 월 구매자 수로 계산
        '==================================================
        If yearBuyerTotal > 0 Then
            yearRetentionRate = yearRetainedTotal / yearBuyerTotal
        Else
            yearRetentionRate = 0
        End If
%>

        <tr class="total-row">
            <td align="center">
                월별 인원 합계
            </td>

            <td align="right"
                style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(yearBuyerTotal, 0)%>
            </td>

            <td align="right"
                style="mso-number-format:'\#\,\#\#0';">
                <%=FormatNumber(yearRetainedTotal, 0)%>
            </td>

            <td align="right"
                style="mso-number-format:'0.00%';">
                <%=FormatNumber(yearRetentionRate, 4)%>
            </td>

            <td align="center">
                표시된 월 기준 가중 유지율
            </td>
        </tr>

<%
    End If
%>

    </tbody>
</table>

<%
Next
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