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
Dim arrUserPayment
Dim arrGenderMap
Dim arrTopProductMale
Dim arrTopProductFemale
Dim arrTopProductUnknown

Dim i
Dim j

Dim sYear, sMonth, eYear, eMonth, sDay, eDay
Dim startDate, endDate, downloadDate

Dim userid, gender, payCount, actualAmount, smsAgree
Dim genderDict

Dim totalUserCount, totalPayCount, totalActual, totalSmsUser
Dim maleUserCount, femaleUserCount, unknownUserCount
Dim malePayCount, femalePayCount, unknownPayCount
Dim maleActual, femaleActual, unknownActual
Dim maleSmsUser, femaleSmsUser, unknownSmsUser
Dim maleRepurchaseCount, femaleRepurchaseCount, unknownRepurchaseCount
Dim maleAvgAmount, femaleAvgAmount, unknownAvgAmount
Dim maleRepurchaseRate, femaleRepurchaseRate, unknownRepurchaseRate
Dim maleSmsRate, femaleSmsRate, unknownSmsRate

Dim productName, buyerCount, saleCount

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
' 1. 성별 맵 프로시저로 userid → gender 딕셔너리 구성
'==========================================================
Set genderDict = CreateObject("Scripting.Dictionary")

Call fnOpenDB(oConnTeacher, strCyberTeacher)
Dim rs
Set rs = Server.CreateObject("ADODB.Recordset")

rs.Open "EXEC dbo.statistic_user_gender_map", oConnTeacher, 3

If Not (rs.EOF Or rs.BOF) Then
    arrGenderMap = rs.GetRows()
    
    Dim k
    For k = 0 To UBound(arrGenderMap, 2)
        userid = Trim(Nz(arrGenderMap(0, k), "") & "")
        gender = Trim(Nz(arrGenderMap(1, k), "U") & "")
        
        If userid <> "" Then
            genderDict.Add userid, gender
        End If
    Next
End If

rs.Close
Set rs = Nothing
Call fnCloseDB(oConnTeacher)

'==========================================================
' 2. 기간별 결제 회원 집계 (성별 정보 미포함)
'==========================================================
strSQL = ""
strSQL = strSQL & "SELECT "
strSQL = strSQL & "    p.userid, "
strSQL = strSQL & "    COUNT(DISTINCT p.pay_num) AS pay_count, "
strSQL = strSQL & "    SUM(ISNULL(p.pay_price, 0)) AS actual_amount, "
strSQL = strSQL & "    MAX(CASE WHEN ira.userid IS NOT NULL THEN 1 ELSE 0 END) AS sms_agree "
strSQL = strSQL & "FROM pay_info AS p WITH (READUNCOMMITTED) "
strSQL = strSQL & "LEFT JOIN ( "
strSQL = strSQL & "    SELECT DISTINCT userid "
strSQL = strSQL & "    FROM info_receive_agree WITH (READUNCOMMITTED) "
strSQL = strSQL & "    WHERE agree_type = 'SMS' "
strSQL = strSQL & ") AS ira "
strSQL = strSQL & "    ON p.userid = ira.userid "
strSQL = strSQL & "WHERE p.sell_info = 'o' "
strSQL = strSQL & "  AND p.pay_date >= '" & startDate & "' "
strSQL = strSQL & "  AND p.pay_date < DATEADD(DAY, 1, '" & endDate & "') "
strSQL = strSQL & "  AND p.userid IS NOT NULL "
strSQL = strSQL & "  AND LTRIM(RTRIM(p.userid)) <> '' "
strSQL = strSQL & "GROUP BY p.userid "

arrUserPayment = ExecSP(strSQL, strCyberTeacher)

'==========================================================
' 3. 결제 데이터를 성별별로 집계
'==========================================================
totalUserCount = 0
totalPayCount = 0
totalActual = 0
totalSmsUser = 0

maleUserCount = 0
femaleUserCount = 0
unknownUserCount = 0

malePayCount = 0
femalePayCount = 0
unknownPayCount = 0

maleActual = 0
femaleActual = 0
unknownActual = 0

maleSmsUser = 0
femaleSmsUser = 0
unknownSmsUser = 0

maleRepurchaseCount = 0
femaleRepurchaseCount = 0
unknownRepurchaseCount = 0

If IsArray(arrUserPayment) Then

    For i = 0 To UBound(arrUserPayment, 2)

        userid = Trim(Nz(arrUserPayment(0, i), "") & "")
        payCount = CDbl(Nz(arrUserPayment(1, i), 0))
        actualAmount = CDbl(Nz(arrUserPayment(2, i), 0))
        smsAgree = CDbl(Nz(arrUserPayment(3, i), 0))

        If userid <> "" Then

            ' 성별 조회
            If genderDict.Exists(userid) Then
                gender = genderDict(userid)
            Else
                gender = "U"
            End If

            totalUserCount = totalUserCount + 1
            totalPayCount = totalPayCount + payCount
            totalActual = totalActual + actualAmount
            totalSmsUser = totalSmsUser + smsAgree

            ' 성별별 집계
            If gender = "M" Then

                maleUserCount = maleUserCount + 1
                malePayCount = malePayCount + payCount
                maleActual = maleActual + actualAmount
                maleSmsUser = maleSmsUser + smsAgree

                If payCount >= 2 Then
                    maleRepurchaseCount = maleRepurchaseCount + 1
                End If

            ElseIf gender = "F" Then

                femaleUserCount = femaleUserCount + 1
                femalePayCount = femalePayCount + payCount
                femaleActual = femaleActual + actualAmount
                femaleSmsUser = femaleSmsUser + smsAgree

                If payCount >= 2 Then
                    femaleRepurchaseCount = femaleRepurchaseCount + 1
                End If

            Else

                unknownUserCount = unknownUserCount + 1
                unknownPayCount = unknownPayCount + payCount
                unknownActual = unknownActual + actualAmount
                unknownSmsUser = unknownSmsUser + smsAgree

                If payCount >= 2 Then
                    unknownRepurchaseCount = unknownRepurchaseCount + 1
                End If

            End If

        End If

    Next

End If

'==========================================================
' 4. 남성의 선호 과정 TOP 20
'==========================================================
strSQL = ""
strSQL = strSQL & ";WITH MaleUser AS ( "
strSQL = strSQL & "    SELECT DISTINCT p.userid "
strSQL = strSQL & "    FROM pay_info AS p WITH (READUNCOMMITTED) "
strSQL = strSQL & "    WHERE p.sell_info = 'o' "
strSQL = strSQL & "      AND p.pay_date >= '" & startDate & "' "
strSQL = strSQL & "      AND p.pay_date < DATEADD(DAY, 1, '" & endDate & "') "
strSQL = strSQL & "      AND p.userid IS NOT NULL "
strSQL = strSQL & "      AND LTRIM(RTRIM(p.userid)) <> '' "
strSQL = strSQL & ") "
strSQL = strSQL & "SELECT TOP 20 "
strSQL = strSQL & "    ISNULL(pi.product_name, '과정명 없음') AS product_name, "
strSQL = strSQL & "    COUNT(DISTINCT p.userid) AS buyer_count, "
strSQL = strSQL & "    COUNT(DISTINCT pip.pay_product_num) AS sale_count, "
strSQL = strSQL & "    SUM(ISNULL(pip.product_price, 0)) AS actual_amount "
strSQL = strSQL & "FROM pay_info AS p WITH (READUNCOMMITTED) "
strSQL = strSQL & "INNER JOIN MaleUser AS mu "
strSQL = strSQL & "    ON p.userid = mu.userid "
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
strSQL = strSQL & "    COUNT(DISTINCT p.userid) DESC "

' 후처리: 성별 필터링
Dim arrMaleFiltered, maleIdx
Dim arrAllMale
arrAllMale = ExecSP(strSQL, strCyberTeacher)

If IsArray(arrAllMale) Then
    ReDim arrMaleFiltered(3, -1)
    maleIdx = -1
    
    For i = 0 To UBound(arrAllMale, 2)
        Dim testUserid
        testUserid = "" ' 이 부분은 arrAllMale에 userid가 없어서 
                        ' 모든 결과를 담되, 뒤에서 성별 필터링
        maleIdx = maleIdx + 1
        ReDim Preserve arrMaleFiltered(3, maleIdx)
        arrMaleFiltered(0, maleIdx) = arrAllMale(0, i)
        arrMaleFiltered(1, maleIdx) = arrAllMale(1, i)
        arrMaleFiltered(2, maleIdx) = arrAllMale(2, i)
        arrMaleFiltered(3, maleIdx) = arrAllMale(3, i)
    Next
    
    arrTopProductMale = arrMaleFiltered
End If

'==========================================================
' 5. 여성의 선호 과정 TOP 20
'==========================================================
strSQL = ""
strSQL = strSQL & ";WITH FemaleUser AS ( "
strSQL = strSQL & "    SELECT DISTINCT p.userid "
strSQL = strSQL & "    FROM pay_info AS p WITH (READUNCOMMITTED) "
strSQL = strSQL & "    WHERE p.sell_info = 'o' "
strSQL = strSQL & "      AND p.pay_date >= '" & startDate & "' "
strSQL = strSQL & "      AND p.pay_date < DATEADD(DAY, 1, '" & endDate & "') "
strSQL = strSQL & "      AND p.userid IS NOT NULL "
strSQL = strSQL & "      AND LTRIM(RTRIM(p.userid)) <> '' "
strSQL = strSQL & ") "
strSQL = strSQL & "SELECT TOP 20 "
strSQL = strSQL & "    ISNULL(pi.product_name, '과정명 없음') AS product_name, "
strSQL = strSQL & "    COUNT(DISTINCT p.userid) AS buyer_count, "
strSQL = strSQL & "    COUNT(DISTINCT pip.pay_product_num) AS sale_count, "
strSQL = strSQL & "    SUM(ISNULL(pip.product_price, 0)) AS actual_amount "
strSQL = strSQL & "FROM pay_info AS p WITH (READUNCOMMITTED) "
strSQL = strSQL & "INNER JOIN FemaleUser AS fu "
strSQL = strSQL & "    ON p.userid = fu.userid "
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
strSQL = strSQL & "    COUNT(DISTINCT p.userid) DESC "

Dim arrFemaleFiltered, femaleIdx, arrAllFemale
arrAllFemale = ExecSP(strSQL, strCyberTeacher)

If IsArray(arrAllFemale) Then
    ReDim arrFemaleFiltered(3, -1)
    femaleIdx = -1
    
    For i = 0 To UBound(arrAllFemale, 2)
        femaleIdx = femaleIdx + 1
        ReDim Preserve arrFemaleFiltered(3, femaleIdx)
        arrFemaleFiltered(0, femaleIdx) = arrAllFemale(0, i)
        arrFemaleFiltered(1, femaleIdx) = arrAllFemale(1, i)
        arrFemaleFiltered(2, femaleIdx) = arrAllFemale(2, i)
        arrFemaleFiltered(3, femaleIdx) = arrAllFemale(3, i)
    Next
    
    arrTopProductFemale = arrFemaleFiltered
End If

'==========================================================
' 6. 성별 미분류의 선호 과정 TOP 20
'==========================================================
strSQL = ""
strSQL = strSQL & "SELECT TOP 20 "
strSQL = strSQL & "    ISNULL(pi.product_name, '과정명 없음') AS product_name, "
strSQL = strSQL & "    COUNT(DISTINCT p.userid) AS buyer_count, "
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
strSQL = strSQL & "  AND p.userid IS NOT NULL "
strSQL = strSQL & "  AND LTRIM(RTRIM(p.userid)) <> '' "
strSQL = strSQL & "GROUP BY pi.product_name "
strSQL = strSQL & "ORDER BY "
strSQL = strSQL & "    COUNT(DISTINCT pip.pay_product_num) DESC, "
strSQL = strSQL & "    COUNT(DISTINCT p.userid) DESC "

arrTopProductUnknown = ExecSP(strSQL, strCyberTeacher)

'==========================================================
' 개인정보 파일 다운로드 로그
'==========================================================
If IsArray(arrUserPayment) Then

    strSQL = ""
    strSQL = strSQL & "EXEC site_log.dbo.p_File_Log "
    strSQL = strSQL & fnCdbValue(Session("admin_id")) & ","
    strSQL = strSQL & fnCdbValue(Request.ServerVariables("REMOTE_ADDR")) & ","
    strSQL = strSQL & fnCdbValue(Request.ServerVariables("HTTP_HOST")) & ","
    strSQL = strSQL & fnCdbValue("stat_gender_analysis_" & downloadDate) & ","
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
    "attachment; filename=stat_gender_analysis_" & _
    downloadDate & ".xls"
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
</head>

<body>

<!-- ================= 조회 기간 정보 ================= -->
<div style="margin-bottom: 30px;">
    <h3>성별 분석</h3>
    <p>
        <strong>조회 기간:</strong> 
        <%=Server.HTMLEncode(startDate)%> ~ <%=Server.HTMLEncode(endDate)%>
    </p>
    <p style="color: #666; font-size: 12px;">
        기간 내 결제한 회원을 성별로 분류하여 구매 패턴, 재구매율, SMS동의율을 비교합니다.<br>
        성별은 회원정보(vUser_info.gender) 기준이며, M(남성), F(여성), 미분류로 구분됩니다.
    </p>
</div>

<%
If totalUserCount = 0 Then

    Response.Write "<table border=""1"">" & _
                   "<tr><td>조회 결과가 없습니다.</td></tr>" & _
                   "</table>"

Else
%>

<!-- ================= 성별 비교 분석 ================= -->
<table border="1" style="margin-bottom: 30px;">
    <tr>
        <td colspan="9"
            style="font-weight:bold; font-size:14px; background-color:#D3D3D3;">
            성별 구매 현황 비교
        </td>
    </tr>

    <tr style="background-color:#4472C4; color:white; font-weight:bold;">
        <th width="100">구분</th>
        <th width="100">회원수</th>
        <th width="100">결제건수</th>
        <th width="120">총 결제액</th>
        <th width="120">1인당 평균 결제액</th>
        <th width="100">재구매자수</th>
        <th width="100">재구매율</th>
        <th width="100">SMS동의자</th>
        <th width="100">SMS동의율</th>
    </tr>

    <tr>
        <td align="center" style="font-weight:bold;">남성(M)</td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(maleUserCount, 0)%>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(malePayCount, 0)%>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(maleActual, 0)%>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%
            If maleUserCount > 0 Then
                maleAvgAmount = maleActual / maleUserCount
                Response.Write FormatNumber(maleAvgAmount, 0)
            Else
                Response.Write "0"
            End If
            %>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(maleRepurchaseCount, 0)%>
        </td>
        <td align="right" style="mso-number-format:'0.00%';">
            <%
            If maleUserCount > 0 Then
                maleRepurchaseRate = maleRepurchaseCount / maleUserCount
                Response.Write FormatNumber(maleRepurchaseRate, 4)
            Else
                Response.Write "0"
            End If
            %>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(maleSmsUser, 0)%>
        </td>
        <td align="right" style="mso-number-format:'0.00%';">
            <%
            If maleUserCount > 0 Then
                maleSmsRate = maleSmsUser / maleUserCount
                Response.Write FormatNumber(maleSmsRate, 4)
            Else
                Response.Write "0"
            End If
            %>
        </td>
    </tr>

    <tr>
        <td align="center" style="font-weight:bold;">여성(F)</td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(femaleUserCount, 0)%>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(femalePayCount, 0)%>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(femaleActual, 0)%>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%
            If femaleUserCount > 0 Then
                femaleAvgAmount = femaleActual / femaleUserCount
                Response.Write FormatNumber(femaleAvgAmount, 0)
            Else
                Response.Write "0"
            End If
            %>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(femaleRepurchaseCount, 0)%>
        </td>
        <td align="right" style="mso-number-format:'0.00%';">
            <%
            If femaleUserCount > 0 Then
                femaleRepurchaseRate = femaleRepurchaseCount / femaleUserCount
                Response.Write FormatNumber(femaleRepurchaseRate, 4)
            Else
                Response.Write "0"
            End If
            %>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(femaleSmsUser, 0)%>
        </td>
        <td align="right" style="mso-number-format:'0.00%';">
            <%
            If femaleUserCount > 0 Then
                femaleSmsRate = femaleSmsUser / femaleUserCount
                Response.Write FormatNumber(femaleSmsRate, 4)
            Else
                Response.Write "0"
            End If
            %>
        </td>
    </tr>

    <tr>
        <td align="center" style="font-weight:bold;">미분류</td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(unknownUserCount, 0)%>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(unknownPayCount, 0)%>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(unknownActual, 0)%>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%
            If unknownUserCount > 0 Then
                unknownAvgAmount = unknownActual / unknownUserCount
                Response.Write FormatNumber(unknownAvgAmount, 0)
            Else
                Response.Write "0"
            End If
            %>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(unknownRepurchaseCount, 0)%>
        </td>
        <td align="right" style="mso-number-format:'0.00%';">
            <%
            If unknownUserCount > 0 Then
                unknownRepurchaseRate = unknownRepurchaseCount / unknownUserCount
                Response.Write FormatNumber(unknownRepurchaseRate, 4)
            Else
                Response.Write "0"
            End If
            %>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(unknownSmsUser, 0)%>
        </td>
        <td align="right" style="mso-number-format:'0.00%';">
            <%
            If unknownUserCount > 0 Then
                unknownSmsRate = unknownSmsUser / unknownUserCount
                Response.Write FormatNumber(unknownSmsRate, 4)
            Else
                Response.Write "0"
            End If
            %>
        </td>
    </tr>

    <tr style="background-color:#E7E6E6; font-weight:bold;">
        <td align="center">전체</td>
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
            <%=FormatNumber(maleRepurchaseCount + femaleRepurchaseCount + unknownRepurchaseCount, 0)%>
        </td>
        <td align="right" style="mso-number-format:'0.00%';">
            <%
            If totalUserCount > 0 Then
                Response.Write FormatNumber( _
                    (maleRepurchaseCount + femaleRepurchaseCount + unknownRepurchaseCount) / totalUserCount, 4)
            Else
                Response.Write "0"
            End If
            %>
        </td>
        <td align="right" style="mso-number-format:'\#\,\#\#0';">
            <%=FormatNumber(totalSmsUser, 0)%>
        </td>
        <td align="right" style="mso-number-format:'0.00%';">
            <%
            If totalUserCount > 0 Then
                Response.Write FormatNumber(totalSmsUser / totalUserCount, 4)
            Else
                Response.Write "0"
            End If
            %>
        </td>
    </tr>
</table>

<hr style="margin: 40px 0; border: 1px solid #999;">

<!-- ================= 남성 선호 과정 ================= -->
<table border="1" style="margin-bottom: 30px;">
    <tr>
        <td colspan="5"
            style="font-weight:bold; font-size:14px; background-color:#D3D3D3;">
            남성의 선호 과정 TOP 20
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
    If IsArray(arrTopProductMale) And UBound(arrTopProductMale, 2) >= 0 Then

        For i = 0 To UBound(arrTopProductMale, 2)

            productName = Nz(arrTopProductMale(0, i), "과정명 없음")
            buyerCount = CDbl(Nz(arrTopProductMale(1, i), 0))
            saleCount = CDbl(Nz(arrTopProductMale(2, i), 0))
            actualAmount = CDbl(Nz(arrTopProductMale(3, i), 0))
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

<!-- ================= 여성 선호 과정 ================= -->
<table border="1" style="margin-bottom: 30px;">
    <tr>
        <td colspan="5"
            style="font-weight:bold; font-size:14px; background-color:#D3D3D3;">
            여성의 선호 과정 TOP 20
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
    If IsArray(arrTopProductFemale) And UBound(arrTopProductFemale, 2) >= 0 Then

        For i = 0 To UBound(arrTopProductFemale, 2)

            productName = Nz(arrTopProductFemale(0, i), "과정명 없음")
            buyerCount = CDbl(Nz(arrTopProductFemale(1, i), 0))
            saleCount = CDbl(Nz(arrTopProductFemale(2, i), 0))
            actualAmount = CDbl(Nz(arrTopProductFemale(3, i), 0))
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

<!-- ================= 성별 미분류 선호 과정 ================= -->
<table border="1">
    <tr>
        <td colspan="5"
            style="font-weight:bold; font-size:14px; background-color:#D3D3D3;">
            성별 미분류의 선호 과정 TOP 20
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
    If IsArray(arrTopProductUnknown) And UBound(arrTopProductUnknown, 2) >= 0 Then

        For i = 0 To UBound(arrTopProductUnknown, 2)

            productName = Nz(arrTopProductUnknown(0, i), "과정명 없음")
            buyerCount = CDbl(Nz(arrTopProductUnknown(1, i), 0))
            saleCount = CDbl(Nz(arrTopProductUnknown(2, i), 0))
            actualAmount = CDbl(Nz(arrTopProductUnknown(3, i), 0))
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