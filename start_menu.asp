<%@ LANGUAGE="VBScript" %>
<%
Option Explicit
Response.Expires = -1

Dim menuTitle
menuTitle = "2027 사업계획 통계"
%>
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<title><%=menuTitle%></title>
<style>
body { font-family: gulim; font-size:12px; }
table { border-collapse:collapse; width:1000px; }
td, th { border:1px solid #999; padding:8px 10px; text-align:left; }
a { text-decoration:none; color:#0d5; font-weight:bold; }
</style>
</head>
<body>
<h2><%=menuTitle%></h2>

<table>
<tr>
  <th>통계명</th>
  <th>설명</th>
  <th>다운로드</th>
</tr>

<tr>
  <td><a href="stat_monthly_sales.asp">월별 매출 통계</a></td>
  <td>월별 결제건수, 구매자수, 원금액, 할인액, 실제결제액, 객단가</td>
  <td><a href="stat_monthly_sales.asp?download=1">엑셀 다운로드</a></td>
</tr>

<tr>
  <td><a href="stat_regional_analysis.asp">지역별 분석</a></td>
  <td>교육청/지역별 회원수, 구매건수, 매출액, 선호과정</td>
  <td><a href="stat_regional_analysis.asp?download=1">엑셀 다운로드</a></td>
</tr>

<tr>
  <td><a href="stat_age_analysis.asp">연령대별 분석</a></td>
  <td>연령대별 회원수, 매출액, 선호과정, SMS동의자</td>
  <td><a href="stat_age_analysis.asp?download=1">엑셀 다운로드</a></td>
</tr>

<tr>
  <td><a href="stat_subject_analysis.asp">담당과목별 분석</a></td>
  <td>담당과목별 회원수, 매출액, 구매건수, 인기과정</td>
  <td><a href="stat_subject_analysis.asp?download=1">엑셀 다운로드</a></td>
</tr>

<tr>
  <td><a href="stat_school_subject_cross.asp">학교급-담당과목 교차분석</a></td>
  <td>학교급별 과목별 분포 및 선호도</td>
  <td><a href="stat_school_subject_cross.asp?download=1">엑셀 다운로드</a></td>
</tr>

<tr>
  <td><a href="stat_top_products.asp">상위 판매 과정 순위</a></td>
  <td>과정별 판매건수, 매출, 구매자수, 지역/연령 분포</td>
  <td><a href="stat_top_products.asp?download=1">엑셀 다운로드</a></td>
</tr>

<tr>
  <td><a href="stat_repurchase_analysis.asp">2회 이상 구매자 분석</a></td>
  <td>재구매자, 평균구매횟수, 쿠폰사용여부, 선호과정</td>
  <td><a href="stat_repurchase_analysis.asp?download=1">엑셀 다운로드</a></td>
</tr>

<tr>
  <td><a href="stat_coupon_effect.asp">쿠폰 효과 분석</a></td>
  <td>쿠폰 사용 여부, 재구매율, 평균결제금액, 할인효과</td>
  <td><a href="stat_coupon_effect.asp?download=1">엑셀 다운로드</a></td>
</tr>

<tr>
  <td><a href="stat_gender_analysis.asp">성별 분석</a></td>
  <td>성별별 회원수, 구매금액, SMS동의자, 구매패턴</td>
  <td><a href="stat_gender_analysis.asp?download=1">엑셀 다운로드</a></td>
</tr>

<tr>
  <td><a href="stat_product_contribution.asp">과정별 매출 기여도</a></td>
  <td>과정별 판매현황과 매출기여도를 분석</td>
  <td><a href="stat_product_contribution.asp?download=1">엑셀 다운로드</a></td>
</tr>
<tr>
  <td><a href="stat_first_purchase_repurchase.asp">재구매 분석</a></td>
  <td>첫구매고객의 재구매 분석</td>
  <td><a href="stat_first_purchase_repurchase.asp?download=1">엑셀 다운로드</a></td>
</tr>
<tr>
  <td><a href="stat_monthly_retention_rate.asp">월간 고객 유지율 분석</a></td>
  <td>월간 고객 유지율 분석</td>
  <td><a href="stat_monthly_retention_rate.asp?download=1">엑셀 다운로드</a></td>
</tr>
<tr>
  <td><a href="stat_dashboard.asp">통합 대시보드</a></td>
  <td>핵심 요약 + KPI</td>
  <td><a href="stat_dashboard.asp?download=1">엑셀 다운로드</a></td>
</tr>
</table>

</body>
</html>