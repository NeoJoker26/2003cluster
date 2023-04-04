<%@ Language=VBScript %>
<% Option Explicit %>

<html>

<head>
<meta http-equiv="Content-Language" content="en-gb">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Greens on Screen</title>

<link rel="stylesheet" type="text/css" href="gos2.css">

<style>
<!--
p {font-size: 11px;}
-->
</style>

<script type="text/javascript"><!--
function countChar(txtBox,message1Span,message2Div)
    {
        try
        {
           count = txtBox.value.length;
           charLeft1 = Math.abs(750 - count);
           charLeft2 = Math.abs(1500 - count);
           message1Span.innerHTML=count;
            if (count < 750)
            {
                txt = charLeft1 + " characters until the minimum length; please add more";
                message2Div.innerHTML="<font color=blue>" + txt + "</font>";
            }
            
            else if (count < 1500)
            {
                txt = charLeft1 + " over the minimum and " + charLeft2 + " under the maximum; this is an ideal length";
                message2Div.innerHTML="<font color=green>" + txt + "</font>";
            }

            else
            {
                txt = charLeft2 + " characters over the maximum length; please reduce";
                message2Div.innerHTML="<font color=red>" + txt + "</font>";
            }
            }
            catch ( e )
            {
            }
    }

// -->
</script>

</head>
  
<body><!--#include file="top_code.htm"-->

<p style="text-align: center; margin: 24px 0 18px;">
<font color="#47784D"><span style="font-size: 18px"><b>Match Report</b></span></font></p>

<%
Dim output, error, oldcode, newcode, initials, matchdate, headline, report, oldreportind, reportpart, reportparts, acknowledge

output = ""
error = 0

if Request.Form("matchdate") > "" then
	matchdate = Request.Form("matchdate")
  else
   	matchdate = Request.Querystring("date")
end if

if Request.Form("code") > "" then
	oldcode = Request.Form("code")
  elseif Request.Querystring("code") > "" then
	oldcode = Request.Querystring("code")
  else
   	oldcode = Session("Reporter")
end if

acknowledge = Request.Querystring("acknowledge")						'Comes from the url when preceded by the code/date screen ...
if acknowledge = "" then acknowledge = Request.Form("acknowledge")	' ... but from a form when the text is being amended

oldreportind = Request.Querystring("oldreportind")						'Comes from the url when preceded by the code/date screen ...
if oldreportind = "" then oldreportind = Request.Form("oldreportind")	' ... but from a form when the text is being amended


output = output & "<div style=""width:700; margin:12 auto;"">"

  	select case oldcode
  		case "PL82pb:ML"
			initials="ML"
		case "PL71qq:MT"
			initials="MT"
		case "RH150rq:AH"
			initials="AH"
		case "PL65au:SD"
			initials="SD"
		case "PL47be:GB"
			initials="GB"
		case "PL51lt:BW"
			initials="BW"
		case "LS131dt:AC"
			initials="AC"
		case "PL68br:JE"
			initials="JE"
		case else 
			error = 1
	end select
	
	if error = 0 then
	
		if isdate(matchdate) then
		
			newcode = matchdate & ":" & initials
		
			Dim conn, sql, rs
			Set conn = Server.CreateObject("ADODB.Connection")
			Set rs = Server.CreateObject("ADODB.Recordset")

			%><!--#include file="conn_read.inc"--><%
		
			sql = "select date, headline, report "
			sql = sql & "from match_extra " 
			sql = sql & "where date = '" & matchdate & "'"
		
			rs.open sql,conn,1,2
			
			report = ""

			if rs.RecordCount > 0 then
				headline = rs.Fields("headline")
				report = rs.Fields("report")
				if not isnull(report) then report = replace(report,"&pound;","£")						'convert &pound; to £
			end if
			rs.close

			conn.close
			
			output = output & "<div style=""width:800; margin:0 auto; text-align:left"">"
			output = output & "<form action=""matchreport1_action.asp"" method=""post"" name=""form2"">"
			output = output & "<input type=""text"" name=""headline"" placeholder=""Headline"" value=""" & headline & """ size=""50""></p>"
 
			output = output & "<textarea style=""margin: 9px 0 9px 0; padding: 5px"" name=""report"" cols=85 rows=18 placeholder=""Report"" onkeyup=""countChar(report,message1,message2);"" onkeydown=""if(event.keyCode == 13){document.getElementById('btnSendTextMessage').click();}"">"
			if report > "" then
				reportparts = split(report,"|p|")
				for each reportpart in reportparts
					if trim(reportpart) > "" then output = output & trim(reportpart) & vbCrLf & vbCrLf
				next	
				output = left(output,len(output)-4)	'Remove final skip to new lines
			end if
			output = output & "</textarea>"
			output = output & "<p style=""margin: 9px 0"">Ideal report length: 750 - 1,500 characters; current count: <span id=""message1"">" & len(report) & "</span></p>"
			output = output & "<div class=""style1"" id=""message2"" style=""text-align: left; margin: 6px 0""></div>"
			output = output & "<input type=""hidden"" value=""" & oldcode & """ name=""code1"">"
			output = output & "<input type=""hidden"" value=""" & newcode & """ name=""code2"">"
			output = output & "<input type=""hidden"" value=""" & oldreportind & """ name=""oldreportind"">"	
			output = output & "<input type=""hidden"" value=""" & acknowledge & """ name=""acknowledge"">"
			output = output & "<p style=""margin:18px auto;""><input type=""submit"" name=""b2"" value=""Submit Report"" style=""font-size: 12px; padding:2px 5px;""></p>"
			output = output & "</form>"		
			output = output & "</div>"
					
		  else error = 2
		end if
		
	end if	

	output = output & "</div>"

select case error
	case 1
		response.write("<p>Incorrect format for the code, please start again</p>")
	case 2
		response.write("<p>Invalid match date passed</p>")
	case else
		response.write(output)
end select

%>

<!--#include file="base_code.htm"-->
</body>

</html>