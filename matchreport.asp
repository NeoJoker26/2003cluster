<%@ Language=VBScript %>
<% Option Explicit %>

<html>

<head>
<meta http-equiv="Content-Language" content="en-gb">
<title>Greens on Screen</title>

<link rel="stylesheet" type="text/css" href="gos2.css">
<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">

<style>
<!--
.centre {text-align: center}
td {padding: 1px 4px; vertical-align: middle;}
th {padding: 3px 4px; vertical-align: middle;}
input[type=radio] {margin: 1px 0 1px 0; padding: 0;}
-->
</style>

<script src="https://code.jquery.com/jquery-1.12.4.js"></script>
<script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>

<script>
        $(function() {
            $( "#datepicker" ).datepicker({
				dateFormat: "yy-mm-dd",
				changeMonth:true,
				changeYear:true,
				showMonthAfterYear:true,
				yearRange:"1903:+0"
            });
         });
</script>

</head>
  
<body><!--#include file="top_code.htm"-->

<p style="text-align: center; margin: 24px 0 18px;">
<font color="#47784D"><span style="font-size: 18px"><b>Match Report</b></span></font></p>

<%
Dim output, error, code, initials, i, j, k, x, y 
Dim lowyear, lowlist(40,3), highlist(40,3), tablerows, diff, lastmonth1, lastmonth2, checked1, checked2, checked3, checked4, cellstyle, reportdate, rowcount, source, acknowledge

output = ""
error = 0
code = Request.Form("code")
source = Request.Form("source")
reportdate = Request.Form("reportdate")

if code = "" then

	output = output & "<div style=""width:300; margin:0 auto;"">"
	output = output & "<form action=""matchreport.asp"" method=""post"" name=""form1"">"
	output = output & "<p style=""margin: 12px auto;"">Code: <input type=""text"" name=""code"" size=""11""></p>"
	output = output & "<p style=""margin: 12px auto;"">Source: <input type=""text"" name=""source"" size=""10""></p>"
	output = output & "<p>Date: &nbsp;<input type=""text"" id=""datepicker"" name=""reportdate"" size=""11""></p></div>"
	output = output & "<p style=""margin:200px auto 100px;""><input type=""submit"" name=""b1"" value=""Next"" style=""font-size: 12px; padding: 5px 2px;""></p>"
	output = output & "</form>"
	output = output & "</div>"

  else

  	select case code
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
	
	acknowledge = ""
	if source > "" then 
		select case source
			case "AR"
				acknowledge="A"
			case "HL"
				acknowledge="H"
			case "PA"
				acknowledge="P"
			case else 
				error = 2
		end select
	end if

	
	if error = 0 then
	
		Session("Reporter") = code
		
		Dim conn, sql, rs
		Set conn = Server.CreateObject("ADODB.Connection")
		Set rs = Server.CreateObject("ADODB.Recordset")

		%><!--#include file="conn_read.inc"--><%
		
		if reportdate = "" then
				
				output = output & "<form action=""matchreport_action.asp"" method=""post"" name=""form2"">"
			
				output = output & "<table border=""1"" style=""border-collapse: collapse; margin: 9 auto;"" width=""700px"">"
  
  				
				sql = "select date, name_then_short, homeaway, reporter "
				sql = sql & "from season_this a join opposition b on a.opposition = b.name_then " 
				sql = sql & "order by date "
		
				rs.open sql,conn,1,2
		
				i = 0
				j = 0
				lowyear = year(rs.Fields("date"))
				output = output & "<th colspan= ""2"">" & lowyear & "</th><th>Opposition</th><th>H/A</th><th>?</th><th>ML</th><th>MT</th><th>SD</th><th style=""border-width:0""></th>"
				output = output & "<th colspan= ""2"">" & lowyear+1 & "</th><th>Opposition</th><th>H/A</th><th>?</th><th>ML</th><th>MT</th><th>SD</th>" 
			
   				Do While Not rs.EOF
 
					if year(rs.Fields("date")) = lowyear then
						lowlist(i,0) = rs.Fields("date")
						lowlist(i,1) = rs.Fields("name_then_short")
						lowlist(i,2) = rs.Fields("homeaway")
  						lowlist(i,3) = rs.Fields("reporter")
						i = i + 1
			  		else
						highlist(j,0) = rs.Fields("date")
						highlist(j,1) = rs.Fields("name_then_short")
						highlist(j,2) = rs.Fields("homeaway")
  						highlist(j,3) = rs.Fields("reporter")
						j = j + 1
					end if	
			
					tablerows = i-1
			
					if i < j then 
						tablerows = j 
						diff = j-i
						for k = 1 to diff
							lowlist(i+k,0) = ""
							lowlist(i+k,1) = ""
							lowlist(i+k,2) = ""
							lowlist(i+k,3) = ""	
						next 
					end if 
					
					if j < i then
						diff = i-j
						for k = 1 to diff
							highlist(j+k,0) = ""
							highlist(j+k,1) = ""
							highlist(j+k,2) = ""
							highlist(j+k,3) = ""	
						next 
					end if

					rs.MoveNext
		
				Loop
				rs.close
		
				x = 1
				y = i+1
		
				for i = 0 to tablerows
		
  					output = output & "<tr>"  
  			
  					if lowlist(i,0) = "" then
  				
  						output = output & "<td colspan=""8""></td>"
  			  
  			  		else
  			  		
   						select case lowlist(i,3)
   							case "ML"
   								checked1 = "" : checked2 = "checked" : checked3 = "" : checked4 = "" 
   							case "MT"
   								checked1 = "" : checked2 = "" : checked3 = "checked" : checked4 = ""
   							case "SD"
   								checked1 = "" : checked2 = "" : checked3 = "" : checked4 = "checked"
   							case else
   								checked1 = "checked" : checked2 = "" : checked3 = "" : checked4 = ""
   						end select
				
						cellstyle = ""
						if CDate(lowlist(i,0)) < CDate(Date) then cellstyle = " style=""background-color:#f0f0f0"""
				
						if month(lowlist(i,0)) = lastmonth1 then
							output = output & "<td" & cellstyle & "></td>"
				 		 else			
							output = output & "<td" & cellstyle & ">" & monthname(month(lowlist(i,0)),true) & "</td>"
							lastmonth1 = month(lowlist(i,0))
						end if
						output = output & "<td" & cellstyle & ">" & weekdayname(weekday(lowlist(i,0)),true) & " " & day(lowlist(i,0)) & "</td>"
						output = output & "<td" & cellstyle & "><a href=""matchreport1.asp?date=" & lowlist(i,0) & """>" & lowlist(i,1) & "</a></td>"
						output = output & "<td" & cellstyle & " class=""centre"">" & lowlist(i,2) & "</td>"
						output = output & "<td" & cellstyle & " class=""centre""><input type=""radio"" value=""? "" " & checked1 & " name=""willbe" & x & """></td>"
						output = output & "<td" & cellstyle & " class=""centre""><input type=""radio"" value=""ML"" " & checked2 & " name=""willbe" & x & """></td>"
						output = output & "<td" & cellstyle & " class=""centre""><input type=""radio"" value=""MT"" " & checked3 & " name=""willbe" & x & """></td>"
						output = output & "<td" & cellstyle & " class=""centre""><input type=""radio"" value=""SD"" " & checked4 & " name=""willbe" & x & """></td>"
						output = output & "<input type=""hidden"" value=""" & lowlist(i,3) & """ name=""was" & x & """>"
						output = output & "<input type=""hidden"" value=""" & lowlist(i,0) & """ name=""date" & x & """>"
		
					end if
				
					output = output & "<td style=""border-width:0""></td>"
					x = x + 1
				
					if highlist(i,0) = "" then
  				
  						output = output & "<td colspan=""8""></td>"
  			  
  			  			else	
		
   						select case highlist(i,3)
   							case "ML"
   								checked1 = "" : checked2 = "checked" : checked3 = "" : checked4 = "" 
   							case "MT"
   								checked1 = "" : checked2 = "" : checked3 = "checked" : checked4 = ""
   							case "SD"
   								checked1 = "" : checked2 = "" : checked3 = "" : checked4 = "checked"
   							case else
   								checked1 = "checked" : checked2 = "" : checked3 = "" : checked4 = ""
   						end select
				
						cellstyle = ""
						if CDate(highlist(i,0)) < CDate(Date) then cellstyle = " style=""background-color:#f0f0f0"""	
				
						if month(highlist(i,0)) = lastmonth2 then
							output = output & "<td" & cellstyle & "></td>"
				  		else			
							output = output & "<td" & cellstyle & ">" & monthname(month(highlist(i,0)),true) & "</td>"
							lastmonth2 = month(highlist(i,0))
						end if	
						output = output & "<td" & cellstyle & ">" & weekdayname(weekday(highlist(i,0)),true) & " " & day(highlist(i,0)) & "</td>"
						output = output & "<td" & cellstyle & "><a href=""matchreport1.asp?date=" & highlist(i,0) & """>" & highlist(i,1) & "</a></td>"
						output = output & "<td" & cellstyle & " class=""centre"">" & highlist(i,2) & "</td>"
						output = output & "<td" & cellstyle & " class=""centre""><input type=""radio"" value=""? "" " & checked1 & " name=""willbe" & y & """></td>"
						output = output & "<td" & cellstyle & " class=""centre""><input type=""radio"" value=""ML"" " & checked2 & " name=""willbe" & y & """></td>"
						output = output & "<td" & cellstyle & " class=""centre""><input type=""radio"" value=""MT"" " & checked3 & " name=""willbe" & y & """></td>"
						output = output & "<td" & cellstyle & " class=""centre""><input type=""radio"" value=""SD"" " & checked4 & " name=""willbe" & y & """></td>"
						output = output & "<input type=""hidden"" value=""" & highlist(i,3) & """ name=""was" & y & """>"				
						output = output & "<input type=""hidden"" value=""" & highlist(i,0) & """ name=""date" & y & """>"
				
					end if	
			
					y = y + 1

					output = output & "</tr>"
	
				next
			
				output = output & "</table>"
				output = output & "<input type=""hidden"" name=""initials"" value=""" & initials & """>"
				output = output & "<input type=""hidden"" name=""matchcount"" value=""" & y-1 & """>"
				output = output & "<p style=""margin:18 auto 24;""><input type=""submit"" name=""b2"" value=""Amend Schedule"" style=""width: 120; font-size: 12px; margin-left:0; margin-right:0; padding:0;""></p>"
				output = output & "</form>"
				
			else
	
				'Date supplied, so it must be a report for an old match
				
				if isdate(reportdate) then
				
					sql = "select count(*) as count "
					sql = sql & "from match " 
					sql = sql & "where date = '" & reportdate & "' "
		
					rs.open sql,conn,1,2
					rowcount = rs.Fields("count")
					rs.close

					if rowcount = 1 then
					
						'This is a report for an old date - check for Alec or John - if so, acknowledge
								 
						select case initials
							case "AH"
								acknowledge="L"
							case "JE"
								acknowledge="J"
						end select
											
						response.redirect("matchreport1.asp?date=" & reportdate & "&oldreportind=y&acknowledge=" & acknowledge)
					
					  else
					  
						output = output & "<p style=""margin:24px auto 300px;"">No match was played on this date</p>"
						
					end if
	
				  else
	
					output = output & "<p style=""margin:24px auto 300px;"">Invalid date supplied</p>"

				end if
				
			end if
			
	  else
	  
		select case error
			case 1
				output = output & "<p style=""margin:24px auto 300px;"">Invalid code supplied</p>"		
			case 2
				output = output & "<p style=""margin:24px auto 300px;"">Invalid source supplied</p>"
		end select
			
	end if
	
end if

response.write(output)
	
%>	

<!--#include file="base_code.htm"-->
</body>

</html>