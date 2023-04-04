<%@ Language=VBScript %>
<% Option Explicit %>

<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Greens on Screen</title>
<base target="_self">

<link rel="stylesheet" type="text/css" href="gos2.css">
</head>

<body><!--#include file="top_code.htm"-->


<% Dim i, j, output, thisday, lastday, thismon, lastmon, thisyear, reqday, reqmon, reqyear, yearturn, selday, selmon, pass, code, entriesforemail
Dim fs, f, line, entrydate, diaryday, diaryentries, diaryentry, diaryparas, diarypara, entry_no, para_no, author
Dim conn,sql,rs 
Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
%>

<div style="width: 980px; margin: 0 auto;">
<center>
<h3 style="text-align: center; margin-bottom: 20; margin-top:10">
<font color="#457B44">
<b>Diary Update</b></font></h3>


<% 
pass = Request.QueryString("pass")

select case pass
	case 1
		Call AskDate
	case 2
		Call DisplayDay
	case 3
		Call ProcessDay
	case else
		Call AskCode
end select
%>

<% Sub AskCode %> 
<form action="diaryupdate.asp?pass=1" method="post" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript" name="FrontPage_Form1">
<center>
<table border="0" cellspacing="5" style="border-collapse: collapse; text-align:center" bordercolor="#111111" width="700">
  <tr>
    <td class="style1">
    <p style="margin-top: 0; margin-bottom: 0; text-align:center">Code:&nbsp;&nbsp;&nbsp;&nbsp;
    <input name="code" size="10" maxlength="25" type="password">  
     </tr>
	<tr>
    <td class="style1">
    <p style="text-align: center; margin-bottom:10; margin-top:10">
	<input type="submit" name="b1" value="Validate" style="margin-top: 0; margin-bottom: 0; font-size:10px">&nbsp; <b>
	<font color="#808080">|</font>&nbsp;<a target="_top" href="index.asp">Cancel</a></b></td>
  	</tr>
</table>
</form>
	<% End Sub %>
	
<% Sub AskDate %>
<%
code = Request.Form("code")
if code = "Adam2010" or code = "Sam1903" or code = "1920Andy" or code = "1886Keith" then 
	%> 
	<form action="diaryupdate.asp?pass=2" method="post" onsubmit="return FrontPage_Form2_Validator(this)" language="JavaScript" name="FrontPage_Form2"> 
	<center>
	<table border="0" cellspacing="5" style="border-collapse: collapse; text-align:center" bordercolor="#111111" width="700">
  	<tr>
   	<td class="style1">
   	<p style="margin-top: 0; margin-bottom: 0; text-align:center">Day: <select size="1" name="day">

	<%  
	thismon = month(date)
    thisday = day(date)
	thisyear = year(date)

	if reqday > "" then selday = CInt(reqday) else selday = thisday
	if reqmon > "" then selmon = CInt(reqmon) else selmon = thismon

	output = "" 	
	for i = 1 to 31
 	   output = output & "<option "
 	   if i = selday then output = output & "selected "
 	   output = output & "value=""" & i & """>" & i & "</option>"   
    next
    response.write(output)   
	%>
    </select>
	<%  

	output = ""
	
	yearturn="N"
	i = 5
	Do While i <> thismon
 	 output = output & "<option "
 	 if i = selmon then output = output & "selected "
 	 output = output & "value=""" & i & """>" & MonthName(i,true) & "</option>"
 	 i = i + 1
 	 	if i = 13 then 
 	 		i = 1
 	 		yearturn="Y"
 	 	end if	
 	loop
 	'do final month (this month)
 	output = output & "<option value=""" & i & """"
 	if i = selmon then output = output & " selected"
 	output = output & ">" & MonthName(i,true) & "</option>"

	output = output & "<input type=""hidden"" name=""year"" value=""" & thisyear & """>" 
	output = output & "<input type=""hidden"" name=""yearturn"" value=""" & yearturn & """>" 
	output = output & "<input type=""hidden"" name=""code"" value=""" & code & """>" 
	%>
	&nbsp;  Month: <select size="1" name="month">
	<%	
	response.write(output) 
	%>
	</select> </td>
  	</tr>
	<tr>
    <td class="style1">
    <p style="text-align: center; margin-bottom:10; margin-top:0">
	<input type="submit" name="b2" value="Get Entry" style="margin-top: 0; margin-bottom: 0; font-size:10px">&nbsp;&nbsp;<b><font color="#808080">|</font>&nbsp;<a target="_top" href="index.asp">Cancel</a></b></td>
	</tr>
	</table>
	</form>
 <%
  else
 	output = "<centre><font color=""#FF0000""><b>Invalid Code</b></font></center>"
 	response.write(output)
 end if
 %>
<% End Sub %>

<% Sub DisplayDay %>
<%
reqday = Request.Form("day")
reqmon = Request.Form("month")
reqyear = Request.Form("year")
yearturn = Request.Form("yearturn")
code = Request.Form("code")

if MonthName(reqmon,true) = "Dec" and yearturn = "Y" then reqyear = reqyear-1
entrydate = reqday & " " & MonthName(reqmon,true) & " " & reqyear

if IsDate(entrydate) = 0 then
	response.write("<centre><font color=""#FF0000""><b>Date is not valid - please check</b></font></center>")

  elseif DateValue(entrydate) > Date + 1  then
	response.write("<centre><font color=""#FF0000""><b>Date in the future - please check</b></font></center>")
    
  else
  
	%><!--#include file="conn_read.inc"--><%

	diaryday = ""
	
	sql = "select entry_no, entry_para_no, entry_para "
	sql = sql & "from daily_diary " 
	sql = sql & "where date = '" & entrydate & "' "
	sql = sql & "order by entry_no, entry_para_no "

	rs.open sql,conn,1,2
		
	Do While Not rs.EOF	
			 
		if rs.Fields("entry_no") = 1 and rs.Fields("entry_para_no") = 1 then
			diaryday = diaryday & rs.Fields("entry_para") & Chr(13)&Chr(10)
		  elseif rs.Fields("entry_no") > 1 and rs.Fields("entry_para_no") = 1 then
		  	diaryday = diaryday & ">>>" & rs.Fields("entry_para") & Chr(13)&Chr(10)	
		  else
			diaryday = diaryday & rs.Fields("entry_para") & Chr(13)&Chr(10)
		end if
	
		rs.MoveNext
	Loop
	
	rs.close
	conn.close
	
	if len(diaryday) = 0 then
		response.write("<centre><font style=""font-size: 12px"" color=""green""><b>No entry found; enter new details, then store, or change date ...</b></font></center><br>")
	  else
		response.write("<centre><font style=""font-size: 12px"" color=""#F1A629""><b>Entry already exists; change or add details, then store, or change date ...</b></font></center><br>")
		diaryday = left(diaryday,len(diaryday)-2)	'remove final skip to new line
	end if	  
		
	%><%'="a" %><%
	%>
	<% Call Askdate %>

<center>
	<form action="diaryupdate.asp?pass=3" method="post" onsubmit="return FrontPage_Form3_Validator(this)" language="JavaScript" name="FrontPage_Form3">
	<center>
	<table border="0" cellspacing="5" style="border-collapse: collapse; text-align:center" bordercolor="#111111" width="700">	
	<tr>
    <td style="text-align: left" >
    <p class="style1" style="margin-top: 6px; margin-bottom: 3px">Notes:</p>
    <ol style="margin: 10px 0 20px -20px">
	<li class="style1">Each new diary entry must begin with >>>, except for a day's first. So the first and those starting with >>> will show with a leading green dot.</p>
	<li class="style1">For a new paragraph within an entry, simply skip to a new line.</p>
	<li class="style1">If you forget to begin a new diary entry with >>>, it will appear as another paragraph under the prior entry. To correct this, go back to that date and insert >>> at the appropriate point.</p>
    </ol>
    <p style="text-align: center"><textarea rows="27" name="entry" cols="100" wrap="wrap"><%response.write(diaryday) %></textarea>
    <%response.write("<input type=""hidden"" name=""entrydate"" value=""" & entrydate & """>") 
      response.write("<input type=""hidden"" name=""code"" value=""" & code & """>")%>
    </td>
  	</tr>
  	<tr>
    <td style="text-align: left" >
    <p style="text-align: center" class="style1">
	<input type="submit" name="b3" value="Store Entry" style="font-size: 10px">&nbsp; <b>
	<font color="#808080">|</font>&nbsp;<a target="_top" href="diaryupdate.asp">Cancel</a></b></td>
  	</tr>
	</table>
	</form>
<%
end if
%>	
<% End Sub %>

<% Sub ProcessDay %>
<%


code = Request.Form("code")
entrydate = Request.Form("entrydate")
diaryday = Request.Form("entry")
if left(diaryday,3) = ">>>" then diaryday = mid(diaryday,4)		'No need for a new entry indicator before the first entry (in fact, it causes problems if it's there) 
diaryday = replace(diaryday,"'","''")
diaryday = replace(diaryday,"‘","''")
diaryday = replace(diaryday,"’","''")
diaryday = replace(diaryday,"“","""")
diaryday = replace(diaryday,"”","""")	

entry_no = 1

	%><!--#include file="conn_update.inc"--><%

	sql = "delete from daily_diary "
	sql = sql & "where date = '" & entrydate & "' "
			
	on error resume next
	conn.Execute sql
	if err <> 0 then response.write("<p class=""style1boldred"">SQL ERROR! Statement: " & sql & "  Error: " & err.description & "</p>")
	On Error GoTo 0
	
	entriesforemail = "<p>" & entrydate
	
	diaryentries = split(diaryday,">>>")
	
	for each diaryentry in diaryentries
	
		para_no = 1
	
		diaryparas = split(diaryentry,Chr(13)&Chr(10))
		
			for each diarypara in diaryparas
			
				if diarypara > "" then
				
					entriesforemail = entriesforemail & "<p>" & diarypara
				
					sql = "insert into daily_diary (date, entry_no, entry_para_no, entry_para) "
					sql = sql & "values ("
					sql = sql & "'" & entrydate & "',"
					sql = sql & entry_no & ","
					sql = sql & para_no & ","
					sql = sql & "'" & diarypara & "'" 	
					sql = sql & ")"	
			
					on error resume next
					conn.Execute sql
					if err <> 0 then response.write("<p class=""style1boldred"">SQL ERROR! Statement: " & sql & "  Error: " & err.description & "</p>")
					On Error GoTo 0
			
					para_no = para_no + 1
					
				end if
			next
			
		entry_no = entry_no + 1
	next
	
	conn.close


Dim strTo,strFrom,strCc,strBcc,message,subject
	   								
strTo = "diary_updated@greensonscreen.co.uk"
strFrom = "diary_updated@greensonscreen.co.uk"
strCc = ""
message = "<p>Diary updated by " & code
message = message & "<p><a href=""http://www.greensonscreen.co.uk/sv-diary.asp"">Show Diary</a>" & entriesforemail
subject = "GoS Diary Update"
   		   				
%><!--#include file="emailcode.asp"--><%

output = "<p class=""style1bold"" style=""margin:12px auto;"">The diary entry(ies) have been added.</p>"
output = output & "<p class=""style1bold"" style=""margin: 24px auto; font-size: 12px;""><a target=""_blank"" href=""dailydiary.asp"">Show Diary</a></p>"
response.write(output) 

%>
<% End Sub %>


<%
%><%'="a"%><%
%>

</div>&nbsp;
<!--#include file="base_code.htm"-->
</body></html>