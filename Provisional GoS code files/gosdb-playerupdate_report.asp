<%@ Language=VBScript %> 
<% Option Explicit %>
<%

' This script is run weekly by a Plesk trigger, to send an email about recently updated player pen-pictures

Dim conn,sql,rs,count
Dim strTo,strFrom,strCc,strBcc,message,subject

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
%>
<!--#include file="conn_read.inc"-->
<%  
sql = "select player_id_spell1, surname, initials, penpic_date "
sql = sql & "			   from player "
sql = sql & "where penpic_date >= dateadd(dd,-7,cast(getdate() as date)) "
sql = sql & "order by 4,2,3 "

count = 0

rs.open sql,conn,1,2

message = "<p>Hi John,<p>Here are the latest pen-picture updates:" 	
	
	Do While Not rs.EOF
		
		count = count + 1
		message = message & "<p>" & rs.Fields("penpic_date") & ": " & "<a href=http://greensonscreen.co.uk/gosdb-players2.asp?pid=" & rs.Fields("player_id_spell1") & ">" & trim(rs.Fields("surname")) & ", " & trim(rs.Fields("initials")) & "</a>" 
		
  		rs.MoveNext
	Loop

	rs.close
	conn.close
 
 	if count = 0 then message = "<p>Hi John,<p>There have been no pen-picture updates this week." 
	message = message & "<p>Many thanks,<p>Steve" 
								
strTo = "jeales@sky.com"
strFrom = "GoSDBprofiles@greensonscreen.co.uk" 
strBcc = "steve@greensonscreen.co.uk"	
subject = "Greens on Screen player pen-pic updates"
%><!--#include file="emailcode.asp"-->
