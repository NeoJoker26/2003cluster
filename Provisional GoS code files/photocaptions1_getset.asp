<%@ Language=VBScript %>
<% Option Explicit %>

<%
Dim eventdate, settype

'Get parameters passed to calling page
eventdate = left(Request.Form("code"),10)
settype = Request.Form("type")

'Check parameter lengths to prevent SQL injection   
if len(eventdate) = 10 and len(settype) = 1 then 


	Dim conn, sql, rs
	Set conn = Server.CreateObject("ADODB.Connection")
	Set rs = Server.CreateObject("ADODB.Recordset")

	%><!--#include file="conn_update.inc"--><%

	sql = "select photo_set_next " 
	sql = sql & "from photo_set_control " 
	sql = sql & "where date = '" & eventdate & "' "	
	sql = sql & "  and type = '" & settype & "' "		
			
	rs.open sql,conn,1,2
	
	if rs.RecordCount = 0 then
		
		sql = "insert into photo_set_control "
		sql = sql & "values("
		sql = sql & "'" & eventdate & "', "	
		sql = sql & "'" & settype & "', "	
		sql = sql & "2) "
		
		On Error Resume Next
		conn.Execute sql
		if err <> 0 then 
			response.write("<p>SQL ERROR! Statement: " & sql & "  Error: " & err.description & "</p>")
		  else
		  	On Error GoTo 0
		  	Session("thisset") = 1
		end if
		
	  else
	  
		Session("thisset") = rs.Fields("photo_set_next")			

		rs.close

		sql = "update photo_set_control "
		sql = sql & "set photo_set_next = photo_set_next + 1 "
		sql = sql & "where date = '" & eventdate & "' "	
		sql = sql & "  and type = '" & settype & "' "		
	
		on error resume next
		conn.Execute sql
		if err <> 0 then response.write("<p>SQL ERROR! Statement: " & sql & "  Error: " & err.description & "</p>")
	
	end if
	
	 else 
	 
	 	Session("thisset") = 9		'Shouldn't happen, but better to have an extreme number than none at all
	
end if  

%>