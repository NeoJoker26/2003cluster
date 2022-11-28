
<%@ Language=VBScript %>
<% Option Explicit %>

<%
Dim objXML, resXML, xmlHTTP, OK, n1, n2, results(500), resultsindex(500,2), total, urlpart, urlpart2, page, page1, searchstring, sortstring, title, resultprefix
Dim resultentries(500), resultentry, resulttext, hitcount, hitsplit, searchwords, output, displaychoices, displaychoice, resultcounts(20,1)
Dim monthhold, team, teamcode, season, matchpart1, matchpart2, seasonpart1, seasonpart2, homeaway, checkboxoutput, checkvar, errormessage
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<html>
<head>
<meta http-equiv="Content-Language" content="en-gb">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
<title>Greens on Screen</title>
<link href="gos2.css" rel="stylesheet" />

</head>
<body><!--#include file="top_code.htm"-->
<div id="search">

<%

Server.ScriptTimeout = 30 

' Initialise array 
for n1 = 0 to UBound(resultcounts,1)
	resultcounts(n1,1) = 0   'initialise each count
next

errormessage = ""

' Test whether this is the first time in, or whether results are being redisplayed
if Request.Form("result1") = "" then

'*** New search to be processed 

searchstring = Request.Form("searchin") 
searchstring = Replace(searchstring,"""","") 'remove double danglers - they cause problems

Set xmlHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
Set objXML = CreateObject("MSXML.DOMDocument")
objXML.Async=False
 
'open the GET connection to the web server
xmlHTTP.Open "GET", "http://www.jrank.org/api/search/v2.xml?key=22f4d3d851d20885038605162cc03cda56e4517e&q=" & searchstring & "&limit=1000", False
'establish the connection
xmlHTTP.Send

'receive the response 
OK = objXML.load(xmlHTTP.responseXML)

if OK then
	
' Get match details from Search Results XML feed

set	resXML = objXML.getElementsByTagName("meta")

total = resXML(0).childnodes(12).text 	'<end> 

searchwords = resXML(0).childnodes(6).text

if total > 0 and total < 500 then

 set	resXML = objXML.getElementsByTagName("entry")
 hitcount = 0

 for n1 = 0 to total-1

 	hitsplit = split(resXML(n1).childnodes(2).text,"<strong>")
 	hitcount = hitcount + UBound(hitsplit)

	'Find page (last part of url)

	urlpart = split(resXML(n1).childnodes(0).text,"/")
	page = urlpart(Ubound(urlpart))
	page1 = Replace(page,"sv-","")
	page1 = Replace(page1,".asp","")
 	title = Trim(Replace(resXML(n1).childnodes(1).text,"Greens on Screen",""))
	
	resultprefix = ""
	if len(title) <= 15 then resultprefix = left(title,15) & string(15-Len(title), " ")
		
	Select case resultprefix
  		case "Daily Diary    "
  				' Adjust for DIARY page
				if left(page,21) = "sv-diary.asp?archive=" then
					monthhold = Month("1 " & Mid(page,22,3))
					if len(monthhold) = 1 then monthhold = "0" & monthhold
					sortstring = "yyydiary" & Mid(page,25,2) & monthhold 
					title = title & ": " & Ucase(Mid(page,22,1)) & Mid(page,23,2) & " 20" & Mid(page,25,2)
				ElseIf left(page,121) = "sv-diary.asp" then
					sortstring = "yyydiary99"
					title = title & ": Current"
				end if
		case "Match Page     "
				' Adjust for MATCH page
				if instr(page,"?team=") > 0 then	' gets over abandoned games, e.g. barnsley09H_abandoned.asp
					matchpart1 = split(page,"?team=")
					matchpart2 = split(matchpart1(1),"&code=")
					seasonpart1 = Left(Right(matchpart2(1),3),2)
					seasonpart2 = seasonpart1 + 1
					if len(seasonpart2) = 1 then seasonpart2 = "0" & seasonpart2
					season = seasonpart1 & "-" & seasonpart2
					homeaway = Right(matchpart2(1),1)
					title = title & ": " & season & " " & Ucase(Mid(matchpart2(0),1,1)) & Mid(matchpart2(0),2) & "(" & homeaway & ")"  
					sortstring = "xxxmatch" & Right(page,3)
				end if
	    case "Memorable Match" 
				sortstring = "wwwmatch" & Right(page1,3)
				resultprefix = "Match Details  "
	    case "Opposition     " 
				sortstring = "vvv" & page
		case "Players        " 
				sortstring = "uuu" & page
		case "Results        "
				title = title & ": " & page1  
				sortstring = "ttt" & page
				resultprefix = "Results&Tables "
		case "Table          " 
				title = title & ": " & Mid(page1,6) 
				sortstring = "sss" & page
		case "Team Photo     " 
				title = title & ": " & Replace(page1,"pic","")  
				sortstring = "rrr" & page
		case "Tour           " 
				sortstring = "qqq" & page
		case Else
		     title = "Miscellaneous: " & page1 
		     resultprefix = "Miscellaneous  "
		     sortstring = "aaa" & page
			 if page = "" then 
			   sortstring = "zzz"
			   title = "Home Page"
			 end if  
	End Select	
	resultsindex(n1,0) = sortstring
	resultsindex(n1,1) = n1
	resulttext = resXML(n1).childnodes(2).text
	results(n1) = resultprefix & "<p class=""result1""><a href=""" & resXML(n1).childnodes(0).text & """>" & title & "</a></p>"
	results(n1) = results(n1) & "<p class=""result2"">" & resulttext & "</p>"
	results(n1) = results(n1) & "<p class=""result3"">" & resXML(n1).childnodes(0).text & " " & sortstring & "</p>"
 next

 Call QuickSort(resultsindex,0,total-1,0)

 n2 = 0
 for n1 = total-1 to 0 step -1
 	resultentries(n2) = results(resultsindex(n1,1))
 	n2 = n2 + 1 
 next

 Call DisplayResults 
 
 else

 	if total = 0 then 
 		errormessage = "<p class=""result2""><font color=""red"">No words found ... try an alternative search</font></p>"
	  else
 	 	errormessage = "<p class=""result2""><font color=""red"">Too many entries ... try an alternative search</font></p>"
 	 end if

  	Call DisplayResults 

end if

else

	response.write("Error reading results file") 
		 
end if

else

'*** Existing results to be redisplayed

searchwords = Request.Form("searchwords")
total = Request.Form("total")
hitcount = Request.Form("hitcount")
n2 = 0
Do While Request.Form("result" & n2) > ""
resultentries(n2) = Request.Form("result" & n2)
n2 = n2 + 1 
loop
 
Call DisplayResults()

end if

response.write(output)
			
%><%'="a" %><%
%>
<center>
<p></p>
</div>

<!--#include file="base_code.htm"-->

</div>

</body>
</html>


<%
Sub DisplayResults()

 output = "<table width=""900px"" border=""0"" cellpadding=""10"" cellspacing=""0"" style=""border-collapse: collapse"">"
 output = output & "<tr>"
 output = output & "<td width=""370px"" valign=""top"">"
 output = output & "<p class=""style1green"" style=""margin: 0 0 3 0""><b><font style=""font-size: 16px"">Search Results</font></b><a href=""http://www.jrank.org/""><img border=""0"" src=""images/jrank.gif"" align=""top"" style=""margin: 0 0 0 10""></a></p>" 
 output = output & "<p class=""style1"" style=""margin: 0 40 3 0""><b>How does it work?</b> All search words appear in each of the following pages. Click on the title to go to that page; use your back button to return.</p>" 
 output = output & "<p class=""style1"" style=""margin: 0 40 3 0"">The results are listed in groups (for instance, Daily Diary), and then ordered by date (when possible, the most recent first). The results are not ranked by word frequency or content relevance.</p>" 

 output = output & "</td>"
 
 for each resultentry in resultentries
   	'Find count position and increment 
   	n1 = 0
   	For n1 = 0 to 20
	   if resultcounts(n1,0) = left(resultentry,15) then 
	   	resultcounts(n1,1) = resultcounts(n1,1) + 1  'increment
	    exit for
	   end if
	   if resultcounts(n1,0) = "" then 
	   	resultcounts(n1,0) = left(resultentry,15)  'set value
	   	resultcounts(n1,1) = resultcounts(n1,1) + 1  'increment
	    exit for
	   end if
    next
 next
 
 displaychoices = Split(Request.Form("choice"),",")
  
 n1 = 0
 n2 = 0
 checkboxoutput = ""
 Do while resultcounts(n1,0) > ""
 	if resultcounts(n1,1) > 0 then 
   	  checkboxoutput = checkboxoutput & "<input type=""checkbox"" name=""choice"" value=""" & resultcounts(n1,0) & """ class=""style3"""
	  if Request.Form("result1") = "" then  'First time display
		checkvar = " checked" 
	   else
		checkvar = ""
		for each displaychoice in displaychoices
	      if trim(resultcounts(n1,0))= trim(displaychoice) then checkvar = " checked"  
	    next
	  end if 
	  checkboxoutput = checkboxoutput & checkvar & ">" & resultcounts(n1,0) & " (" &resultcounts(n1,1) & ")<br>" 
	  n2 = n2 + 1	
    end if
    n1 = n1 + 1
 loop

 output = output & "<td width=""290px"" valign=""top"">"
 if n2 > 1 then
 	output = output & "<p class=""style1""><b>Uncheck to reduce results:</b>"
 	output = output & "<form action=""searchresults.asp"" method=""post"" name=""Form2"">"
 	output = output & checkboxoutput
 	output = output & "<input type=""hidden"" name=""searchwords"" value=""" & searchwords & """ />"
 	output = output & "<input type=""hidden"" name=""total"" value=""" & total & """ />"
 	output = output & "<input type=""hidden"" name=""hitcount"" value=""" & hitcount & """ />"
 	n2 = 0
 	for each resultentry in resultentries
 		resulttext = Replace(resultentry,"""","&quot;")
		output = output & "<input type=""hidden"" name=""result" & n2 & """ value=""" & resulttext & """ />"
		n2 = n2 + 1 
 	next
 	output = output & "<input type=""submit"" name=""b2"" value=""Redisplay Results"" style=""margin:0; font-size: 9px; font-family:Verdana"">"
 	output = output & "</form>"
 end if	
 output = output & "</td>"

 output = output & "<td width=""240px"" valign=""top"">"
 output = output & "<p class=""style1"" style=""margin: 0 0 6 0""><b>Word(s) searched: </b>" & searchwords & "</p>"
 output = output & "<p class=""style1"" style=""margin: 0 0 6 0""><b>Pages found: </b>" & total & "</p>"
 output = output & "<p class=""style1"" style=""margin: 0 0 6 0""><b>Word hits: </b>" & hitcount & "</p>"
 output = output & "<p class=""style1bold"" style=""margin: 0 0 3 0"">Start new search:</p>"
 output = output & "<form action=""searchresults.asp"" method=""post"" name=""Form3"">"
 output = output & "<input name=""searchin"" size=""20"" class=""style3"">"
 output = output & "<input type=""submit"" name=""b3"" value=""Find"" style=""margin: 0 0 0 0; font-size: 9px; font-family:Verdana"">"
 output = output & "</form>"
 output = output & "<p class=""style3"" style=""margin: 0 0 0 0"">Hints: use very specific word(s); avoid unnecessary ones. Response times depend on the number of page links returned.</td>"
 
 output = output & "</tr></table>"
 

 output = output & "<div style=""width: 850; margin-right: auto; margin-left: auto; text-align: left;"">" & "<center>" & errormessage & "</center>"  
 for each resultentry in resultentries
 	if Request.Form("result1") = "" then  'First time display
		output = output & mid(resultentry,16)
	else
 		for each displaychoice in displaychoices
 			if trim(left(resultentry,15)) = trim(displaychoice) then
 				output = output & mid(resultentry,16)
 				exit for
 			end if
 		next
 	end if			
 next
 output = output & "</div>"

'*** End of Display Results 

End Sub 'DisplayResults



Sub SwapRows(ary,row1,row2)
  '== This proc swaps two rows of an array 
  Dim x,tempvar
  For x = 0 to Ubound(ary,2)
    tempvar = ary(row1,x)    
    ary(row1,x) = ary(row2,x)
    ary(row2,x) = tempvar
  Next
End Sub  'SwapRows



Sub QuickSort(vec,loBound,hiBound,SortField)

  '==--------------------------------------------------------==
  '== Sort a 2 dimensional array on SortField                ==
  '==                                                        ==
  '== This procedure is adapted from the algorithm given in: ==
  '==    ~ Data Abstractions & Structures using C++ by ~     ==
  '==    ~ Mark Headington and David Riley, pg. 586    ~     ==
  '== Quicksort is the fastest array sorting routine for     ==
  '== unordered arrays.  Its big O is  n log n               ==
  '==                                                        ==
  '== Parameters:                                            ==
  '== vec       - array to be sorted                         ==
  '== SortField - The field to sort on (2nd dimension value) ==
  '== loBound and hiBound are simply the upper and lower     ==
  '==   bounds of the array's 1st dimension.  It's probably  ==
  '==   easiest to use the LBound and UBound functions to    ==
  '==   set these.                                           ==
  '==--------------------------------------------------------==

  Dim pivot(),loSwap,hiSwap,temp,counter
  Redim pivot (Ubound(vec,2))

  '== Two items to sort
  if hiBound - loBound = 1 then
    if vec(loBound,SortField) > vec(hiBound,SortField) then Call SwapRows(vec,hiBound,loBound)
  End If

  '== Three or more items to sort
  
  For counter = 0 to Ubound(vec,2)
    pivot(counter) = vec(int((loBound + hiBound) / 2),counter)
    vec(int((loBound + hiBound) / 2),counter) = vec(loBound,counter)
    vec(loBound,counter) = pivot(counter)
  Next

  loSwap = loBound + 1
  hiSwap = hiBound
  
  do
    '== Find the right loSwap
    while loSwap < hiSwap and vec(loSwap,SortField) <= pivot(SortField)
      loSwap = loSwap + 1
    wend
    '== Find the right hiSwap
    while vec(hiSwap,SortField) > pivot(SortField)
      hiSwap = hiSwap - 1
    wend
    '== Swap values if loSwap is less then hiSwap
    if loSwap < hiSwap then Call SwapRows(vec,loSwap,hiSwap)


  loop while loSwap < hiSwap
  
  For counter = 0 to Ubound(vec,2)
    vec(loBound,counter) = vec(hiSwap,counter)
    vec(hiSwap,counter) = pivot(counter)
  Next
    
  '== Recursively call function .. the beauty of Quicksort
    '== 2 or more items in first section
    if loBound < (hiSwap - 1) then Call QuickSort(vec,loBound,hiSwap-1,SortField)
    '== 2 or more items in second section
    if hiSwap + 1 < hibound then Call QuickSort(vec,hiSwap+1,hiBound,SortField)

End Sub  'QuickSort
%>