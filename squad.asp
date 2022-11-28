<%@ Language=VBScript %>
<% Option Explicit %>

<!doctype html>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Greens on Screen</title>

<link rel="stylesheet" type="text/css" href="gos2.css">
<style>
<!--
.rowhlt {background-color: #d5e9d7;}
-->
</style>

<script type="text/javascript"  src="jquery/jquery-1.11.1.min.js"></script>
<script>
$(document).ready(function() {
    $("img").on("contextmenu",function(){
       return false;
    }); 
});
</script>

</head>

<body><!--#include file="top_code.htm"-->
 
<div id=squad>
<p style="margin-top: 12; margin-bottom: 3"><font color="#457B44" style="font-size: 18px">PLYMOUTH ARGYLE CURRENT SQUAD</font></p>
<p style="margin: 0" class="style1">Click on
  <img src="images/sort.gif" border="0" hspace="0" align="top"> to re-order; on a name for more details. </p>

<%

Dim conn, rs, sql, output, output1, output2, sortclause, photoname
Dim thhlt, sort

sort = Request.QueryString("sort")
if sort = "" then sort = 2

select case sort
	case 1
		sortclause = "order by surname, forename "
	case 2
		sortclause = "order by squad_no, surname, forename "
	case 3
		sortclause = "order by sortposition, surname, forename "
	case 4
		sortclause = "order by  "
	case 5
		sortclause = "order by  "
	case 6
		sortclause = "order by dob "
	case 7
		sortclause = "order by signed_this_spell "
	case 8
		sortclause = "order by appears desc"
	case 9
		sortclause = "order by goals desc"
end select

output = "<table>"
output = output & "<tr>"
output = output & "<td class=""num""><a href=""squad.asp?sort=2""><img src=""images/sort.gif"" border=""0""></a><br><b>No</b></td>"
output = output & "<td><a href=""squad.asp?sort=1""><img src=""images/sort.gif"" border=""0""></a><br><b>Name</b><br>Click for more details</td>"
output = output & "<td><a href=""squad.asp?sort=3""><img src=""images/sort.gif"" border=""0""></a><br><b>Position</b></td>"
output = output & "<td><a href=""squad.asp?sort=6""><img src=""images/sort.gif"" border=""0""></a><br><b>Born on</b></td>"
output = output & "<td><br><b>Born in</b></td>"
output = output & "<td><br><b>Came From</b></td>"
output = output & "<td><a href=""squad.asp?sort=7""><img src=""images/sort.gif"" border=""0""></a><br><b>Signed</b><br>[this spell]</td>"
output = output & "<td class=""num""><a href=""squad.asp?sort=8""><img src=""images/sort.gif"" border=""0""></a><br><b>Starts<br>(subs)</b></td>"
output = output & "<td class=""num""><a href=""squad.asp?sort=9""><img src=""images/sort.gif"" border=""0""></a><br><b>Goals</b></td>"
output = output & "</tr>"

output2 = "<div id=squadphotos>"


Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_read.inc"--><%

' NB. The first of the 4 unioned selects had a count of zero for each total, without being joined to match_player or match_goal, to ensure
'     that players who have yet to appear will still be picked up. 	
	sql = "with cte as ( "
	sql = sql & "select squad_no, b.dob, b.pob, b.signed_this_spell, b.came_from, b.loan, b.loaned_to, b.position, b.surname, b.forename, rtrim(b.forename) + ' ' + rtrim(b.surname)  as name, b.player_id_spell1, c.photo_exists, c.prime_photo, 0 as starts,0 as subs,0 as goals "
	sql = sql & "from player_squad a join player b on a.player_id = b.player_id join player c on b.player_id_spell1 = c.player_id "
	sql = sql & "where b.last_game_year = 9999 "
	sql = sql & "and season_no = (select max(season_no) from player_squad) "
	sql = sql & "union all "	
	sql = sql & "select squad_no, b.dob, b.pob, b.signed_this_spell, b.came_from, b.loan, b.loaned_to, b.position, b.surname, b.forename, rtrim(b.forename) + ' ' + rtrim(b.surname)  as name, b.player_id_spell1, c.photo_exists, c.prime_photo, 1, 0, 0 "
	sql = sql & "from player_squad a join player b on a.player_id = b.player_id join player c on b.player_id_spell1 = c.player_id join match_player d on d.player_id in (select player_id from player e where e.player_id_spell1 = b.player_id_spell1) "
	sql = sql & "where b.last_game_year = 9999 "
	sql = sql & "and season_no = (select max(season_no) from player_squad) "
	sql = sql & "and startpos > 0 "
	sql = sql & "union all "
	sql = sql & "select squad_no, b.dob, b.pob, b.signed_this_spell, b.came_from, b.loan, b.loaned_to, b.position, b.surname, b.forename, rtrim(b.forename) + ' ' + rtrim(b.surname)  as name, b.player_id_spell1, c.photo_exists, c.prime_photo, 0, 1 ,0 "
	sql = sql & "from player_squad a join player b on a.player_id = b.player_id join player c on b.player_id_spell1 = c.player_id join match_player d on d.player_id in (select player_id from player e where e.player_id_spell1 = b.player_id_spell1) "
	sql = sql & "where b.last_game_year = 9999 "
	sql = sql & "and season_no = (select max(season_no) from player_squad) "
	sql = sql & "and startpos = 0 "
	sql = sql & "union all "
	sql = sql & "select squad_no, b.dob, b.pob, b.signed_this_spell, b.came_from, b.loan, b.loaned_to, b.position, b.surname, b.forename, rtrim(b.forename) + ' ' + rtrim(b.surname)  as name, b.player_id_spell1, c.photo_exists, c.prime_photo, 0, 0, 1 "
	sql = sql & "from player_squad a join player b on a.player_id = b.player_id join player c on b.player_id_spell1 = c.player_id join match_goal d on d.player_id in (select player_id from player e where e.player_id_spell1 = b.player_id_spell1) "
	sql = sql & "where b.last_game_year = 9999 "
	sql = sql & "and season_no = (select max(season_no) from player_squad) "
	sql = sql & ") "
	sql = sql & "select squad_no, dob, pob, signed_this_spell, came_from, loan, loaned_to, "
	sql = sql & "case left(position,3) "
	sql = sql & "  when 'Goa' then '1' + position "
	sql = sql & "  when 'Def' then '2' + position "
	sql = sql & "  when 'Mid' then '3' + position "
	sql = sql & "  when 'For' then '4' + position "
	sql = sql & "  end as sortposition, "
	sql = sql & "surname, forename, name, player_id_spell1, photo_exists, prime_photo, sum(starts) + sum(subs) as appears, sum(starts) as starts, sum(subs) as subs, sum(goals) as goals "	
	sql = sql & "from CTE "
	sql = sql & "group by squad_no, dob, pob, signed_this_spell, came_from, loan, loaned_to, position, surname, forename, name, player_id_spell1, photo_exists, prime_photo  "
	sql = sql & sortclause
	rs.open sql,conn,1,2
	
	Do While Not rs.EOF
		output = output & "<tr onmouseover=""this.className = 'rowhlt';"" onmouseout=""this.className = '';"">"
		output = output & "<td class=""num"">" & rs.Fields("squad_no") & "</td>"
		if not isnull(rs.Fields("player_id_spell1")) then
			output = output & "<td class=""name1"" nowrap><a href=""gosdb-players2.asp?pid=" & rs.Fields("player_id_spell1") & "&from=squad"">" & rs.Fields("name") & "</a>"
		  else
			output = output & "<td class=""name2"" nowrap>" & rs.Fields("name")
		end if
		if not isnull(rs.Fields("loaned_to")) then
			output = output & " *"  
			output1 = output1 & "* " & rs.Fields("name") & " is currently on loan at " & rs.Fields("loaned_to") & "<br>"			
		end if
		output = output & "</td>"
		output = output & "<td>" & mid(rs.Fields("sortposition"),2) & "</td>"
		output = output & "<td>" & rs.Fields("dob") & "</td>"
		output = output & "<td>" & rs.Fields("pob") & "</td>"
		output = output & "<td>" & rs.Fields("came_from")
		if ucase(rs.Fields("loan")) = "Y" then output = output & " on loan"
		output = output & "</td>"
		output = output & "<td>" & rs.Fields("signed_this_spell") & "</td>"
		output = output & "<td class=""num"">" & rs.Fields("starts") & " (" & rs.Fields("subs")  & ")</td>"		
		output = output & "<td class=""num"">" & rs.Fields("goals") & "</td>"
		output = output & "</tr>" 
		
		if isnull(rs.Fields("photo_exists")) then
			photoname = "nophoto"
		  elseif len(rs.Fields("player_id_spell1")) < 4 then 
			photoname = right("00" & rs.Fields("player_id_spell1"),3)
	  	  else
	  		photoname = rs.Fields("player_id_spell1")
	  	end if	
	  	
	  	if not isnull(rs.Fields("photo_exists")) and not isnull(rs.Fields("prime_photo")) then photoname = photoname & "_" & rs.Fields("prime_photo")
	  	
	  	photoname = photoname & ".jpg"
	
		output2 = output2 & "<a href=""gosdb-players2.asp?pid=" & rs.Fields("player_id_spell1") & "&from=squad"">"
		output2 = output2 & "<figure><img width=""192"" src=""gosdb/photos/players/" & photoname & """><figcaption>" & rs.Fields("squad_no") & ". " & rs.Fields("name") & "</figcaption></figure></a>"
		
		rs.MoveNext
	Loop
	
	rs.close

output = output & "<tr><td colspan=""9"">" & output1 & "</td></tr></table>" 
output = output & output2 & "</div>"

response.write(output)
%><%'="a" %><%
%>

</div>

<!--#include file="base_code.htm"-->

</body>

</html>

<%
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