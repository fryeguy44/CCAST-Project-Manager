<%@ Language=VBScript CodePage=65001 %>
<%Response.Charset="UTF-8"%>
<!--#include virtual="/Connections/OBA.asp" -->
<%
' @name	event calendar in classic asp
' @version 	1.0
' @date 	2016/oct/28
' @author	coax
' @website	http://coax.hr/asp/event-calendar
' @license	http://gnu.org/licenses/gpl-3.0.txt
%>
<%
Dim mysql_server, mysql_driver, mysql_db, mysql_user, mysql_pwd

' YOUR OWN CONFIGURATION
'--------------------------------------------------
Session.LCID = 1033 ' Set your Locale ID, we used US locale here
'mysql_server = "localhost" ' Your MySQL server address, keep this if installed on same server
'mysql_driver = "{MySQL ODBC 5.3 Unicode Driver}" ' Replace with your version of installed ODBC driver
'mysql_db = "asp" ' MySQL database name
'mysql_user = "asp" ' MySQL username
'mysql_pwd = "demo" ' MySQL password

' CALENDAR FUNCTIONS
' no need to change anything below this line
'--------------------------------------------------

' Cache
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1

' MySQL database
'--------------------------------------------------
On Error Resume Next
Set Conn = CreateObject("ADODB.Connection") 
'Conn.ConnectionString = "Driver=" & mysql_driver & "; Server=" & mysql_server & "; Database=" & mysql_db & "; User=" & mysql_user & "; Password=" & mysql_pwd & "; Option=3" : Conn.Open
Conn.ConnectionString = MM_OBA_STRING
Conn.Open
Set RS = Server.CreateObject("ADODB.Recordset") : RS.CursorLocation = 3 : RS.CursorType = 3 : RS.LockType = 3
If Err.Number<>0 then
	Response.Write "Error connecting to database."
	Response.End
End if
On Error Goto 0

' QueryStrings for calendar display
'--------------------------------------------------
EventId = Request.QueryString("id")
Action = Request.QueryString("action")
EventPublished = Request.QueryString("start")
EventExpires = Request.QueryString("end")

' Read input fields on form submit
'--------------------------------------------------
If Request.Form<>"" then
	
    EventGroup = CInt(Request.Form("EventGroup"))
	EventTitle = Request.Form("EventTitle")
	EventLong = Request.Form("EventLong")
	EventColor = Request.Form("EventColor")
    Reminder = Request.Form("Reminder")
    RecurringPeriodId = null
    if Request.Form("RecurringPeriodId")<>"" then
        RecurringPeriodId = CInt(Request.Form("RecurringPeriodId"))
    end if
	
    EventPublished = Request.Form("EventPublished") 
    If IsDate(EventPublished) then 
        EventPublished = IsoDate(EventPublished) 
    Else 
        EventPublished = IsoDate(Now())
    End If
	
    EventExpires = Request.Form("EventExpires")
    If IsDate(EventExpires) then 
        EventExpires = IsoDate(EventExpires) 
    Else 
        EventExpires = IsoDate(Now())
    End If
	
    If EventPublished=EventExpires or (Hour(EventPublished) & Minute(EventPublished) = "00" and Hour(EventExpires) & Minute(EventExpires) = "00") then
		EventAllDay = true
        ' remove times
        EventPublished = left(EventPublished, instr(EventPublished, " ")-1)
        EventExpires = left(EventExpires, instr(EventExpires, " ")-1)
	Else
		EventAllDay = false
	End if
End if

' Actions
'--------------------------------------------------
If Action = "1" then Add
If Action = "2" then Edit(EventId)
If Action = "3" then Del(EventId)

' ISO date
'--------------------------------------------------
Function IsoDate(x)
	If IsEmpty(x) then Exit function
	IsoDate = Year(x) & "-" & Right("00" & Month(x),2) & "-" & Right("00" & Day(x),2) & " " & Right("00" & Hour(x),2) & ":" & Right("00" & Minute(x),2)
End function

' Regular Expression Swapping
'--------------------------------------------------
Function Swap(x, y, z)
	If IsNull(x) then Exit function

	Set Re = New RegExp
	Re.Global = True
	Re.IgnoreCase = True
	Re.Pattern = y
	val = Re.Replace(x, z)
	Set Re = Nothing

	Swap = val
End function

' Autoparse URL's
'--------------------------------------------------
Function AutoUrl(x)
	If IsNull(x) then Exit function

	x = " " & x
	x = Swap(x, "(^|[\n ])([\w]+?://[^ ,""\s<]*)", "$1<a href=""$2"">$2</a>")
	x = Swap(x, "(^|[\n ])((www|ftp)\.[^ ,""\s<]*)", "$1<a href=""http://$2"">$2</a>")
	x = Swap(x, "(^|[\n ])([a-z0-9&\-_.]+?)@([\w\-]+\.([\w\-\.]+\.)*[\w]+)", "$1<a href=""mailto:$2@$3"">$2@$3</a>")
	x = Right(x, Len(x)-1)

	AutoUrl = x
End function

Function BoolToBit(val) 
    Dim retVal
    If cbool(val) = True Then
        retVal = 1
    Else
        retVal = 0
    End If
    BoolToBit = retVal
End Function

Function IIf(bClause, sTrue, sFalse)
    If CBool(bClause) Then
        IIf = sTrue
    Else 
        IIf = sFalse
    End If
End Function

' Calendar events
'--------------------------------------------------
If Request.QueryString("_")<>"" then
%>[<%
	SQL = "SELECT EventId, EventTitle, EventLong, EventColor, EventPublished, EventExpires, EventAllDay FROM Events WHERE EventGroup = 1 OR Username = '" & Session("MM_Username") & "' ORDER BY EventId ASC"
	RS.Open SQL, Conn
	If not RS.EOF then Events = RS.GetRows
	RS.Close

	If IsArray(Events) then
		For i=0 to UBound(Events,2)
			EventId = Events(0,i)
			EventTitle = Events(1,i)
			EventLong = Events(2,i)
			EventColor = Events(3,i)
			EventPublished = IsoDate(Events(4,i))
			EventExpires = IsoDate(Events(5,i))
			EventAllDay = Events(6,i)
			If EventAllDay = true then
				EventAllDay = "true"
			Else
				EventAllDay = "false"
			End if
%>{"id": "<%=EventId%>", "title": "<%=EventTitle%>", "allDay": <%=EventAllDay%>, "start": "<%=EventPublished%>", "end": "<%=EventExpires%>", "url": "<%=EventUrl%>", "color": "<%=EventColor%>", "description": ""}<%
			If i<UBound(Events,2) then Response.Write ","
		Next
	End if
%>]<%
	Response.End

ElseIf EventId<>"" and EventPublished<>"" and EventExpires<>"" then
	If (Hour(EventPublished)=0 and Minute(EventPublished)=0) and (Hour(EventPublished)=Hour(EventExpires)) then
		EventAllDay = true
	Else
		EventAllDay = false
	End if

    SQL = "UPDATE Events SET EventPublished='" & EventPublished & "', EventExpires='" & EventExpires & "', EventAllDay=" & BoolToBit(EventAllDay) & " WHERE EventId=" & EventId
	Conn.Execute SQL

	Set RS = Nothing
	Conn.Close : Set Conn = Nothing

	Response.End

ElseIf EventId<>"" then
	SQL = "SELECT EventId, EventGroup, EventTitle, EventLong, EventColor, EventPublished, EventExpires, EventAllDay, Reminder, isNull(RecurringPeriodId,0) as RecurringPeriodId FROM Events WHERE EventId=" & EventId & " ORDER BY EventId ASC"
	RS.Open SQL, Conn
	If not RS.EOF then Events = RS.GetRows
	RS.Close

	If IsArray(Events) then
		For i=0 to UBound(Events,2)
			EventId = Events(0,i)
			EventGroup = Events(1,i)
			EventTitle = Events(2,i)
			EventLong = Events(3,i)
			EventColor = Events(4,i)
			EventPublished = Events(5,i)
			EventExpires = Events(6,i)
			EventAllDay = Events(7,i)
            Reminder = Events(8,i)
            RecurringPeriodId = Events(9,i)
		Next
	End if
	Action = 2

Else
	If IsDate(EventPublished) then
		'If Hour(EventPublished)="0" then
		'	EventPublishedHour = Now()
		'Else
			EventPublishedHour = EventPublished
		'End if
		EventPublished = Year(EventPublished) & "-" & Right("00" & Month(EventPublished),2) & "-" & Right("00" & Day(EventPublished),2) & " " & Right("00" & Hour(EventPublishedHour),2) & ":" & Right("00" & Minute(EventPublishedHour),2)
		EventPublished = CDate(EventPublished)
	End if
	If IsDate(EventExpires) then
		If Hour(EventExpires)="0" then
			EventExpires = DateAdd("d",-1,EventExpires) ' Remove 1 day
		'	EventExpiresHour = Now()
		'Else
			EventExpiresHour = EventExpires
		End if
		EventExpires = Year(EventExpires) & "-" & Right("00" & Month(EventExpires),2) & "-" & Right("00" & Day(EventExpires),2) & " " & Right("00" & Hour(EventExpiresHour),2) & ":" & Right("00" & Minute(EventExpiresHour),2)
		EventExpires = CDate(EventExpires)
	End if
	Action = 1
End if
%>

<% If EventId<>"" and Request.QueryString("edit")="" then %>

  <h2><%=EventTitle%><% If EventGroup=1 then %><br /><span>Business event</span><% End if %></h2>

<% If EventPublished=EventExpires then %>
  <p>All-day event</p>
<% Else %>
  <p>Event start: <%=EventPublished%><br />
    Event end: <%=EventExpires%></p>
<% End if %>
  <p style="border-top:1px solid #dfe0e4; padding-top:15px;"><%=AutoUrl(Replace(EventLong, Chr(10), "<br />" & VbCrLf))%></p>
  <p><button onclick="$.facebox({ajax: 'includes/calendar.asp?id=<%=EventId%>&amp;edit=true'});">Edit</button></p>
  <% Conn.Execute "UPDATE Events SET EventViews=EventViews+1 WHERE EventId=" & EventId %>

<% Else %>

  <form id="form" onsubmit="return Validation('includes/calendar.asp?id=<%=EventId%>&amp;action=<%=Action%>', $(this));">
    <h2><% If Action=1 then %>Add event<% Else %>Edit event<% End if %></h2>
    <p class="selected-date">Selected date: <%=EventPublished%> &ndash; <%=EventExpires%></p>
    <input type="hidden" name="EventPublished" value="<%=EventPublished%>" />
    <input type="hidden" name="EventExpires" value="<%=EventExpires%>" />
    <p>Title</p>
    <input type="text" name="EventTitle" value="<%=EventTitle%>" maxlength="255" placeholder="Event title" class="req">
    <p>Type</p>
    <select name="EventGroup"><%
	Dim Groups(1,1)
	Groups(0,0) = "1"
	Groups(1,0) = "Business"
	Groups(0,1) = "2"
	Groups(1,1) = "Private"

	If IsArray(Groups) then
		For i=0 to UBound(Groups,2)
			GroupId = Groups(0,i)
			GroupTitle = Groups(1,i)
            %><option value="<%=GroupId%>"<% If CInt(EventGroup)=CInt(GroupId) then Response.Write " selected" %>><%=GroupTitle%></option><% 
        Next  
    End if %>
    </select>
    <p>Recurrence</p>
    <select name="RecurringPeriodId">
        <option value=""></option>     
        <%
        dim periodId, periodName
        dim periods : periods = getRecurringPeriodsArray()
        if IsArray(periods) then 
            for i=0 to ubound(periods, 2) 
                periodId = periods(0,i)
                periodName = periods(1,i)
                %><option value="<%=periodId%>"<% If cint(RecurringPeriodId)=cint(periodId) then Response.Write " selected" %>><%=periodName%></option><% 
            next
        end if 
        %>
    </select>
    <p>Description</p>
    <textarea name="EventLong" rows="5" placeholder="Event description" class="req"><%=EventLong%></textarea>
    <p>Send reminder? <input type="checkbox" class="checkbox" name="Reminder" value="true" <% if Reminder = true then response.write "checked" end if %> /></p>
    <p>Color</p>
    <input type="text" name="EventColor" value="<%=EventColor%>" placeholder="Click to change" maxlength="7" style="text-shadow:none;<% If EventColor<>"" then %>background:<%=EventColor%>;<% End if %>">
    <p class="buttons"><button>Save</button><% If EventId<>"" then %><button onclick="return Question('Confirm delete <b><%=EventTitle%></b>:','includes/calendar.asp?id=<%=EventId%>&amp;action=3');">Delete this event</button><% End if %></div>
  </form>

<% End if %>

<%
Sub Add
	If Request.Form="" then Response.Redirect "/"

	Conn.Execute "INSERT INTO Events (EventGroup, EventTitle, EventLong, EventColor, EventViews, EventPublished, EventExpires, EventAllDay, EventDate, Reminder, RecurringPeriodId, Username) VALUES (" & _
        EventGroup & ", '" & EventTitle & "', '" & EventLong & "', '" & EventColor & "', 0, '" & EventPublished & "', '" & EventExpires & "', " & BoolToBit(EventAllDay) & _
        ", getdate(), " & BoolToBit(Reminder) & ", " & iif(isnull(RecurringPeriodId), "null", RecurringPeriodId) & ", '" &  Session("MM_Username") & "' )"

	Set RS = Nothing
	Conn.Close : Set Conn = Nothing

	Response.Write "<p>Event added to calendar.</p>"
	Response.End
End Sub

Sub Edit(EventId)
	If Request.Form="" then Response.Redirect "/"

    SQL = "UPDATE Events SET EventGroup=" & EventGroup & ", EventTitle='" & EventTitle & "', EventLong='" & EventLong & "', EventColor='" & EventColor & _
        "', EventPublished='" & EventPublished & "', EventExpires='" & EventExpires & "', EventAllDay=" & BoolToBit(EventAllDay) & _ 
        ", EventUpdate=getdate(), Reminder=" & BoolToBit(Reminder) & ", RecurringPeriodId=" & iif(isnull(RecurringPeriodId), "null", RecurringPeriodId) & _
		", Username='" &  Session("MM_Username") & "' " & _
        " WHERE EventId=" & EventId

	Conn.Execute SQL

	Set RS = Nothing
	Conn.Close : Set Conn = Nothing

	Response.Write "<p>Event modified.</p>"
	Response.End
End Sub

Sub Del(EventId)
	Conn.Execute "DELETE FROM Events WHERE EventId=" & EventId

	Set RS = Nothing
	Conn.Close : Set Conn = Nothing

	Response.Write "<p>Event deleted.</p>"
	Response.End
End Sub

function getRecurringPeriodsArray()

    set rs = conn.Execute("SELECT RecurringPeriodID, PeriodName FROM RecurringPeriods")
    getRecurringPeriodsArray = rs.GetRows
    rs.Close
    set rs = nothing
    
end function

Set RS = Nothing
Set Conn = Nothing
%>