<%@LANGUAGE="VBSCRIPT"%>
<%
Response.Cookies("MM_Username") = ""
Response.Cookies("UserID") = ""
Response.Cookies("PositionTitle") = ""
Response.Cookies("MM_UsernameLongTerm") = ""
Response.Cookies("UserIDLongTerm") = ""
Response.Cookies("BillingCustomerID") = ""
Session.Abandon
Response.Redirect("logon.asp")
%>
