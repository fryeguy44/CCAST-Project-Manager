<%

' *** Restrict Access To Page: Grant or deny access to this page
MM_authFailedURL="logon.asp"
MM_grantAccess=false
bolEmailNeedsUpdate=false
'Response.Write("MM_Username = " & Session("MM_Username") & "<br />")
If Session("MM_Username") <> "" Then
    MM_grantAccess = true
Else
'Response.Write("PositionID = " & Request.Cookies("PositionID") & "<br />")
	If 	Request.Cookies("UserID") <> "" Then
		Session("MM_Username") = Request.Cookies("MM_Username")
		Session("UserID") = Request.Cookies("UserID")
		Session("PositionID") = Request.Cookies("PositionID")
		Session("VendorID") = Request.Cookies("VendorID")
		Session("ClientID") = Request.Cookies("ClientID")
				'Response.Write("Cookie VendorID = " & Request.Cookies("VendorID") & "<br />")
		MM_grantAccess = true
	Else
'Response.Write("MM_UsernameLongTerm = " & Request.Cookies("MM_UsernameLongTerm") & "<br />")
		If 	Request.Cookies("UserIDLongTerm") = "" Then
			MM_grantAccess = false
		Else
			MM_grantAccess = true

			Set rstSecurityOne_cmd = Server.CreateObject ("ADODB.Command")
			rstSecurityOne_cmd.ActiveConnection = MM_OBA_STRING
		
			rstSecurityOne_cmd.CommandText = "SELECT UserID, UserName, PositionID, VendorID, ClientID FROM Users WHERE (UserID = ?) AND Active = 1"
			rstSecurityOne_cmd.Parameters.Append rstSecurityOne_cmd.CreateParameter("param1", 5, 1, -1, Request.Cookies("UserIDLongTerm")) ' adVarWChar
			Set rstSecurityOne = rstSecurityOne_cmd.Execute
			If rstSecurityOne.EOF or rstSecurityOne.BOF Then
				rstSecurityOne.Close()
				Set rstSecurityOne = Nothing
				Response.Redirect("logoff.asp")

			Else
				'Response.Write("Security VendorID = " & rstSecurityOne.Fields.Item("VendorID").Value & "<br />")
				Session("PositionID") = rstSecurityOne.Fields.Item("PositionID").Value
				Session("MM_Username") = rstSecurityOne.Fields.Item("UserName").Value
				Session("UserID") = rstSecurityOne.Fields.Item("UserID").Value
				Session("VendorID") = rstSecurityOne.Fields.Item("VendorID").Value
				Session("ClientID") = rstSecurityOne.Fields.Item("ClientID").Value
				Response.Cookies("PositionID") = rstSecurityOne.Fields.Item("PositionID").Value
				Response.Cookies("UserID") = rstSecurityOne.Fields.Item("UserID").Value
				Response.Cookies("MM_Username") = rstSecurityOne.Fields.Item("UserName").Value
				Response.Cookies("VendorID") = rstSecurityOne.Fields.Item("VendorID").Value
				Response.Cookies("ClientID") = rstSecurityOne.Fields.Item("ClientID").Value
				
			End If
			Set rstSecurityOne = Nothing
		
		End If
	End If
End If
'Response.Write("MM_Username = " & Session("MM_Username") & "<br />")
'Response.Write("PositionID = " & Session("PositionID") & "<br />")
'Response.Write("MM_UsernameLongTerm = " & Request.Cookies("MM_UsernameLongTerm") & "<br />")
'Response.Write("MM_grantAccess = " & MM_grantAccess & "<br />")

strPageUrl = MID(Request.ServerVariables("URL"),2)
If MM_grantAccess Then

	
	
	
Dim intHelpContextID
Dim rstSecurity
Dim rstSecurity_cmd
Dim rstSecurity_numRows
	Set rstSecurity_cmd = Server.CreateObject ("ADODB.Command")
	
	rstSecurity_cmd.ActiveConnection = MM_OBA_STRING

	rstSecurity_cmd.CommandText = "SELECT Grants.GrantLevelID, Pages.PageTitle, Pages.HelpContextID, Elements.ElementName FROM ((Grants INNER JOIN Elements ON Grants.ElementID = Elements.ElementID) INNER JOIN PageElements ON Elements.ElementID = PageElements.ElementID) INNER JOIN Pages ON PageElements.PageID = Pages.PageID WHERE (Grants.PositionID= ?  AND Pages.PageAddress= ?  AND Pages.Active = 1) GROUP BY Grants.GrantLevelID, Pages.PageTitle, Pages.HelpContextID, Elements.ElementName ORDER BY Elements.ElementName" 
	rstSecurity_cmd.Prepared = true
	rstSecurity_cmd.Parameters.Append rstSecurity_cmd.CreateParameter("param1", 5, 1, -1, Session("PositionID")) ' adVarWChar
	rstSecurity_cmd.Parameters.Append rstSecurity_cmd.CreateParameter("param1", 202, 1, 255, strPageUrl) ' adVarWChar
	Set rstSecurity = rstSecurity_cmd.Execute

Dim strPageTitle
	If rstSecurity.EOF Then
		strPageTitle = "No Security Set for this page"
		intHelpContextID = 120
	Else
		strPageTitle = rstSecurity.Fields.Item("PageTitle").Value
		intHelpContextID = rstSecurity.Fields.Item("HelpContextID").Value
	End If

	Do While Not rstSecurity.EOF

		Select Case rstSecurity.Fields.Item("ElementName").Value
			Case "Support"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolDeveloperViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolDeveloperEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolDeveloperAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolDeveloperDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolDeveloperFullGranted = True
				End If
					
			Case "Security"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolSecurityViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolSecurityEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolSecurityAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolSecurityDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolSecurityFullGranted = True
				End If
		
			Case "Users"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolUsersViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolUsersEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolUsersAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolUsersDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolUsersFullGranted = True
				End If
					
			Case "Developer"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolDeveloperViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolDeveloperEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolDeveloperAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolDeveloperDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolDeveloperFullGranted = True
				End If
			
			Case "Clients"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolClientsViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolClientsEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolClientsAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolClientsDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolClientsFullGranted = True
				End If
		
			Case "Contacts"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolContactsViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolContactsEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolContactsAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolContactsDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolContactsFullGranted = True
				End If
		
			Case "Payments"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolPaymentsViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolPaymentsEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolPaymentsAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolPaymentsDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolPaymentsFullGranted = True
				End If
		
			Case "Vendors"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolVendorsViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolVendorsEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolVendorsAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolVendorsDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolVendorsFullGranted = True
				End If
		
			Case "Financials"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolFinancialsViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolFinancialsEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolFinancialsAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolFinancialsDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolFinancialsFullGranted = True
				End If
		
			Case "TruckTypes"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolTruckTypesViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolTruckTypesEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolTruckTypesAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolTruckTypesDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolTruckTypesFullGranted = True
				End If
		
			Case "TrailerTypes"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolTrailerTypesViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolTrailerTypesEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolTrailerTypesAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolTrailerTypesDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolTrailerTypesFullGranted = True
				End If
		
			Case "Trucks"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolTrucksViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolTrucksEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolTrucksAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolTrucksDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolTrucksFullGranted = True
				End If
		
			Case "Drivers"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolDriversViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolDriversEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolDriversAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolDriversDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolDriversFullGranted = True
				End If
		
			Case "Jobs"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolJobsViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolJobsEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolJobsAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolJobsDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolJobsFullGranted = True
				End If
		
			Case "Companys"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolCompanysViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolCompanysEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolCompanysAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolCompanysDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolCompanysFullGranted = True
				End If
		
			Case "Loads"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolLoadsViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolLoadsEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolLoadsAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolLoadsDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolLoadsFullGranted = True
				End If
		
			Case "CannedPhrases"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolCannedPhrasesViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolCannedPhrasesEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolCannedPhrasesAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolCannedPhrasesDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolCannedPhrasesFullGranted = True
				End If

			Case "Bids"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolBidsViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolBidsEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolBidsAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolBidsDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolBidsFullGranted = True
				End If
		
			Case "FSC"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolFSCViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolFSCEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolFSCAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolFSCDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolFSCFullGranted = True
				End If
		
			Case "TimePunch"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolTimePunchViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolTimePunchEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolTimePunchAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolTimePunchDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolTimePunchFullGranted = True
				End If
		
			Case "DayOffs"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolDayOffsViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolDayOffsEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolDayOffsAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolDayOffsDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolDayOffsFullGranted = True
				End If
		
			Case "Maintenance"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolMaintenanceViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolMaintenanceEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolMaintenanceAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolMaintenanceDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolMaintenanceFullGranted = True
				End If
		
			Case "Projects"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolProjectsViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolProjectsEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolProjectsAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolProjectsDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolProjectsFullGranted = True
				End If
		
			Case "Invoices"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolInvoicesViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolInvoicesEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolInvoicesAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolInvoicesDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolInvoicesFullGranted = True
				End If
		
			Case "Mileage"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolMileageViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolMileageEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolMileageAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolMileageDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolMileageFullGranted = True
				End If
		
			Case "Insurance"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolInsuranceViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolInsuranceEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolInsuranceAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolInsuranceDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolInsuranceFullGranted = True
				End If
		
			Case "TransferFiles"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolTransferFilesViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolTransferFilesEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolTransferFilesAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolTransferFilesDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolTransferFilesFullGranted = True
				End If
		
			Case "COI"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolCOIViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolCOIEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolCOIAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolCOIDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolCOIFullGranted = True
				End If
		
			Case "Training"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolTrainingViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolTrainingEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolTrainingAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolTrainingDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolTrainingFullGranted = True
				End If
		
			Case "Tickets"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolTicketsViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolTicketsEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolTicketsAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolTicketsDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolTicketsFullGranted = True
				End If
		
			Case "AxonInvoices"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolAxonInvoicesViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolAxonInvoicesEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolAxonInvoicesAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolAxonInvoicesDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolAxonInvoicesFullGranted = True
				End If
		
			Case "Briefings"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolBriefingsViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolBriefingsEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolBriefingsAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolBriefingsDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolBriefingsFullGranted = True
				End If
		
			Case "AxonJobs"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolAxonJobsViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolAxonJobsEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolAxonJobsAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolAxonJobsDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolAxonJobsFullGranted = True
				End If
		
			Case "UnprocessedTickets"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolUnprocessedTicketsViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolUnprocessedTicketsEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolUnprocessedTicketsAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolUnprocessedTicketsDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolUnprocessedTicketsFullGranted = True
				End If
		
			Case "Incidents"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolIncidentsViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolIncidentsEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolIncidentsAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolIncidentsDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolIncidentsFullGranted = True
				End If
		
			Case "ClientOnly"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolClientOnlyViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolClientOnlyEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolClientOnlyAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolClientOnlyDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolClientOnlyFullGranted = True
				End If
		
			Case "VendorOnly"
				If rstSecurity.Fields.Item("GrantLevelID").Value >  0 Then
					bolVendorOnlyViewGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  1 Then
					bolVendorOnlyEditGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  2 Then
					bolVendorOnlyAddGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  3 Then
					bolVendorOnlyDeleteGranted = True
				End If
				If rstSecurity.Fields.Item("GrantLevelID").Value >  4 Then
					bolVendorOnlyFullGranted = True
				End If
		
		
		End Select

		rstSecurity.MoveNext
	Loop

	rstSecurity.Close()
	Set rstSecurity = Nothing
		
	dteAccessGranted = Now()

Else
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
'Response.Write("MM_Username = " & Session("MM_Username") & "<br />" & "PositionID = |" & Request.Cookies("PositionID") & "|<br />" & "MM_UsernameLongTerm = " & Request.Cookies("MM_UsernameLongTerm") & "<br />")
End If
%>