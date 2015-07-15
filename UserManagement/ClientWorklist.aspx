<%@ Page Language="vb" AutoEventWireup="false" Codebehind="ClientWorklist.aspx.vb" Inherits="UserManagement.ClientWorklist"%>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
	<head>
		<title runat="server" id="PageCaption"></title>
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
		<meta content="Microsoft Visual Studio.NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<link title="BVIStyle" href="BVI.css" type="text/css" rel="stylesheet">
		<link href="menu.css" rel="stylesheet">
		<script language="JavaScript" src="menu.js"></script>
		<script language="JavaScript" src="menu_items.js"></script>
		<script language="JavaScript" src="menu_tpl.js"></script>
	</head>
	<body>
		<form id="form1" name="form1" action="ClientWorklist.aspx" method="post" runat="server">
			<input type="hidden" name="hdAction"> <input type="hidden" name="hdSortField"> <input type="hidden" name="hdClientID">
			<input type='hidden' name='hdFilterShowHideToggle' id='hdFilterShowHideToggle' value='0'>
			<table class="PrimaryTbl" style="LEFT: 140px; POSITION: absolute; TOP: 14px" cellspacing="0"
				cellpadding="0" width="650" border="0">
				<tr style="DISPLAY: none">
					<td width="100"></td>
					<td width="300"></td>
				</tr>
				<tr>
					<td class="PrimaryTblTitle" colspan="2">Clients</td>
				</tr>
				<tr>
					<td class="CellSeparator" colspan="2"></td>
				</tr>
				<tr>
					<td colspan="2"><asp:literal id="litDG" runat="server" enableviewstate="False"></asp:literal></td>
				</tr>
			</table>
			<asp:label id="lblCurrentRights" runat="server"></asp:label>
			<asp:literal id="litEnviro" runat="server" enableviewstate="False"></asp:literal>
			<script language='javascript'>				
				document.write("<table style='background:#eeeedd;PADDING-LEFT: 4px;FONT: 8pt Arial, Helvetica, sans-serif; POSITION: absolute;TOP: 14px' cellSpacing='0' cellPadding='0' width='125' border='0'><tr><td width='20'>User:</td><td>" + form1.hdLoggedInUserID.value + "</td></tr><tr><td>Site:</td><td WRAP=HARD>" + form1.hdDBHost.value + "</td></tr></table>")				
				
				new menuitems("userlic1", form1.currentrights.value);
				new tpl("v", "admin");
				new menu (MENU_ITEMS, MENU_TPL);				
			</script>
			<asp:literal id="litMsg" runat="server" enableviewstate="False"></asp:literal><asp:literal id="litFilterHiddens" runat="server" enableviewstate="False"></asp:literal>
			<script language="javascript"> 			

			function Sort(vField) {
				form1.hdAction.value = "Sort"
				form1.hdSortField.value = vField
				form1.submit()
			}
	
			function ToggleShowFilter()  {
				form1.hdFilterShowHideToggle.value = 1
				form1.hdAction.value = "ApplyFilter"
				form1.submit()
			}				
	
			function ApplyFilter()
			{		
				form1.hdAction.value = "ApplyFilter"
				form1.submit()				
			}						 						
						
			function ExistingRecord(vClientID)
			{
				form1.hdAction.value = "ExistingRecord"
				form1.hdClientID.value = vClientID
				form1.submit()
			}
			
			function Delete(vClientID)  {
				vClientID = vClientID.replace("~", "'");
				var OKToDelete = confirm("Are you sure you wish to delete " + vClientID + "?");
				if (OKToDelete == true) {
					form1.hdAction.value = "Delete"
					form1.hdClientID.value = vClientID
					form1.submit()		
				}		
			}	
			
			/*
			function License(vClientID)
			{
				form1.hdAction.value = "License"
				form1.hdClientID.value = vClientID
				form1.submit()
			}
			*/						
			
			function NewRecord() {
				form1.hdAction.value = "NewRecord"
				form1.submit()
			}				

			function SubmitOnEnterKey(e) {
				var keypressevent = e ? e : window.event
				if (keypressevent.keyCode == 13) {	
					form1.hdAction.value = "ApplyFilter"						 	
					form1.submit()
				}			
			}		
																												
			</script>
		</form>
	</body>
</html>
