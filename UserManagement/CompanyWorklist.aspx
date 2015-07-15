<%@ Page Language="vb" AutoEventWireup="false" Codebehind="CompanyWorklist.aspx.vb" Inherits="UserManagement.CompanyWorklist" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title runat="server" id="PageCaption"></title>
		<META http-equiv="Content-Type" content="text/html; charset=windows-1252">
		<meta content="Microsoft Visual Studio.NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK title="BVIStyle" href="BVI.css" type="text/css" rel="stylesheet">
		<LINK href="menu.css" rel="stylesheet">
		<script language="JavaScript" src="menu.js"></script>
		<script language="JavaScript" src="menu_items.js"></script>
		<script language="JavaScript" src="menu_tpl.js"></script>
	</HEAD>
	<body>
		<form id="form1" name="form1" action="Company.aspx" method="post" runat="server">
			<input type="hidden" name="hdAction"> <input type="hidden" name="hdSortField"> <input type="hidden" name="hdCompanyID">
			<input type='hidden' name='hdFilterShowHideToggle' id='hdFilterShowHideToggle' value='0'>
			<table class="PrimaryTbl" style="LEFT: 140px; POSITION: absolute; TOP: 14px" cellSpacing="0"
				cellPadding="0" width="650" border="0">
				<tr style="DISPLAY: none">
					<td width="100"></td>
					<td width="300"></td>
				</tr>
				<tr>
					<td class="PrimaryTblTitle" colSpan="2">Company Contacts</td>
				</tr>
				<tr>
					<td class="CellSeparator" colSpan="2"></td>
				</tr>
				<tr>
					<td colSpan="2"><asp:literal id="litDG" runat="server" EnableViewState="False"></asp:literal></td>
				</tr>
			</table>
			<asp:Label id="lblCurrentRights" runat="server"></asp:Label>
			<asp:Literal id="litEnviro" runat="server" EnableViewState="False"></asp:Literal>			
			<script language='javascript'>				
				document.write("<table style='background:#eeeedd;PADDING-LEFT: 4px;FONT: 8pt Arial, Helvetica, sans-serif; POSITION: absolute;TOP: 14px' cellSpacing='0' cellPadding='0' width='125' border='0'><tr><td width='20'>User:</td><td>" + form1.hdLoggedInUserID.value + "</td></tr><tr><td>Site:</td><td WRAP=HARD>" + form1.hdDBHost.value + "</td></tr></table>")				
				
				new menuitems("userlic1", form1.currentrights.value);
				new tpl("v", "admin");
				new menu (MENU_ITEMS, MENU_TPL);				
			</script>
			<asp:literal id="litMsg" runat="server" EnableViewState="False"></asp:literal><asp:literal id="litHiddens" runat="server" EnableViewState="False"></asp:literal>
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
						
			function ExistingRecord(vCompanyID)
			{
				form1.hdAction.value = "ExistingRecord"
				form1.hdCompanyID.value = vCompanyID
				form1.submit()
			}
			
			function Delete(vCompanyID)  {
				vCompanyID = vCompanyID.replace("~", "'");
				var OKToDelete = confirm("Are you sure you wish to delete " + vCompanyID + "?");
				if (OKToDelete == true) {
					form1.hdAction.value = "Delete"
					form1.hdCompanyID.value = vCompanyID
					form1.submit()		
				}		
			}				
				
			function License(vCompanyID)
			{
				form1.hdAction.value = "License"
				form1.hdCompanyID.value = vCompanyID
				form1.submit()
			}						
			
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
</HTML>
