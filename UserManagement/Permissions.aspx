<%@ Page Language="vb" AutoEventWireup="false" Codebehind="Permissions.aspx.vb" Inherits="UserManagement.Permissions" %>
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
		<form id="form1" name="form1" action="Permissions.aspx" method="post" runat="server">
			<asp:Literal id="litResponseAction" runat="server" EnableViewState="False"></asp:Literal>
			<input type="hidden" name="hdAction"> <input type="hidden" name="hdSortField"> <input type="hidden" name="hdUserID">
			<table class="PrimaryTbl" style="LEFT: 140px; POSITION: absolute; TOP: 14px" cellSpacing="0"
				cellPadding="0" width="650" border="0">
				<tr>
					<td class="PrimaryTblTitle" colSpan="2">
						<asp:literal id="litHeading" runat="server"></asp:literal></td>
				</tr>
				<tr>
					<td class="CellSeparator" colSpan="2"></td>
				</tr>
				<tr>
					<td width="100">Name</td>
					<td width="550"><asp:label id="lblFullName" runat="server"></asp:label></td>
				</tr>
				<tr>
					<td width="150">Role</td>
					<td><asp:label id="lblRole" runat="server"></asp:label></td>
				</tr>
				<tr>
					<td width="150">Company</td>
					<td><asp:label id="lblCompany" runat="server"></asp:label></td>
				</tr>
				<tr>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<td colSpan="2"><asp:literal id="litDG" runat="server"></asp:literal></td>
				</tr>
				<tr>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<td colSpan="2"></td>
				</tr>
				<tr>
					<td class="Cell9Reg" align="left" width="150">&nbsp;</td>
					<td align="left" colSpan="2"><asp:literal id="litUpdate" runat="server" EnableViewState="False"></asp:literal>&nbsp;&nbsp;<INPUT onclick="ReturnToParentPage()" type="button" value="Return"></td>
				</tr>
			</table>
			<asp:label id="lblCurrentRights" runat="server"></asp:label>
			<asp:Literal id="litEnviro" runat="server" EnableViewState="False"></asp:Literal>
			<script language='javascript'>	
				document.write("<table style='background:#eeeedd;PADDING-LEFT: 4px;FONT: 8pt Arial, Helvetica, sans-serif; POSITION: absolute;TOP: 14px' cellSpacing='0' cellPadding='0' width='125' border='0'><tr><td width='20'>User:</td><td>" + form1.hdLoggedInUserID.value + "</td></tr><tr><td>Site:</td><td WRAP=HARD>" + form1.hdDBHost.value + "</td></tr></table>")			
						
				new menuitems("userlic1", form1.currentrights.value);
				new tpl("v", "admin");
				new menu (MENU_ITEMS, MENU_TPL);				
			</script>
			<asp:literal id="litMsg" runat="server" EnableViewState="False"></asp:literal><asp:literal id="litFilterHiddens" runat="server" EnableViewState="False"></asp:literal>
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
			
			function Update(vUserID)
			{
				form1.hdAction.value = "Update"
				form1.hdUserID.value = vUserID
				form1.submit()
			}									 						
						
			function ExistingRecord(vUserID)
			{
				form1.hdAction.value = "ExistingRecord"
				form1.hdUserID.value = vUserID
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
			function ReturnToParentPage()  {
				form1.hdAction.value = "return"
				form1.submit()
			}				
																									
			</script>
		</form>
	</body>
</HTML>
