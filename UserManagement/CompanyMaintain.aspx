<%@ Page Language="vb" AutoEventWireup="false" Codebehind="CompanyMaintain.aspx.vb" Inherits="UserManagement.CompanyMaintain" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title runat="server" id="PageCaption"></title>
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
		<form id="form1" name="form1" action="CompanyMaintain.aspx" method="post" runat="server">
			<input type="hidden" name="hdAction"><input type="hidden" id="hdConfirm" name="hdConfirm">
			<script language='javascript'>
				function ConfirmKeyChange(vMsg)  {
					var OKToChange = confirm(vMsg)
					//var OKToChange = confirm("There are users currently associated with this Company ID. If you proceed with this change in Company ID, these users will be updated to the new Company ID value. Do you wish to proceed with this change?");
					if (OKToChange == true) {
						form1.hdConfirm.value = "yes"
					}  else {
						form1.hdConfirm.value = "no"
					}					
					form1.hdAction.value = "keychange"
					form1.submit()
				}				
			</script>
			<table class="PrimaryTbl" style="LEFT: 140px; POSITION: absolute; TOP: 14px" cellSpacing="0"
				cellPadding="0" width="650" border="0">
				<tr style="DISPLAY: none">
					<td width="250">&nbsp;</td>
					<td>&nbsp;
					</td>
				</tr>
				<tr>
					<td class="PrimaryTblTitle" colSpan="2"><asp:literal id="litHeading" runat="server"></asp:literal></td>
				</tr>
				<tr>
					<td class="cellseparator" colSpan="2"></td>
				</tr>
				<tr>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<td class="Cell9Reg" align="left">Company ID:</td>
					<td><asp:textbox id="txtCompanyID" runat="server" MaxLength="50" width="300px"></asp:textbox></td>
				</tr>
				<tr>
					<td class="Cell9Reg" align="left">Description:</td>
					<td><asp:textbox id="txtDescription" runat="server" MaxLength="100" width="300px"></asp:textbox></td>
				</tr>
				<tr>
					<td class="Cell9Reg" align="left" width="120">Primary Contact Name:</td>
					<td><asp:textbox id="txtPrimaryContactName" runat="server" MaxLength="100" width="300px"></asp:textbox></td>
				</tr>
				<tr>
					<td class="Cell9Reg" align="left">Phone #:</td>
					<td><asp:textbox id="txtPrimaryContactPhone" runat="server" MaxLength="15" width="300px"></asp:textbox></td>
				</tr>
				<tr>
					<td class="Cell9Reg" align="left">Email:</td>
					<td><asp:textbox id="txtPrimaryContactEmail" runat="server" MaxLength="50" width="300px"></asp:textbox></td>
				</tr>
				<tr>
					<td class="Cell9Reg" align="left" valign='top'>Notes:</td>
					<td>
						<asp:textbox id="txtContactNotes" runat="server" TextMode="MultiLine" Rows="10" Width="300px"></asp:textbox></td>
				</tr>
				<tr>
					<td colSpan="2">&nbsp;</td>
				</tr>
				<tr>
					<td style="WIDTH: 555px" colSpan="2">&nbsp;</td>
				</tr>
				<tr>
					<td class="Cell9Reg" align="left" width="120">&nbsp;</td>
					<td align="left" colSpan="2"><asp:literal id="litUpdate" runat="server" EnableViewState="False"></asp:literal>&nbsp;&nbsp;<INPUT onclick="ReturnToParentPage()" type="button" value="Return"></td>
				</tr>
				<tr>
					<td style="WIDTH: 555px" colSpan="2">&nbsp;</td>
				</tr>
			</table>
			<asp:literal id="litResponseAction" runat="server" EnableViewState="False"></asp:literal>
			<asp:Label id="lblCurrentRights" runat="server"></asp:Label>
			<asp:Literal id="litEnviro" runat="server" EnableViewState="False"></asp:Literal>
			
			<script language='javascript'>		
				document.write("<table style='background:#eeeedd;PADDING-LEFT: 4px;FONT: 8pt Arial, Helvetica, sans-serif; POSITION: absolute;TOP: 14px' cellSpacing='0' cellPadding='0' width='125' border='0'><tr><td width='20'>User:</td><td>" + form1.hdLoggedInUserID.value + "</td></tr><tr><td>Site:</td><td WRAP=HARD>" + form1.hdDBHost.value + "</td></tr></table>")
				
				new menuitems("userlic1", form1.currentrights.value);
				new tpl("v", "admin");
				new menu (MENU_ITEMS, MENU_TPL);				
			</script>
			<asp:literal id="litMsg" runat="server" EnableViewState="False"></asp:literal>
			<script type="text/javascript">
			function legallength(txtarea, legalmax)  {	
				if (txtarea.getAttribute && txtarea.value.length>legalmax) {
					txtarea.value=txtarea.value.substring(0,legalmax)
				}
			}
			function Update()  {
				form1.hdAction.value = "update"
				form1.submit()
			}
			
			function ReturnToParentPage()  {
				form1.hdAction.value = "return"
				form1.submit()
			}
			
       	function ToggleCompany(vIn) {
				if (vIn == "BVI") {
					if (form1.chkOther.checked) {
						form1.chkOther.checked = false
					} else if (!form1.chkBVI.checked)  {
						form1.chkBVI.checked = true
					}
				}			
				else if (vIn == "Other") {
					if (form1.chkOther.checked) {
						form1.chkBVI.checked = false
					} else if (!form1.chkOther.checked) {
						form1.chkOther.checked = true
					}
				}						
			}						
			
			</script>
		</form>
	</body>
</HTML>
