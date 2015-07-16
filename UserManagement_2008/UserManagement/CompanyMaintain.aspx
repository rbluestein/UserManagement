<%@ Page Language="VB" AutoEventWireup="false" CodeFile="CompanyMaintain.aspx.vb" Inherits="CompanyMaintain" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<title runat="server" id="PageCaption"></title>			
        <link href="css/BVI.css" rel="stylesheet" type="text/css" />
        <link href="css/menu.css" rel="stylesheet" type="text/css" />
        <script src="jscripts/menu_items.js" type="text/javascript"></script>
        <script src="jscripts/menu.js" type="text/javascript"></script>
        <script src="jscripts/menu_tpl.js" type="text/javascript"></script>			
	</head>
	<body>
		<form id="form1" name="form1" action="CompanyMaintain.aspx" method="post" runat="server">
			<input type="hidden" id="hdAction" name="hdAction" /><input type="hidden" id="hdConfirm" name="hdConfirm" />
			<script type="text/javascript">
				function ConfirmKeyChange(vMsg)  {
					var OKToChange = confirm(vMsg)
					//var OKToChange = confirm("There are users currently associated with this Company ID. If you proceed with this change in Company ID, these users will be updated to the new Company ID value. Do you wish to proceed with this change?");
					if (OKToChange == true) {
						document.getElementById("hdConfirm").value = "yes"
					}  else {
						document.getElementById("hdConfirm").value = "no"
					}					
					document.getElementById("hdAction").value = "keychange"
					form1.submit()
				}				
			</script>
			<table class="PrimaryTbl" style="LEFT: 140px; POSITION: absolute; TOP: 14px" cellspacing="0"
				cellpadding="0" width="650" border="0">
				<tr style="DISPLAY: none">
					<td width="250">&nbsp;</td>
					<td>&nbsp;
					</td>
				</tr>
				<tr>
					<td class="PrimaryTblTitle" colspan="2"><asp:literal id="litHeading" runat="server"></asp:literal></td>
				</tr>
				<tr>
					<td class="CellSeparator" colspan="2"></td>
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
					<td colspan="2">&nbsp;</td>
				</tr>
				<tr>
					<td style="WIDTH: 555px" colspan="2">&nbsp;</td>
				</tr>
				<tr>
					<td class="Cell9Reg" align="left" width="120">&nbsp;</td>
					<td align="left" colspan="2"><asp:literal id="litUpdate" runat="server" EnableViewState="False"></asp:literal>&nbsp;&nbsp;<input onclick="ReturnToParentPage()" type="button" value="Return" /></td>
				</tr>
				<tr>
					<td style="WIDTH: 555px" colspan="2">&nbsp;</td>
				</tr>
			</table>
			<asp:literal id="litResponseAction" runat="server" EnableViewState="False"></asp:literal>
			<asp:Label id="lblCurrentRights" runat="server"></asp:Label>
			<asp:Literal id="litEnviro" runat="server" EnableViewState="False"></asp:Literal>
			
			<script type="text/javascript">
			    document.write("<table style='background:#eeeedd;PADDING-LEFT: 4px;FONT: 8pt Arial, Helvetica, sans-serif; POSITION: absolute;TOP: 14px' cellspacing='0' cellpadding='0' width='125' border='0'><tr><td width='20'>User:</td><td>" + document.getElementById("hdLoggedInUserID").value + "</td></tr><tr><td>Site:</td><td WRAP=HARD>" + document.getElementById("hdDBHost").value + "</td></tr></table>")
				
				//new menuitems("userlic1", form1.currentrights.value);
				new menuitems("userlic1", document.getElementById("currentrights").value);
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
			    document.getElementById("hdAction").value = "update"
				form1.submit()
			}
			
			function ReturnToParentPage()  {
			    document.getElementById("hdAction").value = "return"
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
</html>

