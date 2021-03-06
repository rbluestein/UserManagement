﻿<%@ Page Language="VB" AutoEventWireup="false" CodeFile="CarrierMaintain.aspx.vb" Inherits="CarrierMaintain" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
	<head>	
		<title id="PageCaption" runat="server"></title>
        <link href="css/BVI.css" rel="stylesheet" type="text/css" />
        <link href="css/menu.css" rel="stylesheet" type="text/css" />
        <link href="css/BVI.css" rel="stylesheet" type="text/css" />
        <script src="jscripts/calendar2.js" type="text/javascript"></script>
        <script src="jscripts/menu_items.js" type="text/javascript"></script>
        <script src="jscripts/menu.js" type="text/javascript"></script>
        <script src="jscripts/menu_tpl.js" type="text/javascript"></script>	        
        <link href="css/BVI.css" rel="stylesheet" type="text/css" />
        <link href="css/menu.css" rel="stylesheet" type="text/css" />
	</head>
	<body>
		<form id="form1" name="form1" action="CarrierMaintain.aspx" method="post" runat="server">
			<input type="hidden" name="hdAction" /><input id="hdConfirm" type="hidden" name="hdConfirm" />
			<script type="text/javascript">
			//	function GetUserConfirmation()  {
			//		var OKToChange = confirm("There are users currently associated with this Carrier ID. If you proceed with this change in Carrier ID, these users will be updated to the new Carrier ID value. Do you wish to proceed with this change?");
			//		if (OKToChange == true) {
			//			form1.hdConfirm.value = "yes"
			//		}  else {
			//			form1.hdConfirm.value = "no"
			//		}					
			//		form1.hdAction.value = "confirmation"
			//		form1.submit()
			//	}				
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
					<td class="Cell9Reg" align="left">Carrier ID:</td>
					<td><asp:textbox id="txtCarrierID" runat="server" width="300px" maxlength="50"></asp:textbox></td>
				</tr>
				<tr>
					<td class="Cell9Reg" valign="top" align="left">Description:</td>
					<td><asp:textbox id="txtDescription" runat="server" width="300px" rows="5" textmode="MultiLine"></asp:textbox></td>
				</tr>
				<tr>
					<td colspan="2">&nbsp;</td>
				</tr>
				<tr>
					<td style="WIDTH: 555px" colspan="2">&nbsp;</td>
				</tr>
				<tr>
					<td class="Cell9Reg" align="left" width="120">&nbsp;</td>
					<td align="left" colspan="2"><asp:literal id="litUpdate" runat="server" enableviewstate="False"></asp:literal>&nbsp;&nbsp;<input onclick="ReturnToParentPage()" type="button" value="Return" /></td>
				</tr>
				<tr>
					<td style="WIDTH: 555px" colspan="2">&nbsp;</td>
				</tr>
			</table>
			<asp:literal id="litResponseAction" runat="server" enableviewstate="False"></asp:literal><asp:label id="lblCurrentRights" runat="server"></asp:label>
			<asp:literal id="litEnviro" runat="server" enableviewstate="False"></asp:literal>
			<script type="text/javascript">	
				document.write("<table style='background:#eeeedd;PADDING-LEFT: 4px;FONT: 8pt Arial, Helvetica, sans-serif; POSITION: absolute;TOP: 14px' cellspacing='0' cellpadding='0' width='125' border='0'><tr><td width='20'>User:</td><td>" + form1.hdLoggedInUserID.value + "</td></tr><tr><td>Site:</td><td WRAP=HARD>" + form1.hdDBHost.value + "</td></tr></table>")
				
				new menuitems("userlic1", form1.currentrights.value);
				new tpl("v", "admin");
				new menu (MENU_ITEMS, MENU_TPL);				
			</script>
			<asp:literal id="litMsg" runat="server" enableviewstate="False"></asp:literal>
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
</html>
