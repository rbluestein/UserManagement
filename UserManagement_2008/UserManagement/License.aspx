<%@ Page Language="VB" AutoEventWireup="false" CodeFile="License.aspx.vb" Inherits="License" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">

	<head>
		<title runat="server" id="PageCaption"></title>
        <link href="css/BVI.css" rel="stylesheet" type="text/css" />
        <link href="css/menu.css" rel="stylesheet" type="text/css" />
        <script src="jscripts/menu.js" type="text/javascript"></script>
        <script src="jscripts/menu_items.js" type="text/javascript"></script>
        <script src="jscripts/menu_tpl.js" type="text/javascript"></script>
        <script src="jscripts/calendar2.js" type="text/javascript"></script>
	</head>
	<body>
		<form id="form1" name="form1" action="License.aspx" method="post" runat="server">
			<input type="hidden" name="hdAction" /><input type="hidden" name="hdSubAction" />
			<table class="PrimaryTbl" style="POSITION: absolute; TOP: 14px; LEFT: 140px" cellspacing="0"
				cellpadding="0" width="650" border="0">
				<tr style="DISPLAY: none">
					<td colspan="2">&nbsp;</td>
				</tr>
				<tr>
					<td class="PrimaryTblTitle" colspan="2"><asp:literal id="litHeading" runat="server"></asp:literal></td>
				</tr>
				<tr>
					<td class="CellSeparator" colspan="2"></td>
				</tr>
				<tr>
					<td colspan="2"><asp:label id="lblEnrollerName" runat="server"></asp:label></td>
				</tr>
				<tr>
					<td colspan="2">&nbsp;</td>
				</tr>
				<tr>
					<td class="Cell9Reg" style="WIDTH: 120px" align="left">State:</td>
					<td><asp:dropdownlist id="ddState" runat="server"></asp:dropdownlist><asp:textbox id="txtState" runat="server" readonly="True" width="112px" maxlength="50"></asp:textbox></td>
				</tr>
				<tr>
					<td class="Cell9Reg" style="WIDTH: 120px" align="left">License Number:</td>
					<td><asp:textbox id="txtLicenseNumber" runat="server" maxlength="50"></asp:textbox></td>
				</tr>
				<tr>
					<td class="Cell9Reg" style="WIDTH: 120px" align="left">Application Date:</td>
					<td class="Cell9Reg"><asp:textbox id="txtApplicationDate" runat="server" readonly="True" width="112px" maxlength="50"></asp:textbox>&nbsp;&nbsp;
						<asp:label id="lblApplicationDateLink" runat="server">
							<a href="javascript:GetDate('ApplicationDate')">Get Date</a></asp:label></td>
				</tr>
				<tr>
					<td class="Cell9Reg" style="WIDTH: 120px" align="left" width="120">Effective Date:</td>
					<td class="Cell9Reg"><asp:textbox id="txtEffectiveDate" runat="server" readonly="True" width="112px" maxlength="50"></asp:textbox>&nbsp;&nbsp;
						<asp:label id="lblEffectiveDateLink" runat="server">
							<a href="javascript:GetDate('EffectiveDate')">Get Date</a></asp:label></td>
				</tr>
				<tr>
					<td class="Cell9Reg" style="WIDTH: 120px" align="left">Expiration Date:</td>
					<td class="Cell9Reg"><asp:textbox id="txtExpirationDate" runat="server" readonly="True" width="112px" maxlength="50"></asp:textbox>&nbsp;&nbsp;
						<asp:label id="lblExpirationDateLink" runat="server">
							<a href="javascript:GetDate('ExpirationDate')">Get Date</a></asp:label></td>
				</tr>
				<tr>
					<td class="Cell9Reg" style="WIDTH: 120px" align="left">OK to renew?</td>
					<td><asp:dropdownlist id="ddOKToRenewInd" runat="server"></asp:dropdownlist>
						<asp:textbox style="Z-INDEX: 0" id="txtOKToRenewInd" runat="server" maxlength="50" width="112px"
							readonly="True"></asp:textbox></td>
				</tr>
				<tr>
					<td class="Cell9Reg" style="WIDTH: 120px" align="left">Renewal Date Sent:</td>
					<td class="Cell9Reg"><asp:textbox id="txtRenewalDateSent" runat="server" readonly="True" width="112px" maxlength="50"></asp:textbox>&nbsp;&nbsp;
						<asp:label id="lblRenewalDateSentLink" runat="server">
							<a href="javascript:GetDate('RenewalDateSent')">Get Date</a></asp:label></td>
				</tr>
				<tr>
					<td class="Cell9Reg" style="WIDTH: 120px" align="left">Renewal Date Recd:</td>
					<td class="Cell9Reg"><asp:textbox id="txtRenewalDateRecd" runat="server" readonly="True" width="112px" maxlength="50"></asp:textbox>&nbsp;&nbsp;
						<asp:label id="lblRenewalDateRecdLink" runat="server"><a href="javascript:GetDate('RenewalDateRecd')">
								Get Date</a>&nbsp;&nbsp</asp:label></td>
				</tr>
				<asp:placeholder id="plLongTermCareStateSpecific" runat="server">
					<tr>
						<td colspan="2">&nbsp;
						</td>
					</tr>
					<tr>
						<td colspan="2"><b><u>Long Term Care State Specific Certification</u></b></td>
					</tr>
					<tr>
						<td style="WIDTH: 120px" class="Cell9Reg" align="left">Effective Date:</td>
						<td class="Cell9Reg">
							<asp:textbox id="txtLongTermCareStateSpecificEffectiveDate" runat="server" maxlength="50" width="112px"
								readonly="True"></asp:textbox>&nbsp;&nbsp;
							<asp:label id="lblLongTermCareStateSpecificEffectiveDateLink" runat="server"><a href="javascript:GetDate('LongTermCareStateSpecificEffectiveDate')">
									Get Date</a>&nbsp;&nbsp</asp:label></td>
					</tr>
					<tr>
						<td style="WIDTH: 120px" class="Cell9Reg" align="left">Expiration Date:</td>
						<td class="Cell9Reg">
							<asp:textbox id="txtLongTermCareStateSpecificExpirationDate" runat="server" maxlength="50" width="112px"
								readonly="True"></asp:textbox>&nbsp;&nbsp;
							<asp:label id="lblLongTermCareStateSpecificExpirationDateLink" runat="server"><a href="javascript:GetDate('LongTermCareStateSpecificExpirationDate')">
									Get Date</a>&nbsp;&nbsp</asp:label></td>
					</tr>
					<tr>
						<td colspan="2">&nbsp;
						</td>
					</tr>
				</asp:placeholder>
				<tr>
					<td class="Cell9Reg" style="WIDTH: 120px" valign="top" align="left" width="35">Notes:</td>
					<td><asp:textbox id="txtNotes" runat="server" rows="10" textmode="MultiLine"></asp:textbox></td>
				</tr>
				<tr>
					<td colspan="2">&nbsp;</td>
				</tr>
				<tr>
					<td align="center" colspan="2"><asp:literal id="litUpdate" runat="server" enableviewstate="False"></asp:literal>&nbsp;&nbsp;<input onclick="ReturnToParentPage()" type="button" value="Return" /></td>
				</tr>
				<tr>
					<td colspan="2">&nbsp;</td>
				</tr>
			</table>
			<asp:literal id="litResponseAction" runat="server" enableviewstate="False"></asp:literal><asp:label id="lblCurrentRights" runat="server"></asp:label>
			<asp:literal id="litEnviro" runat="server" enableviewstate="False"></asp:literal>
			<script type="text/javascript">		
				document.write("<table style='background:#eeeedd;PADDING-LEFT: 4px;FONT: 8pt Arial, Helvetica, sans-serif; POSITION: absolute;TOP: 14px' cellSpacing='0' cellPadding='0' width='125' border='0'><tr><td width='20'>User:</td><td>" + form1.hdLoggedInUserID.value + "</td></tr><tr><td>Site:</td><td WRAP=HARD>" + form1.hdDBHost.value + "</td></tr></table>")			
					
				new menuitems("userlic1", form1.currentrights.value);
				new tpl("v", "admin");
				new menu (MENU_ITEMS, MENU_TPL);				
			</script>
			<asp:literal id="litMsg" runat="server" enableviewstate="False"></asp:literal><asp:literal id="litHiddens" runat="server" enableviewstate="False"></asp:literal>
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
					
					function ChangeYear(vIn)  {
						form1.hdAction.value = vIn
						form1.submit()
					}
					
					function GetDate(vIn)  {
						eval(vIn).popup()
					}		
					
					function ClientSelectionChanged()  {
						form1.hdAction.value = "ClientSelectionChanged"
						form1.submit()
					}
										
																								
			</script>
		</form>
		<script type="text/javascript">
			var ApplicationDate = new calendar2(document.forms['form1'].elements['txtApplicationDate']); 
			ApplicationDate.year_scroll = true; ApplicationDate.time_comp = false;
				
			var EffectiveDate = new calendar2(document.forms['form1'].elements['txtEffectiveDate']); 
			EffectiveDate.year_scroll = true; EffectiveDate.time_comp = false;	
			
			var ExpirationDate = new calendar2(document.forms['form1'].elements['txtExpirationDate']); 
			ExpirationDate.year_scroll = true; ExpirationDate.time_comp = false;
				
			var RenewalDateSent = new calendar2(document.forms['form1'].elements['txtRenewalDateSent']); 
			RenewalDateSent.year_scroll = true; RenewalDateSent.time_comp = false;	
			
			var RenewalDateRecd = new calendar2(document.forms['form1'].elements['txtRenewalDateRecd']); 
			RenewalDateRecd.year_scroll = true; RenewalDateRecd.time_comp = false;				

			var txtBox=document.forms['form1'].elements['txtLongTermCareStateSpecificEffectiveDate']; 
			if (txtBox != null) {
				var LongTermCareStateSpecificEffectiveDate = new calendar2(document.forms['form1'].elements['txtLongTermCareStateSpecificEffectiveDate']); 
				LongTermCareStateSpecificEffectiveDate.year_scroll = true; LongTermCareStateSpecificEffectiveDate.time_comp = false;				
				
				var LongTermCareStateSpecificExpirationDate = new calendar2(document.forms['form1'].elements['txtLongTermCareStateSpecificExpirationDate']); 
				LongTermCareStateSpecificExpirationDate.year_scroll = true; LongTermCareStateSpecificExpirationDate.time_comp = false;	
			}			
							
		</script>
	</body>
</html>

