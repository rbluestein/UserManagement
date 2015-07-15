<%@ Page Language="vb" AutoEventWireup="false" Codebehind="BulkLicense.aspx.vb" Inherits="UserManagement.BulkLicense"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
		<title runat="server" id="PageCaption"></title>
		<meta content="Microsoft Visual Studio.NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<link title="BVIStyle" href="BVI.css" type="text/css" rel="stylesheet">
		<script language="JavaScript" src="Gn.js"></script>
		<link href="menu.css" rel="stylesheet">
		<script language="JavaScript" src="menu.js"></script>
		<script language="JavaScript" src="menu_items.js"></script>
		<script language="JavaScript" src="menu_tpl.js"></script>
		<script language="JavaScript" src="calendar2.js"></script>
	</head>
	<body>
		<form id="form1" name="form1" action="BulkLicense.aspx" method="post" runat="server">
			<input type="hidden" name="hdAction"><input type="hidden" name="hdSubAction">
			<asp:literal id="litHiddens" runat="server" enableviewstate="False"></asp:literal>
			<table class="PrimaryTbl" style="LEFT: 140px; POSITION: absolute; TOP: 14px" cellspacing="0"
				cellpadding="0" width="650" border="0">
				<tr style="DISPLAY: none">
					<td>&nbsp;</td>
				</tr>
				<tr>
					<td class="PrimaryTblTitle"><asp:literal id="litHeading" runat="server"></asp:literal></td>
				</tr>
				<tr>
					<td class="cellseparator"></td>
				</tr>
				<tr>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<td colspan="1">
						<table class="PrimaryTblEmbedded" cellspacing="0" cellpadding="0" width="650" border="0">
							<tr>
								<td style="FONT: 16pt Arial, Helvetica, sans-serif" colspan="4">Enroller</td>
							</tr>
							<tr>
								<td width="250"><a href="javascript:Letter('A')">A</a><a href="javascript:Letter('B')">B</a><a href="javascript:Letter('C')">C</a><a href="javascript:Letter('D')">D</a><a href="javascript:Letter('E')">E</a><a href="javascript:Letter('F')">F</a><a href="javascript:Letter('G')">G</a><a href="javascript:Letter('H')">H</a><a href="javascript:Letter('I')">I</a><a href="javascript:Letter('J')">J</a><a href="javascript:Letter('K')">K</a><a href="javascript:Letter('L')">L</a><a href="javascript:Letter('M')">M</a><a href="javascript:Letter('N')">N</a><a href="javascript:Letter('O')">O</a><a href="javascript:Letter('P')">P</a><a href="javascript:Letter('Q')">Q</a><a href="javascript:Letter('R')">R</a><a href="javascript:Letter('S')">S</a><a href="javascript:Letter('T')">T</a><a href="javascript:Letter('U')">U</a><a href="javascript:Letter('V')">V</a><a href="javascript:Letter('W')">W</a><a href="javascript:Letter('X')">X</a><a href="javascript:Letter('Y')">Y</a><a href="javascript:Letter('Z')">Z</a>
									<asp:dropdownlist id="ddEnrollersSource" runat="server"></asp:dropdownlist></td>
								<td align="center" width="50"><asp:button id="btnEnrollerAdd" runat="server" width="50px" text="  >>  "></asp:button><asp:button id="btnEnrollerRemove" runat="server" width="50px" text="  <<  "></asp:button></td>
								<td align="right" width="250">&nbsp;<br>
									<asp:textbox id="txtEnrollerTarget" runat="server"></asp:textbox></td>
								<td width="100"></td>
							</tr>
							<tr>
								<td colspan="4">&nbsp;</td>
							</tr>
							<tr>
								<td style="FONT: 16pt Arial, Helvetica, sans-serif" colspan="4">State</td>
							</tr>
							<tr>
								<td width="250"><asp:listbox id="lstStatesSource" runat="server" width="160px" rows="7"></asp:listbox></td>
								<td align="center" width="50"><asp:button id="btnStateAdd" runat="server" width="50px" text="  >>  "></asp:button><asp:button id="btnStateRemove" runat="server" width="50px" text="  <<  "></asp:button></td>
								<td align="right" width="250"><asp:textbox id="txtStateFullTarget" runat="server"></asp:textbox></td>
								<td width="100"></td>
							</tr>
							<tr>
								<td colspan="4">&nbsp;</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td colspan="2">
						<table class="PrimaryTblEmbedded" cellspacing="0" cellpadding="0" width="650" border="0">
							<tr>
								<td class="Cell9Reg" style="WIDTH: 200px" align="left">Status:</td>
								<td><asp:textbox id="txtStatus" runat="server" maxlength="50"></asp:textbox></td>
							</tr>
							<tr>
								<td class="Cell9Reg" style="WIDTH: 200px" align="left">License Number:</td>
								<td><asp:textbox id="txtLicenseNumber" runat="server" maxlength="50"></asp:textbox></td>
							</tr>
							<tr>
								<td class="Cell9Reg" style="WIDTH: 200px" align="left">Application Date:</td>
								<td class="Cell9Reg"><asp:textbox id="txtApplicationDate" runat="server" width="112px" maxlength="50" readonly="True"></asp:textbox>&nbsp;&nbsp;
									<asp:label id="lblApplicationDateLink" runat="server">
										<a href="javascript:GetDate('ApplicationDate')">Get Date</a></asp:label></td>
							</tr>
							<tr>
								<td class="Cell9Reg" style="WIDTH: 200px" align="left">Effective Date:</td>
								<td class="Cell9Reg"><asp:textbox id="txtEffectiveDate" runat="server" width="112px" maxlength="50" readonly="True"></asp:textbox>&nbsp;&nbsp;
									<asp:label id="lblEffectiveDateLink" runat="server">
										<a href="javascript:GetDate('EffectiveDate')">Get Date</a></asp:label></td>
							</tr>
							<tr>
								<td class="Cell9Reg" style="WIDTH: 200px" align="left">Expiration Date:</td>
								<td class="Cell9Reg"><asp:textbox id="txtExpirationDate" runat="server" width="112px" maxlength="50" readonly="True"></asp:textbox>&nbsp;&nbsp;
									<asp:label id="lblExpirationDateLink" runat="server">
										<a href="javascript:GetDate('ExpirationDate')">Get Date</a></asp:label></td>
							</tr>
							<tr>
								<td class="Cell9Reg" style="WIDTH: 200px" align="left">Renewal Date Sent:</td>
								<td class="Cell9Reg"><asp:textbox id="txtRenewalDateSent" runat="server" width="112px" maxlength="50" readonly="True"></asp:textbox>&nbsp;&nbsp;
									<asp:label id="lblRenewalDateSentLink" runat="server">
										<a href="javascript:GetDate('RenewalDateSent')">Get Date</a></asp:label></td>
							</tr>
							<asp:placeholder id="plRenewalDateRecd" runat="server" visible="False">
								<tr>
									<td class="Cell9Reg" style="WIDTH: 200px" align="left">Renewal Date Recd:</td>
									<td class="Cell9Reg">
										<asp:textbox id="txtRenewalDateRecd" runat="server" width="112px" maxlength="50" readonly="True"></asp:textbox>&nbsp;&nbsp;
										<asp:label id="lblRenewalDateRecdLink" runat="server"><a href="javascript:GetDate('RenewalDateRecd')">
												Get Date</a>&nbsp;&nbsp</asp:label></td>
								</tr>
							</asp:placeholder>
							<asp:placeholder id="plLongTermCareStateSpecific" runat="server">
								<tr>
									<td class="Cell9Reg" style="WIDTH: 200px" align="left">LTC State Specific Eff Date:</td>
									<td>
										<asp:textbox id="txtLongTermCareStateSpecificEffectiveDate" runat="server" width="112px" maxlength="50"
											readonly="True"></asp:textbox>&nbsp;&nbsp;
										<asp:label id="lblLongTermCareStateSpecificEffectiveDateLink" runat="server"><a href="javascript:GetDate('LongTermCareStateSpecificEffectiveDate')">
												Get Date</a>&nbsp;&nbsp</asp:label></td>
								</tr>
								<tr>
									<td class="Cell9Reg" style="WIDTH: 200px; HEIGHT: 23px" align="left">LTC State 
										Specific&nbsp;Exp Date:</td>
									<td class="Cell9Reg" style="HEIGHT: 23px">
										<asp:textbox id="txtLongTermCareStateSpecificExpirationDate" runat="server" width="112px" maxlength="50"
											readonly="True"></asp:textbox>&nbsp;&nbsp;
										<asp:label id="lblLongTermCareStateSpecificExpirationDateLink" runat="server"><a href="javascript:GetDate('LongTermCareStateSpecificExpirationDate')">
												Get Date</a>&nbsp;&nbsp</asp:label></td>
								</tr>
							</asp:placeholder>
							<tr>
								<td class="Cell9Reg" style="WIDTH: 200px" valign="top" align="left" width="35">Notes:</td>
								<td><asp:textbox id="txtNotes" runat="server" rows="3" textmode="MultiLine"></asp:textbox></td>
							</tr>
							<tr>
								<td colspan="2">&nbsp;</td>
							</tr>
							<tr>
								<td align="center" colspan="2"><asp:literal id="litUpdate" runat="server" enableviewstate="False"></asp:literal>&nbsp;&nbsp;</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td colspan="2">&nbsp;</td>
				</tr>
			</table>
			<asp:literal id="litResponseAction" runat="server" enableviewstate="False"></asp:literal><asp:label id="lblCurrentRights" runat="server"></asp:label>
			<asp:literal id="litEnviro" runat="server" enableviewstate="False"></asp:literal>
			<script language="javascript">	
				document.write("<table style='background:#eeeedd;PADDING-LEFT: 4px;FONT: 8pt Arial, Helvetica, sans-serif; POSITION: absolute;TOP: 14px' cellSpacing='0' cellPadding='0' width='125' border='0'><tr><td width='20'>User:</td><td>" + form1.hdLoggedInUserID.value + "</td></tr><tr><td>Site:</td><td WRAP=HARD>" + form1.hdDBHost.value + "</td></tr></table>")
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
				var OKToUpdate
				OKToUpdate = confirm("Warning: If you proceed, you will overwrite any existing license information for the enroller in the state you specify.")
				if (OKToUpdate == true)  {	
					form1.hdAction.value = "update"
					form1.submit()
				} 
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
			
			function Letter(vIn)  {
				form1.hdAction.value = "enrollerletterclick"
				form1.hdSubAction.value = vIn
				form1.submit()
			}	
			
			function ChangeSelection(vIn)  {
				form1.hdAction.value = vIn
				//form1.submit()			
			}
	
			function EnrollerAdd()  {
				form1.hdAction.value = "enrolleradd"
				form1.submit()
			}
			
			function EnrollerRemove()  {
				form1.hdAction.value = "enrollerremove"
				form1.submit()
			}
			
			function GoToAppt()  {
				form1.hdAction.value = "gotoappt"
				form1.submit()
			}			
		
																		
			</script>
		</form>
		<script language="javascript">
			var ApplicationDate = new calendar2(document.forms['form1'].elements['txtApplicationDate']); 
			ApplicationDate.year_scroll = true; ApplicationDate.time_comp = false;
				
			var EffectiveDate = new calendar2(document.forms['form1'].elements['txtEffectiveDate']); 
			EffectiveDate.year_scroll = true; EffectiveDate.time_comp = false;	
			
			var ExpirationDate = new calendar2(document.forms['form1'].elements['txtExpirationDate']); 
			ExpirationDate.year_scroll = true; ExpirationDate.time_comp = false;		
						
			var RenewalDateSent = new calendar2(document.forms['form1'].elements['txtRenewalDateSent']); 
			RenewalDateSent.year_scroll = true; RenewalDateSent.time_comp = false;
				
			//var RenewalDateRecd = new calendar2(document.forms['form1'].elements['txtRenewalDateRecd']); 
			//RenewalDateRecd.year_scroll = true; RenewalDateRecd.time_comp = false;	
			
			//var LongTermCareExpirationDate = new calendar2(document.forms['form1'].elements['txtLongTermCareExpirationDate']); 
			//LongTermCareExpirationDate.year_scroll = true; LongTermCareExpirationDate.time_comp = false;			
			
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
