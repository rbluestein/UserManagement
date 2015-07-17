<%@ Page Language="VB" AutoEventWireup="false" CodeFile="UserMaintain.aspx.vb" Inherits="UserMaintain" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">

	<head>
		<title id="PageCaption" runat="server"></title>
        <link href="css/BVI.css" rel="stylesheet" type="text/css" />
        <link href="css/menu.css" rel="stylesheet" type="text/css" />
        <script src="jscripts/calendar2.js" type="text/javascript"></script>
        <script src="jscripts/menu_tpl.js" type="text/javascript"></script>
        <script src="jscripts/menu.js" type="text/javascript"></script>
        <script src="jscripts/menu_items.js" type="text/javascript"></script>
	</head>
	<body>
		<form id="form1" runat="server">
			<input type="hidden" id="hdAction" name="hdAction" /><input type="hidden" id="hdSubAction" name="hdSubAction" /><input type="hidden" name="hdTermDate" />
			<asp:literal id="litHiddens" runat="server" enableviewstate="False"></asp:literal>
			<table style="POSITION: absolute; TOP: 14px; LEFT: 140px" class="PrimaryTbl" border="0"
				cellspacing="0" cellpadding="0" width="650">
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
					<td colspan="2"><a href="javascript:fnToggleSection('PersonalData')"><b>Personal Data</b></a></td>
				</tr>
				<asp:placeholder id="plPersonalData" runat="server">
					<tr>
						<td class="Cell9Reg" width="120" align="left">First Name:</td>
						<td>
							<asp:textbox id="txtFirstName" runat="server" width="300px" maxlength="15"></asp:textbox></td>
					</tr>
					<tr>
						<td class="Cell9Reg" width="120" align="left">Middle Initial:</td>
						<td>
							<asp:textbox id="txtMI" runat="server" width="300px" maxlength="1"></asp:textbox></td>
					</tr>
					<tr>
						<td class="Cell9Reg" width="120" align="left">Last Name:</td>
						<td>
							<asp:textbox id="txtLastName" runat="server" width="300px" maxlength="30"></asp:textbox></td>
					</tr>
					<tr>
						<td class="Cell9Reg" width="120" align="left">Primary Phone #:</td>
						<td>
							<asp:textbox id="txtPrimaryContactNumber" runat="server" width="300px" maxlength="15"></asp:textbox></td>
					</tr>
					<tr>
						<td class="Cell9Reg" width="120" align="left">Extension #:</td>
						<td>
							<asp:textbox id="txtPhoneExtension" runat="server" width="300px" maxlength="20"></asp:textbox></td>
					</tr>
					<tr>
						<td class="Cell9Reg" width="120" align="left">Alternate Phone #:</td>
						<td>
							<asp:textbox id="txtAltContactNumber" runat="server" width="300px" maxlength="15"></asp:textbox></td>
					</tr>
					<tr>
						<td class="Cell9Reg" width="120" align="left">Email:</td>
						<td>
							<asp:textbox id="txtEmail" runat="server" width="300px" maxlength="50"></asp:textbox></td>
					</tr>
					<tr>
						<td class="Cell9Reg" width="120" align="left">Resident State:</td>
						<td>
							<asp:dropdownlist id="ddResidentState" runat="server"></asp:dropdownlist>
							<asp:textbox id="txtResidentState" runat="server" width="300px"></asp:textbox></td>
					</tr>
				</asp:placeholder>
				<tr>
					<td colspan="2">&nbsp;</td>
				</tr>
				<tr>
					<td colspan="2"><a href="javascript:fnToggleSection('EmploymentData')"><b>Employment Data</b></a></td>
				</tr>
				<asp:placeholder id="plEmploymentData" runat="server">
					<tr>
						<td class="Cell9Reg" width="120" align="left">User ID:</td>
						<td>
							<asp:textbox id="txtUserID" runat="server" width="300px"></asp:textbox></td>
					</tr>
					<tr>
						<td class="Cell9Reg" width="120" align="left">Status:</td>
						<td>
							<asp:dropdownlist id="ddStatusCode" runat="server"></asp:dropdownlist>
							<asp:textbox id="txtStatusCode" runat="server" width="300px"></asp:textbox></td>
					</tr>
					<tr>
						<td colspan="2">&nbsp;</td>
					</tr>
					<tr>
						<td class="Cell9Reg" colspan="2" align="left"><u>BVI User</u></td>
					</tr>
					<tr>
						<td class="Cell9Reg" width="120" align="left">BVI</td>
						<td>
							<asp:checkbox id="chkBVI" runat="server"></asp:checkbox>
							<asp:textbox id="txtBVI" runat="server" width="300px"></asp:textbox></td>
					</tr>
					<tr>
						<td class="Cell9Reg" width="120" align="left">Location:</td>
						<td>
							<asp:dropdownlist id="ddLocationID" runat="server"></asp:dropdownlist>
							<asp:textbox id="txtLocationID" runat="server" width="300px"></asp:textbox></td>
					</tr>
					<tr>
						<td class="Cell9Reg" width="120" align="left">Role:</td>
						<td>
							<asp:dropdownlist id="ddRole" runat="server"></asp:dropdownlist>
							<asp:textbox id="txtRole" runat="server" width="300px"></asp:textbox></td>
					</tr>
					<tr>
						<td class="Cell9Reg" width="120" align="left">Work State:</td>
						<td>
							<asp:dropdownlist id="ddWorkState" runat="server"></asp:dropdownlist>
							<asp:textbox id="txtWorkState" runat="server" width="300px"></asp:textbox></td>
					</tr>
					<tr>
						<td colspan="2">&nbsp;</td>
					</tr>
					<tr>
						<td class="Cell9Reg" colspan="2" align="left"><u>Non-BVI User</u></td>
					</tr>
					<tr>
						<td class="Cell9Reg" width="120" align="left">Other</td>
						<td>
							<asp:checkbox id="chkOther" runat="server"></asp:checkbox>
							<asp:textbox id="txtOther" runat="server" width="300px"></asp:textbox></td>
					</tr>
					<tr>
						<td class="Cell9Reg" width="120" align="left">Company:</td>
						<td>
							<asp:dropdownlist id="ddCompanyID" runat="server"></asp:dropdownlist>
							<asp:textbox id="txtCompanyID" runat="server" width="300px"></asp:textbox></td>
					</tr>
				</asp:placeholder>
				<tr>
					<td colspan="2">&nbsp;</td>
				</tr>
				<tr>
					<td><a href="javascript:fnToggleSection('Enroller')"><b>Enroller Details</b></a></td>
				</tr>
				<asp:placeholder id="plEnrollerSection" runat="server">
					<tr>
						<td class="Cell9Reg" width="120" align="left">National Producer #:</td>
						<td>
							<asp:textbox id="txtNationalProducerNumber" runat="server" width="300px"></asp:textbox></td>
					</tr>
					<tr>
						<td style="WIDTH: 120px" class="Cell9Reg" align="left">LTC State:</td>
						<td>
							<asp:dropdownlist id="ddLongTermCareCertState" runat="server"></asp:dropdownlist>
							<asp:textbox id="txtLongTermCareCertState" runat="server" width="300px"></asp:textbox></td>
					</tr>
					<tr>
						<td style="WIDTH: 120px" class="Cell9Reg" align="left">LTC Eff Date:</td>
						<td class="Cell9Reg">
							<asp:textbox id="txtLongTermCareCertEffectiveDate" runat="server" width="112px" maxlength="50"
								readonly="True"></asp:textbox>&nbsp;&nbsp;
							<asp:label id="lblLongTermCareCertEffectiveDateLink" runat="server"><a href="javascript:GetDate('LongTermCareCertEffectiveDate')">
									Get Date</a>&nbsp;&nbsp</asp:label></td>
					</tr>
					<tr>
						<td style="WIDTH: 120px" class="Cell9Reg" align="left">LTC Exp Date:</td>
						<td class="Cell9Reg">
							<asp:textbox id="txtLongTermCareCertExpirationDate" runat="server" width="112px" maxlength="50"
								readonly="True"></asp:textbox>&nbsp;&nbsp;
							<asp:label id="lblLongTermCareCertExpirationDateLink" runat="server"><a href="javascript:GetDate('LongTermCareCertExpirationDate')">
									Get Date</a>&nbsp;&nbsp</asp:label></td>
					</tr>
					<tr>
						<td colspan="2">&nbsp;</td>
					</tr>
<%--					<tr>
						<td class="Cell9Reg" colspan="2" align="left"><u>Falcon</u></td>
					</tr>					
					<tr>
						<td class="Cell9Reg" colspan="2"><asp:literal id="litFalcon" runat="server"  enableviewstate="True"></asp:literal></td>
					</tr>--%>
				</asp:placeholder>
				<tr>
					<td style="WIDTH: 555px" colspan="2">&nbsp;</td>
				</tr>
				<tr>
					<td class="Cell9Reg" width="120" align="left">&nbsp;</td>
					<td align="left"><asp:literal id="litUpdate" runat="server" enableviewstate="False"></asp:literal>&nbsp;&nbsp;<input onclick="ReturnToParentPage()" value="Return" type="button" /></td>
				</tr>
				<tr>
					<td style="WIDTH: 555px" colspan="2">&nbsp;</td>
				</tr>
			</table>
			<asp:literal id="litResponseAction" runat="server" enableviewstate="False"></asp:literal><asp:label id="lblCurrentRights" runat="server"></asp:label>
			<script type="text/javascript">
				var TermDate = new calendar2(document.forms['form1'].elements['hdTermDate']); 
				TermDate.year_scroll = true; TermDate.time_comp = false;				
			</script>
			<asp:literal id="litEnviro" runat="server" enableviewstate="False"></asp:literal>
			<script type="text/javascript">
			    //document.write("<table style='background:#eeeedd;PADDING-LEFT: 4px;FONT: 8pt Arial, Helvetica, sans-serif; POSITION: absolute;TOP: 14px' cellSpacing='0' cellPadding='0' width='125' border='0'><tr><td width='20'>User:</td><td>" + document.getElementById("hdLoggedInUserID").value + "</td></tr><tr><td>Site:</td><td WRAP=HARD>" + document.getElementById("hdHost").value + "</td></tr></table>")
			    document.write("<table style='background:#eeeedd;PADDING-LEFT: 4px;FONT: 8pt Arial, Helvetica, sans-serif; POSITION: absolute;TOP: 14px' cellSpacing='0' cellPadding='0' width='125' border='0'><tr><td width='20'>User:</td><td>" + document.getElementById("hdLoggedInUserID").value + "</td></tr><tr><td>Site:</td><td WRAP=HARD>" + document.getElementById("hdDBHost").value + "</td></tr></table>")			
				
				new menuitems("userlic1", document.getElementById("currentrights").value);
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
			
			function HandleTermDate()  {	
				//if(form1.ddStatusCode.value == "TERMINATED" && (form1.ddRole.value == "ENROLLER" || form1.ddRole.value == "SUPERVISOR"))  {
			    if (document.getElementById("ddStatusCode").value == "TERMINATED" && (document.getElementById("ddRole").value == "ENROLLER" || document.getElementById("ddRole").value == "SUPERVISOR")) {
					eval("TermDate").popup()			
				}
			}
			
			function GetDate(vIn)  {
				eval(vIn).popup()
			}
			
			function fnToggleSection(vIn)  {
			    document.getElementById("hdAction").value = "clientselectionchanged"
			    document.getElementById("hdSubAction").value = vIn
				form1.submit()
			}				
		
		
			function fnAlphaNumericOnly(txtbox)
			{
					//var txtbox = document.getElementByID("adriano")
					var ToTest = txtbox.value
					var re = /[^0-9A-Za-z]/
					if (re.test(ToTest)) {
						txtbox.value = txtbox.value.substring(0, txtbox.value.length-1)
					}
			}				
			
			</script>
		</form>
		<script type="text/javascript">
			if (document.getElementById("txtLongTermCareCertEffectiveDate") != null)  {
				var LongTermCareCertEffectiveDate = new calendar2(document.forms['form1'].elements['txtLongTermCareCertEffectiveDate']); 
				LongTermCareCertEffectiveDate.year_scroll = true; LongTermCareCertEffectiveDate.time_comp = false;
					
				var LongTermCareCertExpirationDate = new calendar2(document.forms['form1'].elements['txtLongTermCareCertExpirationDate']); 
				LongTermCareCertExpirationDate.year_scroll = true; LongTermCareCertExpirationDate.time_comp = false;	
			}
		</script>
	</body>
</html>
