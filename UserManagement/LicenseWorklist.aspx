<%@ Page Language="vb" AutoEventWireup="false" Codebehind="LicenseWorklist.aspx.vb" Inherits="UserManagement.LicenseWorklist" %>
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
		<form id="form1" name="form1" action="LicenseWorklist.aspx" method="post" runat="server">
			<input type="hidden" name="hdAction"> <input type="hidden" name="hdSortField"><input type="hidden" name="hdAppointmentNumber">
			<input type='hidden' name='hdFilterShowHideToggle' id='hdFilterShowHideToggle' value='0'>
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
					<td colSpan="2"><asp:label id="lblEnrollerName" runat="server"></asp:label></td>
				</tr>
				<tr>
					<td colSpan="2"><asp:literal id="litDG" runat="server" EnableViewState="False"></asp:literal></td>
				</tr>
				<tr>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<td colSpan="2"></td>
				</tr>
				<tr>
					<td align="center" colSpan="2"><INPUT onclick="ReturnToParentPage()" type="button" value="Return"></td>
				</tr>
			</table>
			<asp:Literal id="litKeyValues" runat="server" EnableViewState="False"></asp:Literal>
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
				form1.hdAction.value = "update"
				form1.hdUserID.value = vUserID
				form1.submit()
			}									 						
						
			function ExistingLicense(vState, vLicenseNumber, vEffectiveDate)
			{
				form1.hdAction.value = "ExistingLicense"
				form1.hdState.value = vState
				form1.hdLicenseNumber.value = vLicenseNumber
				form1.hdEffectiveDate.value = vEffectiveDate
				form1.submit()
			}
			
			function NewLicense() {
				form1.hdAction.value = "NewLicense"
				form1.submit()
			}	
			
			function NewAppointment(vState) {
				form1.hdState.value = vState
				form1.hdAction.value = "NewAppointment"
				form1.submit()
			}	
			
			function ExistingAppointment(vState, vCarrierID, vEffectiveDate, vAppointmentNumber) {
				form1.hdState.value = vState
				form1.hdCarrierID.value = vCarrierID
				form1.hdEffectiveDate.value = vEffectiveDate
				form1.hdAppointmentNumber.value = vAppointmentNumber			
				form1.hdAction.value = "ExistingAppointment"
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
			
			function ToggleShowFilter()  {
				form1.hdFilterShowHideToggle.value = 1
				form1.hdAction.value = "ApplyFilter"
				form1.submit()
			}				
			
			function ShowHideSubTable(vSubTableInd, vSubTableState)  {
				form1.hdSubTableInd.value = vSubTableInd
				form1.hdSubTableState.value = vSubTableState
				form1.hdAction.value = "ShowHideSubTable"
				form1.submit()
			}				
	
	/*	
			function ShowHideSubTable(vSubTableInd, vSubTableState, vLicenseNumber, vEffectiveDate)  {
				form1.hdSubTableInd.value = vSubTableInd
				form1.hdSubTableState.value = vSubTableState
				form1.hdLicenseNumber.value = vLicenseNumber
				form1.hdEffectiveDate.value = vEffectiveDate
				form1.hdAction.value = "ShowHideSubTable"
				form1.submit()
			}	
	
	*/		
			function DeleteLic(vState, vLicenseNumber, vEffectiveDate)  {
				var OKToDelete = confirm("Deleting this license deletes the license and any appointments associated with this license. Do you wish to proceed with the deletion?");
				if (OKToDelete == true) {
					form1.hdAction.value = "DeleteLic"
					form1.hdState.value = vState
					form1.hdLicenseNumber.value = vLicenseNumber
					form1.hdEffectiveDate.value = vEffectiveDate					
					form1.submit()		
				}		
			}
			
			function DeleteAppt(vState, vCarrierID, vEffectiveDate, vAppointmentNumber)  {
				var OKToDelete = confirm("Are you sure you wish to delete this appointment?");
				if (OKToDelete == true) {
					form1.hdAction.value = "DeleteAppt"
					form1.hdState.value = vState
					form1.hdCarrierID.value = vCarrierID
					form1.hdEffectiveDate.value = vEffectiveDate	
					form1.hdAppointmentNumber.value = vAppointmentNumber								
					form1.submit()		
				}		
			}								
																									
			</script>
		</form>
	</body>
</HTML>
