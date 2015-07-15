<%@ Page Language="VB" AutoEventWireup="false" CodeFile="LicenseWorklist.aspx.vb" Inherits="LicenseWorklist" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<title runat="server" id="PageCaption"></title>
        <link href="css/BVI.css" rel="stylesheet" type="text/css" />
        <link href="css/menu.css" rel="stylesheet" type="text/css" />
        <script src="jscripts/menu.js" type="text/javascript"></script>
        <script src="jscripts/menu_items.js" type="text/javascript"></script>
        <script src="jscripts/menu_tpl.js" type="text/javascript"></script>
	</head>
	<body>
		<form id="form2" runat="server">
			<input type="hidden" name="hdAction" /> <input type="hidden" name="hdSortField" /><input type="hidden" name="hdAppointmentNumber" />
			<input type='hidden' name='hdFilterShowHideToggle' id='hdFilterShowHideToggle' value='0' />
			<table class="PrimaryTbl" style="LEFT: 140px; POSITION: absolute; TOP: 14px" cellspacing="0"
				cellpadding="0" width="650" border="0">
				<tr>
					<td class="PrimaryTblTitle" colspan="2">
						<asp:literal id="litHeading" runat="server"></asp:literal></td>
				</tr>
				<tr>
					<td class="CellSeparator" colspan="2"></td>
				</tr>
				<tr>
					<td colspan="2"><asp:label id="lblEnrollerName" runat="server"></asp:label></td>
				</tr>
				<tr>
					<td colspan="2"><asp:literal id="litDG" runat="server" EnableViewState="False"></asp:literal></td>
				</tr>
				<tr>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<td colspan="2"></td>
				</tr>
				<tr>
					<td align="center" colspan="2"><input onclick="ReturnToParentPage()" type="button" value="Return" /></td>
				</tr>
			</table>
			<asp:Literal id="litKeyValues" runat="server" EnableViewState="False"></asp:Literal>
			<asp:label id="lblCurrentRights" runat="server"></asp:label>
			<asp:Literal id="litEnviro" runat="server" EnableViewState="False"></asp:Literal>			
			<script type="text/javascript">
				document.write("<table style='background:#eeeedd;PADDING-LEFT: 4px;FONT: 8pt Arial, Helvetica, sans-serif; POSITION: absolute;TOP: 14px' cellSpacing='0' cellPadding='0' width='125' border='0'><tr><td width='20'>User:</td><td>" + form1.hdLoggedInUserID.value + "</td></tr><tr><td>Site:</td><td WRAP=HARD>" + form1.hdDBHost.value + "</td></tr></table>")			
							
				new menuitems("userlic1", form1.currentrights.value);
				new tpl("v", "admin");
				new menu (MENU_ITEMS, MENU_TPL);				
			</script>
			<asp:literal id="litMsg" runat="server" EnableViewState="False"></asp:literal><asp:literal id="litFilterHiddens" runat="server" EnableViewState="False"></asp:literal>
			<script type="text/javascript"> 
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
</html>

