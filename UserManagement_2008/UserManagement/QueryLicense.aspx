<%@ Page Language="VB" AutoEventWireup="false" CodeFile="QueryLicense.aspx.vb" Inherits="QueryLicense" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<title runat="server" id="PageCaption"></title>
        <link href="css/BVI.css" rel="stylesheet" type="text/css" />
        <link href="css/menu.css" rel="stylesheet" type="text/css" />
        <script src="jscripts/menu.js" type="text/javascript"></script>
        <script src="jscripts/MonthPicker.js" type="text/javascript"></script>
        <script src="jscripts/menu_items.js" type="text/javascript"></script>
        <script src="jscripts/menu_tpl.js" type="text/javascript"></script>
	</head>
	<body>
		<form id="form1" runat="server">
			<input type="hidden" name="hdAction" /> <input type="hidden" name="hdSortField" /> <input type="hidden" name="hdUserID"/>
			<input id="hdFilterShowHideToggle" type="hidden" value="0" name="hdFilterShowHideToggle" />
			<table class="PrimaryTbl" style="LEFT: 140px; POSITION: absolute; TOP: 14px" cellspacing="0"
				cellpadding="0" width="1000" border="0">
				<tr style="DISPLAY: none">
					<td width="100"></td>
					<td width="300"></td>
				</tr>
				<tr>
					<td class="PrimaryTblTitle" colspan="2">License Query</td>
				</tr>
				<tr>
					<td class="CellSeparator" colspan="2"></td>
				</tr>
				<tr>
					<td colspan="2"><asp:literal id="litDG" runat="server" enableviewstate="False"></asp:literal></td>
				</tr>
			</table>
			<asp:label id="lblCurrentRights" runat="server"></asp:label><asp:literal id="litEnviro" runat="server" enableviewstate="False"></asp:literal>
			<script type="text/javascript">
				document.write("<table style='background:#eeeedd;PADDING-LEFT: 4px;FONT: 8pt Arial, Helvetica, sans-serif; POSITION: absolute;TOP: 14px' cellSpacing='0' cellPadding='0' width='125' border='0'><tr><td width='20'>User:</td><td>" + form1.hdLoggedInUserID.value + "</td></tr><tr><td>Site:</td><td WRAP=HARD>" + form1.hdDBHost.value + "</td></tr></table>")		
				new menuitems("userlic1", form1.currentrights.value);
				new tpl("v", "admin");
				new menu (MENU_ITEMS, MENU_TPL);				
			</script>
			<asp:literal id="litMsg" runat="server" enableviewstate="False"></asp:literal><asp:literal id="litFilterHiddens" runat="server" enableviewstate="False"></asp:literal>
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
							
					function View()
					{		
						form1.hdAction.value = "ApplyFilter"
						form1.submit()				
					}			
					
					function ApplyFilter()
					{		
						form1.hdAction.value = "ApplyFilter"
						form1.submit()				
					}										
					
					function Download()
					{		
						form1.hdAction.value = "Download"
						form1.submit()				
					}				

					function ExistingRecord(vUserID)
					{
						form1.hdAction.value = "ExistingRecord"
						form1.hdUserID.value = vUserID
						form1.submit()
					}
					function Permissions(vUserID)
					{
						form1.hdAction.value = "Permissions"
						form1.hdUserID.value = vUserID
						form1.submit()
					}	
					
					function License(vUserID)
					{
						form1.hdAction.value = "License"
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
					
					function Delete(vUserID, vUserName)  {
						vUserName = vUserName.replace("~", "'");
						var OKToDelete = confirm("Are you sure you wish to delete " + vUserName + "?");
						if (OKToDelete == true) {
							form1.hdAction.value = "Delete"
							form1.hdUserID.value = vUserID
							form1.submit()		
						}		
					}		
					
					function GetDateRange()  {
							vIn = "MonthPicker"
							eval(vIn).popup()				
					}		
					
																												
			</script>
		</form>
		<script type="text/javascript">
					var MonthPicker = new MonthPicker(document.forms['form1'].elements['hdDateRange']); 
					//var MonthPicker = new MonthPicker(document.forms['form1'].elements['hdDateRange'], document.getElementById("GetDateRange"));
		</script>
	</body>
</html>

