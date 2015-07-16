<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ClientWorklist.aspx.vb" Inherits="ClientWorklist" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<title id="PageCaption" runat="server"></title>					
        <link href="css/BVI.css" rel="stylesheet" type="text/css" />
        <link href="css/menu.css" rel="stylesheet" type="text/css" />
        <script src="jscripts/menu_items.js" type="text/javascript"></script>
        <script src="jscripts/menu.js" type="text/javascript"></script>
        <script src="jscripts/menu_tpl.js" type="text/javascript"></script>		
	</head>	
	<body>
		<form id="form1" name="form1" runat="server">
			<input type="hidden" name="hdAction" /> <input type="hidden" name="hdSortField" /> <input type="hidden" name="hdClientID" />
			<input type='hidden' name='hdFilterShowHideToggle' id='hdFilterShowHideToggle' value='0' />
			<table class="PrimaryTbl" style="LEFT: 140px; POSITION: absolute; TOP: 14px" cellspacing="0"
				cellpadding="0" width="650" border="0">
				<tr style="DISPLAY: none">
					<td width="100"></td>
					<td width="300"></td>
				</tr>
				<tr>
					<td class="PrimaryTblTitle" colspan="2">Clients</td>
				</tr>
				<tr>
					<td class="CellSeparator" colspan="2"></td>
				</tr>
				<tr>
					<td colspan="2"><asp:literal id="litDG" runat="server" enableviewstate="False"></asp:literal></td>
				</tr>
			</table>
			<asp:label id="lblCurrentRights" runat="server"></asp:label>
			<asp:literal id="litEnviro" runat="server" enableviewstate="False"></asp:literal>
			<script type="text/javascript">				
				//document.write("<table style='background:#eeeedd;PADDING-LEFT: 4px;FONT: 8pt Arial, Helvetica, sans-serif; POSITION: absolute;TOP: 14px' cellspacing='0' cellpadding='0' width='125' border='0'><tr><td width='20'>User:</td><td>" + form1.hdLoggedInUserID.value + "</td></tr><tr><td>Site:</td><td WRAP=HARD>" + form1.hdDBHost.value + "</td></tr></table>")
			    document.write("<table style='background:#eeeedd;PADDING-LEFT: 4px;FONT: 8pt Arial, Helvetica, sans-serif; POSITION: absolute;TOP: 14px' cellSpacing='0' cellPadding='0' width='125' border='0'><tr><td width='20'>User:</td><td>" + document.getElementById("hdLoggedInUserID").value + "</td></tr><tr><td>Site:</td><td WRAP=HARD>" + document.getElementById("hdDBHost").value + "</td></tr></table>")			
				
				//new menuitems("userlic1", form1.currentrights.value);
			    new menuitems("userlic1", document.getElementById("currentrights").value);
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
	
			function ApplyFilter()
			{		
				form1.hdAction.value = "ApplyFilter"
				form1.submit()				
			}						 						
						
			function ExistingRecord(vClientID)
			{
				form1.hdAction.value = "ExistingRecord"
				form1.hdClientID.value = vClientID
				form1.submit()
			}
			
			function Delete(vClientID)  {
				vClientID = vClientID.replace("~", "'");
				var OKToDelete = confirm("Are you sure you wish to delete " + vClientID + "?");
				if (OKToDelete == true) {
					form1.hdAction.value = "Delete"
					form1.hdClientID.value = vClientID
					form1.submit()		
				}		
			}	
			
			/*
			function License(vClientID)
			{
				form1.hdAction.value = "License"
				form1.hdClientID.value = vClientID
				form1.submit()
			}
			*/						
			
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
																												
			</script>
		</form>
	</body>
</html>

