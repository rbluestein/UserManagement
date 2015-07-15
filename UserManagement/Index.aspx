<%@ Page Language="vb" AutoEventWireup="false" Codebehind="Index.aspx.vb" Inherits="UserManagement.Index" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
		<title runat="server" id="PageCaption"></title>
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
		<meta content="Microsoft Visual Studio.NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<link title="BVIStyle" href="BVI.css" type="text/css" rel="stylesheet">
		<link href="menu.css" rel="stylesheet">
		<script language="JavaScript" src="menu.js"></script>
		<script language="JavaScript" src="menu_items.js"></script>
		<script language="JavaScript" src="menu_tpl.js"></script>
	</head>
	<body>
		<form id="form1" name="form1" action="Index.aspx" method="post" runat="server">
			<input type="hidden" name="hdAction"> <input type="hidden" name="hdSortField"> <input type="hidden" name="hdUserID">
			<input type='hidden' name='hdFilterShowHideToggle' id='hdFilterShowHideToggle' value='0'>
			<table class="PrimaryTbl" style="POSITION: absolute; TOP: 14px; LEFT: 140px" cellspacing="0"
				cellpadding="0" width="650" border="0">
				<tr style="DISPLAY: none">
					<td width="100"></td>
					<td width="300"></td>
				</tr>
				<tr>
					<td colspan="2" class="PrimaryTblTitle">Users</td>
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
			<script language="javascript">
				//document.write("<table style='background:#eeeedd;PADDING-LEFT: 4px;FONT: 8pt Arial, Helvetica, sans-serif; POSITION: absolute;TOP: 14px' cellSpacing='0' cellPadding='0' width='125' border='0'><tr><td width='20'>User:</td><td wrap>awickenheiseranddaisytoo</td></tr><tr><td>Site:</td><td  WRAP=HARD>" + form1.hdDBHost.value + "</td></tr></table>")
				document.write("<table style='background:#eeeedd;PADDING-LEFT: 4px;FONT: 8pt Arial, Helvetica, sans-serif; POSITION: absolute;TOP: 14px' cellSpacing='0' cellPadding='0' width='125' border='0'><tr><td width='20'>User:</td><td>" + form1.hdLoggedInUserID.value + "</td></tr><tr><td>Site:</td><td WRAP=HARD>" + form1.hdDBHost.value + "</td></tr></table>")
			
				new menuitems("userlic1", form1.currentrights.value);
				new tpl("v", "admin");
				new menu (MENU_ITEMS, MENU_TPL);
			</script>
			<asp:literal id="litMsg" runat="server" enableviewstate="False"></asp:literal><asp:literal id="litFilterHiddens" runat="server" enableviewstate="False"></asp:literal>
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
	/*
	  function Test2()
      {
         //alert(arguments.Value);
         var results
         var re      //Declare variables.
         var source = "94602"
         
         re = new RegExp("/^[1-9][0-9][0-9][0-9]$/");  //Create regular expression object.
         results = re.test(source);  //Test for match.
         alert(results);    //Return results.
      }


	function fnTest3()
	{
		var telnum = "94602"
		exp = "/^[0-9]$/"
		if (exp.test(telnum) != true)
		{
			alert("Please enter only numbers with area code - For example: 01382000000");
		}
	}	
*/

			function ApplyFilter()
			{		
				form1.hdAction.value = "ApplyFilter"
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
																												
			</script>
		</form>
	</body>
</html>
