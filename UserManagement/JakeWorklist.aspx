<%@ Page Language="vb" AutoEventWireup="false" Codebehind="JakeWorklist.aspx.vb" Inherits="UserManagement.JakeWorklist" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>Jake</title>
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
	<BODY>
		<form id="form1" name="form1" action="JakeWorklist.aspx" method="post" runat="server">
			<input type="hidden" name="hdAction"> <input type="hidden" name="hdSortField"> <input type="hidden" name="hdJakeId">
			<table class="PrimaryTbl" style="LEFT: 140px; POSITION: absolute; TOP: 14px" cellSpacing="0"
				cellPadding="0" width="650" border="0">
				<tr style="DISPLAY: none">
					<td width="100"></td>
					<td width="300"></td>
				</tr>
				<tr>
					<td class="PrimaryTblTitle" colSpan="2">Jake
					</td>
				</tr>
				<tr>
					<td class="CellSeparator" colSpan="2"></td>
				</tr>
				<tr>
					<td colSpan="2"><asp:literal id="litDG" runat="server"></asp:literal></td>
				</tr>
			</table>
			<asp:label id="lblCurrentRights" runat="server"></asp:label>
			<asp:Literal id="litFilterHiddens" runat="server" EnableViewState="False"></asp:Literal>
			<script language='javascript'>				
				new menuitems("userlic1", form1.currentrights.value);
				new tpl("v", "admin");
				new menu (MENU_ITEMS, MENU_TPL);				
			</script>
			<script language="javascript"> 
				function Sort(vField) {
				form1.hdAction.value = "Sort"
				form1.hdSortField.value = vField
				form1.submit()
			}
	
			function ApplyFilter()
			{		
				form1.hdAction.value = "ApplyFilter"
				form1.submit()				
			}			
			
			function ToggleShowFilter()  {
				form1.hdFilterShowHideToggle.value = 1
				form1.hdAction.value = "ApplyFilter"
				form1.submit()
			}							 						
						
			function View(vJakeId)
			{
				form1.hdAction.value = "View"
				form1.hdJakeId.value = vJakeId
				form1.submit()
			}
			
			function GoToForm() {
				form1.hdAction.value = "GoToForm"
				form1.submit()
			}	
			
			function ExistingRecord(vUserID)
			{
				form1.hdAction.value = "ExistingRecord"
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
																								
			</script>
		</form>
	</BODY>
</HTML>
