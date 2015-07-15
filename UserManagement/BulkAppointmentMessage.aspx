<%@ Page Language="vb" AutoEventWireup="false" Codebehind="BulkAppointmentMessage.aspx.vb" Inherits="UserManagement.BulkAppointmentMessage"%>
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
  </HEAD>
	<body>
		<form id="form1" name="form1" action="BulkAppointmentMessge.aspx" method="post" runat="server">
			<table class="PrimaryTbl" style="LEFT: 30px; POSITION: absolute; TOP: 14px" cellSpacing="0"
				cellPadding="0" width="500" border="0">
				<tr style="DISPLAY: none">
					<td width="100"></td>
					<td width="300"></td>
				</tr>
				<tr>
					<td class="PrimaryTblTitle" colSpan="2">Save Results</td>
				</tr>
				<tr>
					<td class="CellSeparator" colSpan="2"></td>
				</tr>
				<tr>
					<td colSpan="2"><asp:literal id="litDG" runat="server" EnableViewState="False"></asp:literal></td>
				</tr>
				<tr>
					<td colSpan="2">&nbsp;</td>
				</tr>				
				<tr>
					<td colSpan="2" align="center"><INPUT type="button" value="Close" onclick="CloseWindow()"></td>
				</tr>
			</table>
			<script language="javascript"> 			
		
					//This opens a new page, (non-existent), into a target frame/window, (_parent which of course is the window in which the script is executed, so replacing itself), and defines parameters such as window size etc, (in this case none are defined as none are needed). Now that the browser thinks a script opened a page we can quickly close it in the standard way…
					function CloseWindow()  {
						window.open('','_parent','');
						window.close();			
					}	
																												
			</script>
		</form>
	</body>
</HTML>