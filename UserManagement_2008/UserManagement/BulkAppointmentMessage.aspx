<%@ Page Language="VB" AutoEventWireup="false" CodeFile="BulkAppointmentMessage.aspx.vb" Inherits="BulkAppointmentMessage" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title runat="server" id="PageCaption"></title>
    <link href="css/BVI.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
    <div>
			<table class="PrimaryTbl" style="LEFT: 30px; POSITION: absolute; TOP: 14px" cellspacing="0"
				cellpadding="0" width="500" border="0">
				<tr style="DISPLAY: none">
					<td width="100"></td>
					<td width="300"></td>
				</tr>
				<tr>
					<td class="PrimaryTblTitle" colspan="2">Save Results</td>
				</tr>
				<tr>
					<td class="CellSeparator" colspan="2"></td>
				</tr>
				<tr>
					<td colspan="2"><asp:literal id="litDG" runat="server" EnableViewState="False"></asp:literal></td>
				</tr>
				<tr>
					<td colspan="2">&nbsp;</td>
				</tr>				
				<tr>
					<td colspan="2" align="center"><input type="button" value="Close" onclick="CloseWindow()" /></td>
				</tr>
			</table>  
			<script type="text/javascript"> 			
		
					//This opens a new page, (non-existent), into a target frame/window, (_parent which of course is the window in which the script is executed, so replacing itself), and defines parameters such as window size etc, (in this case none are defined as none are needed). Now that the browser thinks a script opened a page we can quickly close it in the standard way…
					function CloseWindow()  {
						window.open('','_parent','');
						window.close();			
					}	
																												
			</script>			  
    
    </div>
    </form>
</body>
</html>
