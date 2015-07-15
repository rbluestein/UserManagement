<%@ Page Language="VB" AutoEventWireup="false" CodeFile="LittleMessage.aspx.vb" Inherits="LittleMessage" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<title>LittleMessage</title>
        <link href="css/BVI.css" rel="stylesheet" type="text/css" />
	</head>
	<body onload="fnLoad()">
		<form id="form1" method="post" runat="server">
			<table class="PrimaryTbl" style="LEFT: 0px; POSITION: absolute; TOP: 0px" cellspacing="0" cellpadding="0" width="650" border="0">
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td>
						<asp:Label id="lblMessage" runat="server">Label</asp:Label></td>
				</tr>
				<tr height="200px"><td>&nbsp;</td></tr>
			</table>
			
			<script type="text/javascript">
				function fnLoad()  {
				//	window.setTimeout("self.close()", 5000 ) 
					window.setTimeout("window.open('','_parent','');window.close();", 3000)
				}
			</script>
		</form>
	</body>
</html>