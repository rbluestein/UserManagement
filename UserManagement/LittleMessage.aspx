<%@ Page Language="vb" AutoEventWireup="false" Codebehind="LittleMessage.aspx.vb" Inherits="UserManagement.LittleMessage"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>LittleMessage</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<LINK title="BVIStyle" href="BVI.css" type="text/css" rel="stylesheet">
	</HEAD>
	<body onload="fnLoad()">
		<form id="form1" method="post" runat="server">
			<table class="PrimaryTbl" style="LEFT: 0px; POSITION: absolute; TOP: 0px" cellSpacing="0" cellPadding="0" width="650" border="0">
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td>
						<asp:Label id="lblMessage" runat="server">Label</asp:Label></td>
				</tr>
				<tr height="200px"><td>&nbsp;</td></tr>
			</table>
			
			<script language="javascript">
				function fnLoad()  {
				//	window.setTimeout("self.close()", 5000 ) 
					window.setTimeout("window.open('','_parent','');window.close();", 3000)
				}
			</script>
		</form>
	</body>
</HTML>
