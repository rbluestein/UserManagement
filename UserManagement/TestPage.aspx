<%@ Page Language="vb" AutoEventWireup="false" Codebehind="TestPage.aspx.vb" Inherits="UserManagement.TestPage" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>TestPage</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
				<link title="BVIStyle" href="BVI.css" type="text/css" rel="stylesheet">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<input type="hidden" name="myHidden">
			<table>
				<tr>
					<td class="DGMenuItem"width="14.9%">hello</td>
					<td><a id="daisy" name="daisy" href="http://www.google.com" title="thedog">Google</a></td>
					<td><input type="button" onclick="fnTestIT()"</td>
				</tr>
			</table>
			<asp:Label id="Label1" style="Z-INDEX: 101; LEFT: 176px; POSITION: absolute; TOP: 200px" runat="server"
				Width="232px" Height="24px">Label</asp:Label>
			<asp:Label id="Label2" style="Z-INDEX: 102; LEFT: 176px; POSITION: absolute; TOP: 248px" runat="server">Label</asp:Label>
			<asp:Label id="Label3" style="Z-INDEX: 103; LEFT: 176px; POSITION: absolute; TOP: 296px" runat="server">Label</asp:Label>

		<script language="javascript">
			function fnTestIT()  {
				Form1.myHidden.value = "whoisadog"
				var x=document.getElementById("daisy")
				x.title="Good morning"
			}
		</script>
				</form>
	</body>
</HTML>
