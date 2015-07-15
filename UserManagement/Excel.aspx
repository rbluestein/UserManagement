<%@ Page Language="vb" AutoEventWireup="false" Codebehind="Excel.aspx.vb" Inherits="UserManagement.Excel" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>Excel</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<asp:DataGrid id="dgRoleID" style="Z-INDEX: 100; LEFT: 128px; POSITION: absolute; TOP: 32px" runat="server"></asp:DataGrid>
			<asp:DataGrid id="dgWorkState" style="Z-INDEX: 106; LEFT: 128px; POSITION: absolute; TOP: 664px"
				runat="server"></asp:DataGrid>
			<asp:DataGrid id="dgResidentState" style="Z-INDEX: 105; LEFT: 128px; POSITION: absolute; TOP: 504px"
				runat="server"></asp:DataGrid>
			<asp:DataGrid id="dgSiteID" style="Z-INDEX: 104; LEFT: 128px; POSITION: absolute; TOP: 336px"
				runat="server"></asp:DataGrid>
			<asp:DataGrid id="dgClientID" style="Z-INDEX: 103; LEFT: 368px; POSITION: absolute; TOP: 32px"
				runat="server"></asp:DataGrid>
			<asp:DataGrid id="dgStatusCode" style="Z-INDEX: 101; LEFT: 128px; POSITION: absolute; TOP: 184px"
				runat="server"></asp:DataGrid>&nbsp;
		</form>
	</body>
</HTML>
