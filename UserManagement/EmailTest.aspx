<%@ Page Language="vb" AutoEventWireup="false" Codebehind="EmailTest.aspx.vb" Inherits="UserManagement.EmailTest" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>SendEmail</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<asp:Label id="lblTo" style="Z-INDEX: 100; LEFT: 40px; POSITION: absolute; TOP: 32px" runat="server">To</asp:Label>
			<asp:Label id="lblResults" style="Z-INDEX: 109; LEFT: 24px; POSITION: absolute; TOP: 320px"
				runat="server">Results</asp:Label>
			<asp:TextBox id="txtResults" style="Z-INDEX: 108; LEFT: 104px; POSITION: absolute; TOP: 320px"
				runat="server" Width="320px" TextMode="MultiLine" Height="120px" Font-Names="Arial"></asp:TextBox>
			<asp:Label id="lblAttachment" style="Z-INDEX: 105; LEFT: 24px; POSITION: absolute; TOP: 208px"
				runat="server">Attachment</asp:Label>
			<asp:Label id="lblBody" style="Z-INDEX: 103; LEFT: 40px; POSITION: absolute; TOP: 72px" runat="server">From</asp:Label>
			<asp:TextBox id="txtTo" style="Z-INDEX: 101; LEFT: 112px; POSITION: absolute; TOP: 32px" runat="server"
				Width="312px"></asp:TextBox>
			<asp:TextBox id="txtBody" style="Z-INDEX: 102; LEFT: 112px; POSITION: absolute; TOP: 72px" runat="server"
				Width="320px" TextMode="MultiLine" Height="120px" Font-Names="Arial"></asp:TextBox>
			<asp:CheckBox id="chkAttachment" style="Z-INDEX: 104; LEFT: 104px; POSITION: absolute; TOP: 208px"
				runat="server"></asp:CheckBox>
			<asp:Button id="btnSendMail" style="Z-INDEX: 107; LEFT: 104px; POSITION: absolute; TOP: 256px"
				runat="server" Text="Send Mail"></asp:Button></form>
	</body>
</HTML>
