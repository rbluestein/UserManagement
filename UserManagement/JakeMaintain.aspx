<%@ Page Language="vb" AutoEventWireup="false" Codebehind="JakeMaintain.aspx.vb" Inherits="UserManagement.JakeMaintain" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>Jake</title>
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
	<body>
		<form id="form1" name="form1" action="JakeMaintain.aspx" method="post" runat="server">
			<input type="hidden" name="hdAction">
			<asp:Literal id="litHiddens" runat="server" EnableViewState="False"></asp:Literal>
			<table class="PrimaryTbl" style="LEFT: 140px; POSITION: absolute; TOP: 14px" cellSpacing="0"
				cellPadding="0" width="650" border="0">
				<tr style="DISPLAY: none">
					<td width="150"></td>
					<td width="443">&nbsp;</td>
					<td width="150">;</td>
				</tr>
				<tr>
					<td class="PrimaryTblTitle" colSpan="3"><asp:literal id="litHeading" runat="server"></asp:literal></td>
				</tr>
				<tr>
					<td class="CellSeparator" colSpan="3"></td>
				</tr>
				<tr height="24">
					<td class="Cell9Reg" align="left">Ticket Number:</td>
					<td>
						<asp:textbox id="txtTicketNum" runat="server" Width="304px"></asp:textbox></td>
				</tr>
				<tr height="24">
					<td class="Cell9Reg" align="left">Open Date:</td>
					<td>
						<asp:textbox id="txtOpenDt" runat="server" Width="304px"></asp:textbox></td>
				</tr>
				<tr height="24">
					<td class="Cell9Reg" align="left">Close Date:</td>
					<td>
						<asp:textbox id="txtCloseDt" runat="server" Width="304px"></asp:textbox></td>
				</tr>
				<tr>
					<td>&nbsp;</td>
				</tr>
				<tr height="24">
					<td class="Cell9Reg" align="left">Originator:</td>
					<td style="WIDTH: 443px"><asp:textbox id="txtOriginator" runat="server" Width="304px"></asp:textbox></td>
				</tr>
				<tr>
					<td class="Cell9Reg" align="left">Summary:</td>
					<td style="WIDTH: 443px"><asp:textbox id="txtIssueSummary" runat="server" MaxLength="100" Width="400px"></asp:textbox></td>
				</tr>
				<tr>
					<td class="Cell9Reg" vAlign="top" align="left">Description:</td>
					<td style="WIDTH: 443px"><asp:textbox id="txtIssueDescription" runat="server" Width="445px" TextMode="MultiLine" Height="100px"
							Font-Names="Arial"></asp:textbox></td>
				</tr>
				<tr>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<td class="Cell9Reg" align="left">Tech Status:</td>
					<td style="WIDTH: 443px"><asp:dropdownlist id="ddTechStatus" runat="server">
							<asp:ListItem Value="1">Not started</asp:ListItem>
							<asp:ListItem Value="2">In progress</asp:ListItem>
							<asp:ListItem Value="3">Unable to duplicate</asp:ListItem>
							<asp:ListItem Value="4">Completed</asp:ListItem>
						</asp:dropdownlist><asp:textbox id="txtTechStatus" runat="server"></asp:textbox></td>
				</tr>
				<tr>
					<td class="Cell9Reg" vAlign="top" align="left">Tech Comments:</td>
					<td style="WIDTH: 443px"><asp:textbox id="txtTechComments" runat="server" Width="445px" TextMode="MultiLine" Height="100px"
							Font-Names="Arial"></asp:textbox></td>
				</tr>
				<tr>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<td class="Cell9Reg" align="left">QA Status:</td>
					<td style="WIDTH: 443px"><asp:dropdownlist id="ddQAStatus" runat="server">
							<asp:ListItem Value="1">Unresolved</asp:ListItem>
							<asp:ListItem Value="2">Resolved</asp:ListItem>
						</asp:dropdownlist><asp:textbox id="txtQAStatus" runat="server"></asp:textbox></td>
				</tr>
				<tr>
					<td class="Cell9Reg" vAlign="top" align="left">QA Comments:</td>
					<td style="WIDTH: 443px"><asp:textbox id="txtQAComments" runat="server" Width="445px" TextMode="MultiLine" Height="100px"
							Font-Names="Arial"></asp:textbox></td>
				</tr>
				<tr>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<td style="WIDTH: 555px" colSpan="2">&nbsp;</td>
				</tr>
				<tr>
					<td style="WIDTH: 555px" align="center" colSpan="2"><asp:literal id="litUpdate" runat="server" EnableViewState="False"></asp:literal>&nbsp;&nbsp;<INPUT onclick="ReturnToParentPage()" type="button" value="Return"></td>
				</tr>
				<tr>
					<td style="WIDTH: 555px" colSpan="2">&nbsp;</td>
				</tr>
			</table>
			<asp:literal id="litResponseAction" runat="server" EnableViewState="False"></asp:literal><asp:label id="lblCurrentRights" runat="server"></asp:label></form>
		<script language="javascript">				
				new menuitems("userlic1", form1.currentrights.value);
				new tpl("v", "admin");
				new menu (MENU_ITEMS, MENU_TPL);				
		</script>
		<asp:literal id="litMsg" runat="server" EnableViewState="False"></asp:literal>
		<script type="text/javascript">
			function legallength(txtarea, legalmax)  {	
				if (txtarea.getAttribute && txtarea.value.length>legalmax) {
					txtarea.value=txtarea.value.substring(0,legalmax)
				}
			}
			function Update()  {
				form1.hdAction.value = "update"
				form1.submit()
			}
			
			function ReturnToParentPage()  {
				form1.hdAction.value = "return"
				form1.submit()
			}			

			function IsValidChar(sText)  {
				var TextValue;
   				var ValidChars = ".0123456789";
   				var IsValid=true;
   				var DecCount = 0;
   				var Char;   				
   				TextValue = sText.value
   				for (i = 0; i < TextValue.length && IsValid == true; i++)  { 
      				Char = TextValue.charAt(i);
      				if (Char == ".") {
      					DecCount = DecCount + 1
      					if (DecCount == 2) {
      						IsValid = false;
      					}
      				}
      				if (ValidChars.indexOf(Char) == -1) {
         				IsValid = false;
         			}
      			}
   				if (IsValid == false) {
   					sText.value = sText.value.substring(0,i-1)
   				}
   				return;	
   			}					
		</script>
	</body>
</HTML>
