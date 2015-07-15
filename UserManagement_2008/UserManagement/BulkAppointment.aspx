<%@ Page Language="VB" AutoEventWireup="false" CodeFile="BulkAppointment.aspx.vb" Inherits="BulkAppointment" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
		<title id="PageCaption" runat="server"></title>
        <link href="css/BVI.css" rel="stylesheet" type="text/css" />
        <link href="css/menu.css" rel="stylesheet" type="text/css" />
        <link href="css/BVI.css" rel="stylesheet" type="text/css" />        
        <script src="jscripts/calendar2.js" type="text/javascript"></script>
        <script src="jscripts/menu_items.js" type="text/javascript"></script>
        <script src="jscripts/menu.js" type="text/javascript"></script>
        <script src="jscripts/menu_tpl.js" type="text/javascript"></script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
			<input type="hidden" name="hdAction" /><input type="hidden" name="hdSubAction" />
			<asp:literal id="litHiddens" runat="server" enableviewstate="False"></asp:literal>  

			<table class="PrimaryTbl" style="LEFT: 140px; POSITION: absolute; TOP: 14px" cellspacing="0"
				cellpadding="0" width="650" border="0">
				<tr style="DISPLAY: none">
					<td colspan="2">&nbsp;</td>
				</tr>
				<tr>
					<td class="PrimaryTblTitle" colspan="2"><asp:literal id="litHeading" runat="server"></asp:literal></td>
				</tr>
				<tr>
					<td class="CellSeparator" colspan="2"></td>
				</tr>
				<tr>
					<td colspan="2">&nbsp;</td>
				</tr>
				<tr>
					<td colspan="2">
						<table class="PrimaryTblEmbedded" cellspacing="0" cellpadding="0" width="650" border="0">
							<tr>
								<td style="FONT: 16pt Arial, Helvetica, sans-serif" colspan="4">Enroller</td>
							</tr>
							<tr>
								<td width="250"><a href="javascript:Letter('A')">A</a><a href="javascript:Letter('B')">B</a><a href="javascript:Letter('C')">C</a><a href="javascript:Letter('D')">D</a><a href="javascript:Letter('E')">E</a><a href="javascript:Letter('F')">F</a><a href="javascript:Letter('G')">G</a><a href="javascript:Letter('H')">H</a><a href="javascript:Letter('I')">I</a><a href="javascript:Letter('J')">J</a><a href="javascript:Letter('K')">K</a><a href="javascript:Letter('L')">L</a><a href="javascript:Letter('M')">M</a><a href="javascript:Letter('N')">N</a><a href="javascript:Letter('O')">O</a><a href="javascript:Letter('P')">P</a><a href="javascript:Letter('Q')">Q</a><a href="javascript:Letter('R')">R</a><a href="javascript:Letter('S')">S</a><a href="javascript:Letter('T')">T</a><a href="javascript:Letter('U')">U</a><a href="javascript:Letter('V')">V</a><a href="javascript:Letter('W')">W</a><a href="javascript:Letter('X')">X</a><a href="javascript:Letter('Y')">Y</a><a href="javascript:Letter('Z')">Z</a>
									<asp:dropdownlist id="ddEnrollersSource" runat="server"></asp:dropdownlist></td>
								<td align="center" width="50"><asp:button id="btnEnrollerAdd" runat="server" text="  >>  " width="50px"></asp:button><asp:button id="btnEnrollerRemove" runat="server" text="  <<  " width="50px"></asp:button></td>
								<td align="right" width="250">&nbsp;<br />
									<asp:textbox id="txtEnrollerTarget" runat="server"></asp:textbox></td>
								<td width="100"></td>
							</tr>
							<tr>
								<td colspan="2">&nbsp;</td>
							</tr>
							<tr>
								<td style="FONT: 16pt Arial, Helvetica, sans-serif" colspan="4">Carrier</td>
							</tr>
							<tr>
								<td width="250"><asp:listbox id="lstCarriersSource" runat="server" width="160px" rows="7"></asp:listbox></td>
								<td align="center" width="50"><asp:button id="btnCarrierAdd" runat="server" text="  >>  " width="50px"></asp:button><asp:button id="btnCarrierRemove" runat="server" text="  <<  " width="50px"></asp:button></td>
								<td align="right" width="250"><asp:listbox id="lstCarriersTarget" runat="server" width="160px" rows="7"></asp:listbox></td>
								<td width="100"></td>
							</tr>
							<tr>
								<td style="FONT: 16pt Arial, Helvetica, sans-serif" colspan="4">State</td>
							</tr>
							<tr>
								<td width="250"><asp:listbox id="lstStatesSource" runat="server" width="160px" rows="7"></asp:listbox></td>
								<td align="center" width="50"><asp:button id="btnStateAdd" runat="server" text="  >>  " width="50px"></asp:button><asp:button id="btnStateRemove" runat="server" text="  <<  " width="50px"></asp:button></td>
								<td align="right" width="250"><asp:listbox id="lstStatesTarget" runat="server" width="160px" rows="7"></asp:listbox></td>
								<td width="100"></td>
							</tr>
							<tr>
								<td colspan="4">&nbsp;</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td class="Cell9Reg" style="WIDTH: 120px" align="left">Current Input:</td>
					<td><asp:textbox id="txtStatus" runat="server" maxlength="50"></asp:textbox></td>
				</tr>
				<tr>
					<td class="Cell9Reg" style="WIDTH: 120px" align="left">Appointment Number:</td>
					<td><asp:textbox id="txtAppointmentNumber" runat="server" maxlength="50"></asp:textbox></td>
				</tr>
				<tr>
					<td class="Cell9Reg" style="WIDTH: 120px" align="left" width="120">
						Pending&nbsp;Date:</td>
					<td class="Cell9Reg"><asp:textbox id="txtApplicationDate" runat="server" width="112px" maxlength="50"></asp:textbox>&nbsp;&nbsp;
						<asp:label id="lblApplicationDateLink" runat="server">
							<a href="javascript:GetDate('ApplicationDate')">Get Date</a></asp:label></td>
				</tr>
				<tr>
					<td class="Cell9Reg" style="WIDTH: 120px" align="left" width="120">Effective Date:</td>
					<td class="Cell9Reg"><asp:textbox id="txtEffectiveDate" runat="server" width="112px" maxlength="50"></asp:textbox>&nbsp;&nbsp;
						<asp:label id="lblEffectiveDateLink" runat="server">
							<a href="javascript:GetDate('EffectiveDate')">Get Date</a></asp:label></td>
				</tr>
				<asp:placeholder id="plExpirationDate" runat="server" visible="False">
					<tr>
						<td class="Cell9Reg" style="WIDTH: 120px" align="left">ExpirationDate:</td>
						<td class="Cell9Reg">
							<asp:textbox id="txtExpirationDate" runat="server" width="112px" maxlength="50"></asp:textbox>&nbsp;&nbsp;
							<asp:label id="lblExpirationDateLink" runat="server">
								<a href="javascript:GetDate('ExpirationDate')">Get Date</a></asp:label></td>
					</tr>
				</asp:placeholder>
				<tr>
					<td class="Cell9Reg" style="WIDTH: 120px" align="left" width="120">
						Status:</td>
					<td>
						<asp:dropdownlist id="ddStatusCode" runat="server"></asp:dropdownlist></td>
				</tr>
				<tr>
					<td colspan="2">&nbsp;</td>
				</tr>
				<tr>
					<td align="center" colspan="2"><asp:literal id="litUpdate" runat="server" enableviewstate="False"></asp:literal>&nbsp;&nbsp;</td>
				</tr>
				<tr>
					<td colspan="2">&nbsp;</td>
				</tr>
			</table>			
			<asp:literal id="litEnviro" runat="server" enableviewstate="False"></asp:literal>
			<asp:literal id="litMsg" runat="server" enableviewstate="False"></asp:literal>
			<asp:literal id="litResponseAction" runat="server" enableviewstate="False"></asp:literal><asp:label id="lblCurrentRights" runat="server"></asp:label>					  
    </div>
    </form>
</body>
</html>
