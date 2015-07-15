<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MonthPicker.aspx.vb" Inherits="MonthPicker" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<title>Select Date</title>
        <link href="css/JSControls.css" rel="stylesheet" type="text/css" />
	</head>
	<body onunload="fnClose()">
		<form id="form2" method="post" runat="server">
			<input type='hidden' id='hdCurMonth' /><input type='hidden' id='hdCurYear' /><input type='hidden' id='hdInitial' value='1' />
			<input type='hidden' id='hdFromMonth' /><input type='hidden' id='hdFromYear' /><input type='hidden' id='hdToMonth' /><input type='hidden' id='hdToYear' />
			<table cellpadding="0" style="LEFT: 10px; POSITION: absolute; TOP: 0px">
				<tr><td style="font: bold 13px Verdana">From Date</td></tr>
			</table>
			<table class="MPtable" cellpadding="0" style="LEFT: 10px; POSITION: absolute; TOP: 20px">
				<tr>
					<td class="MPtd"><a href="javascript:fnSelectYear(0, -1)">Prev Year</a></td>
					<td class="MPtd"><a href="javascript:fnSelectYear(0, 1)">Next Year</a></td>
				</tr>
			</table>
			<table class="MPtable" cellpadding="0" style="LEFT: 10px; POSITION: absolute; TOP: 40px">
				<tr>
					<td class="MPtd"><a href="javascript:fnSelectMonth('0','0')">Jan</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('0','1')">Feb</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('0','2')">Mar</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('0','3')">Apr</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('0','4')">May</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('0','5')">Jun</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('0','6')">Jul</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('0','7')">Aug</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('0','8')">Sep</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('0','9')">Oct</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('0','10')">Nov</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('0','11')">Dec</a></td>
				</tr>
			</table>
			<table cellpadding="2" cellspacing="0" style="LEFT: 10px; POSITION: absolute; TOP: 60px">
				<tr>
					<td><input type="text" id="txtFromMonthYear" name="txtFromMonthYear" readonly="readonly" /></td>
				</tr>
			</table>
			<table class="MPtable" cellpadding="0" style="LEFT: 10px; POSITION: absolute; TOP: 87px">
				<tr>
					<td class="MPtd"><a href="javascript:fnClear(0)">Clear</a></td>
				</tr>
			</table>
			
			
			
			
			<table cellpadding="0" style="LEFT: 10px; POSITION: absolute; TOP: 120px">
				<tr><td style="font: bold 13px Verdana">To Date</td></tr>
			</table>			
			<table class="MPtable" cellpadding="0" style="LEFT: 10px; POSITION: absolute; TOP: 140px">
				<tr>
					<td class="MPtd"><a href="javascript:fnSelectYear(1, -1)">Prev Year</a></td>
					<td class="MPtd"><a href="javascript:fnSelectYear(1, 1)">Next Year</a></td>
				</tr>
			</table>
			<table class="MPtable" cellpadding="0" style="LEFT: 10px; POSITION: absolute; TOP: 160px">
				<tr>
					<td class="MPtd"><a href="javascript:fnSelectMonth('1','0')">Jan</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('1','1')">Feb</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('1','2')">Mar</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('1','3')">Apr</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('1','4')">May</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('1','5')">Jun</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('1','6')">Jul</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('1','7')">Aug</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('1','8')">Sep</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('1','9')">Oct</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('1','10')">Nov</a></td>
					<td class="MPtd"><a href="javascript:fnSelectMonth('1','11')">Dec</a></td>
				</tr>
			</table>
			<table cellpadding="0" style="LEFT: 10px; POSITION: absolute; TOP: 180px">
				<tr>
					<td><input type="text" id="txtToMonthYear" name="txtToMonthYear" readonly="readonly" /></td>
				</tr>
			</table>
			<table class="MPtable" cellpadding="0" style="LEFT: 10px; POSITION: absolute; TOP: 207px">
				<tr>
					<td class="MPtd"><a href="javascript:fnClear(1)">Clear</a></td>
				</tr>
			</table>
			<table class="MPtable" cellpadding="0" style="LEFT: 10px; POSITION: absolute; TOP: 237px">
				<tr>
					<td class="MPtd"><a href="javascript:fnClose()">Done</a></td>
				</tr>
			</table>
			<script type="text/javascript">
			
				// On Document Load
			
					var re_id = new RegExp('id=(\\d+)');
					var num_id = (re_id.exec(String(window.location))  ? new Number(RegExp.$1) : 0);
					var obj_caller = (window.opener ? window.opener.monthpickers[num_id] : null);
					
					var ARR_MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];		
					var ShowDate =  String(ARR_MONTHS[form1.hdFromMonth.value]) + '&nbsp;' + String(form1.hdFromYear.value)	
							
					if (form1.hdInitial.value == 1)  {	
						var ArrValues = obj_caller.target.value.split("|")						
						form1.hdFromMonth.value = ArrValues[0]			
						form1.hdFromYear.value = ArrValues[1]	
						form1.hdToMonth.value = ArrValues[2]	
						form1.hdToYear.value = ArrValues[3]												
						form1.hdInitial.value = "0"
					}						
						
					fnShowDate(0)
					fnShowDate(1)					
					
				// Methods				
					
					function fnClose()  {
						fnSetSourceControl()
						window.close();
					}			

						
		
												
					function fnSelectMonth(vSelect, vValue)  {
						fnAdjustYear(vSelect)			
						if (vSelect == 0)  {				
							form1.hdFromMonth.value = vValue
						} else {
							form1.hdToMonth.value = vValue				
						}			
						fnShowDate(vSelect)
						//fnSetSourceControl()
					}
								
					function fnSelectYear(vSelect, vValue)  {
						fnAdjustMonth(vSelect)				
						if (vSelect == 0) {	
							if (form1.hdFromYear.value == '')  {
								fnAdjustYear(0)			
							}  else {
								form1.hdFromYear.value = parseInt(form1.hdFromYear.value) + vValue										
							}				
						}	
						if (vSelect == 1) {	
							if (form1.hdToYear.value == '')  {
								fnAdjustYear(1)			
							}  else {
								form1.hdToYear.value = parseInt(form1.hdToYear.value) + vValue										
							}			
						}						
						fnShowDate(vSelect)					
						//fnSetSourceControl()			
					}		
					
					function fnShowDate(vSelect) 
					{
						if (vSelect == 0) {
							if (form1.hdFromMonth.value == "" && form1.hdFromYear.value == "") {
								form1.txtFromMonthYear.value = ""
							} else {
								form1.txtFromMonthYear.value  = ARR_MONTHS[form1.hdFromMonth.value] + ' ' + form1.hdFromYear.value	
							}
						}
						
						if (vSelect == 1) {
							if (form1.hdToMonth.value == "" && form1.hdToYear.value == "") {
								form1.txtToMonthYear.value = ""
							} else {
								form1.txtToMonthYear.value  = ARR_MONTHS[form1.hdToMonth.value] + ' ' + form1.hdToYear.value	
							}
						}					
					
					}		
					
					function fnClear(vSelect)  {
						if (vSelect == 0)  {		
							form1.hdFromMonth.value = ''
							form1.hdFromYear.value = ''				
						} else {		
							form1.hdToMonth.value = ''
							form1.hdToYear.value = ''	
						}
						fnShowDate(vSelect)					
						//fnSetSourceControl()		
					}	
					
					function fnAdjustMonth(vSelect) 	{
						if (vSelect == 0  && form1.hdFromMonth.value == "")	{
							form1.hdFromMonth.value = 0
						} 
						if (vSelect == 1 && form1.hdToMonth.value == '')	{
							form1.hdToMonth.value = 0
						}												
					}
					
					function fnAdjustYear(vSelect)  {
						if (vSelect == 0 && form1.hdFromYear.value == '') {
								var d = new Date()
								form1.hdFromYear.value = d.getFullYear()
						}	
						if  (vSelect == 1 && form1.hdToYear.value == '') 	{
								var d = new Date()
								form1.hdToYear.value = d.getFullYear()				
						}			
					}			
																								
					function fnSetSourceControl() {
							obj_caller.target.value = form1.hdFromMonth.value + '|' + form1.hdFromYear.value + '|' + form1.hdToMonth.value + '|' + form1.hdToYear.value
					}	
				
			</script>
		</form>
	</body>
</html>
