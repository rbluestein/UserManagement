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
					//var ShowDate =  String(ARR_MONTHS[form1.hdFromMonth.value]) + '&nbsp;' + String(form1.hdFromYear.value)
					var ShowDate = String(ARR_MONTHS[document.getElementById("hdFromMonth.value")]) + '&nbsp;' + String(document.getElementById("hdFromYear").value)	
							
					if (docment.getElementById("hdInitial").value == 1)  {	
						var ArrValues = obj_caller.target.value.split("|")
						document.getElementById("hdFromMonth").value = ArrValues[0]
						document.getElementById("hdFromYear").value = ArrValues[1]
						document.getElementById("hdToMonth").value = ArrValues[2]
						document.getElementById("hdToYear").value = ArrValues[3]
						document.getElementById("hdInitial").value = "0"
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
							document.getElementById("hdFromMonth").value = vValue
						} else {
							document.getElementById("hdToMonth").value = vValue	
						}			
						fnShowDate(vSelect)
						//fnSetSourceControl()
					}
								
					function fnSelectYear(vSelect, vValue)  {
						fnAdjustMonth(vSelect)				
						if (vSelect == 0) {
						    if (document.getElementById("hdFromYear").value == '') {
								fnAdjustYear(0)			
							}  else {
							    //form1.hdFromYear.value = parseInt(form1.hdFromYear.value) + vValue
							    document.getElementById("hdFromYear").value = parseInt(document.getElementById("hdFromYear").value) + vValue								
							}				
						}	
						if (vSelect == 1) {	
							if (document.getElementById("hdToYear").value == '')  {
								fnAdjustYear(1)			
							}  else {
							    //form1.hdToYear.value = parseInt(form1.hdToYear.value) + vValue
							    document.getElementById("hdToYear").value = parseInt(document.getElementById("hdToYear").value) + vValue																	
							}			
						}						
						fnShowDate(vSelect)					
						//fnSetSourceControl()			
					}		
					
					function fnShowDate(vSelect) 
					{
						if (vSelect == 0) {
						    //if (form1.hdFromMonth.value == "" && form1.hdFromYear.value == "") {
							if (document.getElementById("hdFromMonth").value == "" && document.getElementById("hdFromYear").value == "") {
								document.getElementById("txtFromMonthYear").value = ""
							} else {
							    //form1.txtFromMonthYear.value  = ARR_MONTHS[form1.hdFromMonth.value] + ' ' + form1.hdFromYear.value
							    document.getElementById("txtFromMonthYear").value = ARR_MONTHS[document.getElementById("hdFromMonth").value] + ' ' + document.getElementById("hdFromYear").value	
							}
						}
						
						if (vSelect == 1) {
							if (document.getElementById("hdToMonth").value == "" && document.getElementById("hdToYear").value == "") {
								document.getElementById("txtToMonthYear").value = ""
							} else {
								document.getElementById("txtToMonthYear").value  = ARR_MONTHS[document.getElementById("hdToMonth").value] + ' ' + document.getElementById("hdToYear").value	
							}
						}					
					
					}		
					
					function fnClear(vSelect)  {
						if (vSelect == 0)  {		
							document.getElementById("hdFromMonth").value = '' 
							document.getElementById("hdFromYear").value = ''	
						} else {		
							document.getElementById("hdToMonth").value = ''
							document.getElementById("hdToYear").value = ''
						}
						fnShowDate(vSelect)					
						//fnSetSourceControl()		
					}	
					
					function fnAdjustMonth(vSelect) 	{
						if (vSelect == 0  && document.getElementById("hdFromMonth").value == "")	{
						    document.getElementById("hdFromMonth").value = 0
						} 
						if (vSelect == 1 && document.getElementById("hdToMonth").value == '')	{
							document.getElementById("hdToMonth").value = 0
						}												
					}
					
					function fnAdjustYear(vSelect)  {
						if (vSelect == 0 && document.getElementById("hdFromYear").value == '') {
								var d = new Date()
								document.getElementById("hdFromYear").value = d.getFullYear()
						}	
						if  (vSelect == 1 && document.getElementById("hdToYear").value == '') 	{
								var d = new Date()
								document.getElementById("hdToYear").value = d.getFullYear()				
						}			
					}			
																								
					function fnSetSourceControl() {
							obj_caller.target.value = document.getElementById("hdFromMonth").value + '|' + document.getElementById("hdFromYear").value + '|' + document.getElementById("hdToMonth").value + '|' + document.getElementById("hdToYear").value
					}	
				
			</script>
		</form>
	</body>
</html>
