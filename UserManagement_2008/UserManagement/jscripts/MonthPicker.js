var monthpickers = [];

function MonthPicker(obj_target) {
	this.popup    = pm_popup2;
	this.target = obj_target;
	this.id = monthpickers.length;
	monthpickers[this.id] = this;
}

function pm_popup2()  {
	var obj_calwindow = window.open('MonthPicker.aspx?id=' + this.id, 'Calendar', 'width=318,height=262,status=no,resizable=no,top=200,left=200,dependent=yes,alwaysRaised=yes')
	obj_calwindow.opener = window;
	obj_calwindow.focus();
}