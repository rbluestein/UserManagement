var MENU_ITEMS = [];
var mmitems = [];

function menuitems(apppart, currentrights)  {
	    
	// assign methods and event handlers
	//this.OrigHasPermission	= OrigHasPermission;
	//this.HasPermission			= HasPermission;
	//this.AppendItem				= AppendItem;

	if (apppart == "userlic1") {
		AppendItem({id:'User', parentid:'', arr:['Users', 'Index.aspx?CalledBy=Other'], right:'USV'});
		AppendItem({id:'Client', parentid:'', arr:['Clients', 'ClientWorklist.aspx?CalledBy=Other'], right:'CLV'});		
		AppendItem({id:'Carrier', parentid:'', arr:['Carriers', 'CarrierWorklist.aspx?CalledBy=Other'], right:'CAV'});
		//AppendItem({id:'Company', parentid:'', arr:['Company Contacts', 'CompanyWorklist.aspx?CalledBy=Other'], right:'COV'});
		AppendItem({id:'License', parentid:'', arr:['Licenses', null, null], right:'LIV'});
		AppendItem({id:'BulkLicense', parentid:'License', arr:['License Express', 'BulkLicense.aspx?CalledBy=Other'], right:'LIE'});
		AppendItem({id:'QueryLicense', parentid:'License', arr:['License Query', 'QueryLicense.aspx?CalledBy=Other'], right:'LIV'});		
		AppendItem({id:'QueryCarrier', parentid:'License', arr:['Carrier Query', 'QueryCarrier.aspx?CalledBy=Other'], right:'LIV'});					
	}													
	
	count = 0;
	for (var j=0; j<mmitems.length; j++) {
		if (mmitems[j].parentid == '') {
			MENU_ITEMS[count] = mmitems[j].arr
			count++
		}
	}

	// Load the child menu items //
	count = 0;
	var prnt;
	for (var chld=0; chld<mmitems.length; chld++) {
		//if (mmitems[chld].parentid != ''&& HasPermission(mmitems[chld].right)) {
		if (mmitems[chld].parentid != '') {		
			for (prnt=0; prnt<mmitems.length; prnt++) {        
				if (mmitems[chld].parentid == mmitems[prnt].id) {
					mmitems[prnt].arr[mmitems[prnt].arr.length] = mmitems[chld].arr;
					break;           
				}       
			}
			if (prnt == mmitems.length) {
				alert("Problem creating menu.")
			}  
		}
	}

	function CalendarAddCurMonthYear(PageName) {
		var now = new Date()
		var month = now.getMonth() + 1
		var year = now.getFullYear()
		var value = PageName + '?Month='+ month + '&Year=' + year
		return PageName + '?Month='+ month + '&Year=' + year
	}

	function OrigHasPermission(mitem) {
		// What is the position of the second string in the first string?
		// Is the required right a member of the current rights?
		if (currentrights.toUpperCase().indexOf(mitem.right.toUpperCase() ) == -1) {
			return false
		} else {
			return true
		}
	}
	
	
	function HasPermission(mitem) {
		// What is the position of the second string in the first string?
		// Is the required right a member of the current rights?
	
		var box = mitem.right.toUpperCase().split("|")
		for (var i=0;i<box.length;i++) {	
		
			if (currentrights.toUpperCase().indexOf(box[i]) >= 0) {
				return true;
			}				
		}	
	}	
		
	function AppendItem(mitem) {
		if (HasPermission(mitem)) {
			mmitems[mmitems.length] = mitem;     			
		}
	}
	
}
		