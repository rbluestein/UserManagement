var TPLOrientation = "v";
var MENU_TPL = [];

function tpl(TPLOrientation, apppart) {
    if (TPLOrientation == "v" && apppart=="public") {
        MENU_TPL = [{        
	        //  Item sizes //
	        'height': 24,
	        'width': 125,	        	  
	        //  Offset from parent
	        //	Root: offset from left corner of the page
	        //	Other: offset from upper left corner of parent item
        	'block_top': 200,
	        'block_left': 10,        	        	
	        // Sibling offset	
	        'top': 24,
	        'left': 0,     	  	
	        // Time in milliseconds before menu is hidden after cursor has gone out of any items
	        'hide_delay': 100,
	        'expd_delay': 100,
	        'css' : {
		        'outer' : ['m0l0oout', 'm0l0oover'],
		        'inner' : ['m0l0iout', 'm0l0iover']
	        }
        },{
	        'height': 24,
	        'width': 125,        	
	        'block_top': 5,
	        'block_left': 115,        	        	
	        'top': 23,
	        'left': 0,	
	        'css' : {
		        'outer' : ['m0l1oout', 'm0l1oover'],
		        'inner' : ['m0l1iout', 'm0l1iover']
	        }
        },{
	        'block_top': 5,
	        'block_left': 115
        }
        ]
        
      } else if (TPLOrientation == "v" && apppart=="admin") {      
         MENU_TPL = [{        
	        //  Item sizes //
	        'height': 24,
	        'width': 125,	        	  
	        //  Offset from parent
	        //	Root: offset from left corner of the page
	        //	Other: offset from upper left corner of parent item
        	'block_top': 58,
	        'block_left': 10,        	        	
	        // Sibling offset	
	        'top': 24,
	        'left': 0,     	  	
	        // Time in milliseconds before menu is hidden after cursor has gone out of any items
	        'hide_delay': 100,
	        'expd_delay': 100,
	        'css' : {
		        'outer' : ['m0l0oout', 'm0l0oover'],
		        'inner' : ['m0l0iout', 'm0l0iover']
	        }
        },{
	        'height': 24,
	        'width': 125,        	
	        'block_top': 5,
	        'block_left': 115,        	        	
	        'top': 23,
	        'left': 0,	
	        'css' : {
		        'outer' : ['m0l1oout', 'm0l1oover'],
		        'inner' : ['m0l1iout', 'm0l1iover']
	        }
        },{
	        'block_top': 5,
	        'block_left': 115
        }
        ]       
        
      
    } else if (TPLOrientation == "h") {
        MENU_TPL = [{
	            'height': 24,
	            'width': 130,
	            'block_top': 130,
	            'block_left': 300,          	
        	    'top': 0,
        	    'left': 131,            	            	
                'hide_delay': 200,
	            'expd_delay': 200,
	            'css' : {
		            'outer' : ['m0l0oout', 'm0l0oover'],
		            'inner' : ['m0l0iout', 'm0l0iover']
	            }
            },{
	            'height': 24,
	            'width': 170,
        	    'block_top': 25,
        	    'block_left': 0,
	            'top': 23,
	            'left': 0,	
	            'css' : {
		            'outer' : ['m0l1oout', 'm0l1oover'],
		            'inner' : ['m0l1iout', 'm0l1iover']
	            }
            },{
	            'block_top': 5,
	            'block_left': 160
            }
            ]
    }   
}    
