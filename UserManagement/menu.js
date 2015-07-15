/// Title: tigra menu
// Description: See the demo at url
// URL: http://www.softcomplex.com/products/tigra_menu/
// Version: 2.0 (commented source)
// Date: 04-05-2003 (mm-dd-yyyy)
// Tech. Support: http://www.softcomplex.com/forum/forumdisplay.php?fid=40
// Notes: This script is free. Visit official site for further details.

// --------------------------------------------------------------------------------
// global collection containing all menus on current page
var a_MenuObjects = [];

// --------------------------------------------------------------------------------
// menu class
function menu (a_menuitems, a_tpl) {

    this._bobname = "MenuObject"

	// browser check
	if (!document.body || !document.body.style)
		return;

	// store items structure
	this.a_menuitems = a_menuitems;

	// store template structure
	this.a_tpl = a_tpl;

	// get menu id
	this.n_id = a_MenuObjects.length;

	// declare collections
	this.a_index = [];
	this.a_children = [];

	// assign methods and event handlers
	this.expand      = menu_expand;
	this.collapse    = menu_collapse;

	this.onclick     = menu_onclick;
	this.onmouseout  = menu_onmouseout;
	this.onmouseover = menu_onmouseover;
	this.onmousedown = menu_onmousedown;

	// default level scope description structure 
	this.a_tpl_def = {
		'block_top'  : 16,
		'block_left' : 16,
		'top'        : 20,
		'left'       : 4,
		'width'      : 120,
		'height'     : 22,
		'hide_delay' : 0,
		'expd_delay' : 0,
		'css'        : {
			'inner' : '',
			'outer' : ''
		}
	};
	
	// assign methods and properties required to emulate parent item
	this.getdefaulttplprop = function (s_key) {
		return this.a_tpl_def[s_key];
	};

	this.o_root = this;
	this.n_depth = -1;
	this.n_x = 0;
	this.n_y = 0;

	// 	init items recursively (top-level)
	// this executes one time only
	for (n_order = 0; n_order < a_menuitems.length; n_order++)
		new menu_item(this, n_order);

	// register self in global collection
	a_MenuObjects[this.n_id] = this;

	// make root level visible by setting the style property of the
	// root level menu items to visible
	// this executes one time only
	for (var n_order = 0; n_order < this.a_children.length; n_order++)
		this.a_children[n_order].e_oelement.style.visibility = 'visible';
}

// --------------------------------------------------------------------------------
function menu_collapse (n_id) {
	// cancel item open delay
	clearTimeout(this.o_showtimer);

	// by default collapse to root level
	var n_tolevel = (n_id ? this.a_index[n_id].n_depth : 0);
	
	// hide all items over the level specified
	for (n_id = 0; n_id < this.a_index.length; n_id++) {
		var o_curritem = this.a_index[n_id];
		if (o_curritem.n_depth > n_tolevel && o_curritem.b_visible) {
			o_curritem.e_oelement.style.visibility = 'hidden';
			o_curritem.b_visible = false;
		}
	}

	// reset current item if mouse has gone out of items
	if (!n_id)
		this.o_current = null;
}

// --------------------------------------------------------------------------------
function menu_expand (n_id) {

	// expand only when mouse is over some menu item
	if (this.o_hidetimer)
		return;

	// lookup current item
	var o_item = this.a_index[n_id];

	// close previously opened items
	if (this.o_current && this.o_current.n_depth >= o_item.n_depth)
		this.collapse(o_item.n_id);
	this.o_current = o_item;

	// exit if there are no children to open
	if (!o_item.a_children)
		return;

	// show direct child items
	for (var n_order = 0; n_order < o_item.a_children.length; n_order++) {
		var o_curritem = o_item.a_children[n_order];
		o_curritem.e_oelement.style.visibility = 'visible';
		o_curritem.b_visible = true;
	}
}

// --------------------------------------------------------------------------------
//
// --------------------------------------------------------------------------------
function menu_onclick (n_id) {
	// don't go anywhere if item has no link defined
	return Boolean(this.a_index[n_id].a_menuitems[1]);
}

// --------------------------------------------------------------------------------
function menu_onmouseout (n_id) {

	// lookup new item's object	
	var o_item = this.a_index[n_id];

	// apply rollout
	o_item.e_oelement.className = o_item.getstyle(0, 0);
	o_item.e_ielement.className = o_item.getstyle(1, 0);
	
	// update status line	
	o_item.upstatus(7);

	// run mouseover timer
	this.o_hidetimer = setTimeout('a_MenuObjects['+ this.n_id +'].collapse();',
		o_item.getdefaulttplprop('hide_delay'));
}

// --------------------------------------------------------------------------------
function menu_onmouseover (n_id) {

	// cancel mouseoute menu close and item open delay
	clearTimeout(this.o_hidetimer);
	this.o_hidetimer = null;
	clearTimeout(this.o_showtimer);

	// lookup new item's object	
	var o_item = this.a_index[n_id];

	// update status line	
	o_item.upstatus();

	// apply rollover
	o_item.e_oelement.className = o_item.getstyle(0, 1);
	o_item.e_ielement.className = o_item.getstyle(1, 1);
	
	// if onclick open is set then no more actions required
	if (o_item.getdefaulttplprop('expd_delay') < 0)
		return;

	// run expand timer
	this.o_showtimer = setTimeout('a_MenuObjects['+ this.n_id +'].expand(' + n_id + ');',
		o_item.getdefaulttplprop('expd_delay'));

}

// --------------------------------------------------------------------------------
// called when mouse button is pressed on menu item
// --------------------------------------------------------------------------------
function menu_onmousedown (n_id) {
	
	// lookup new item's object	
	var o_item = this.a_index[n_id];

	// apply mouse down style
	o_item.e_oelement.className = o_item.getstyle(0, 2);
	o_item.e_ielement.className = o_item.getstyle(1, 2);

	this.expand(n_id);
//	this.items[id].switch_style('onmousedown');
}


// --------------------------------------------------------------
// menu item Class
function menu_item (o_parent, n_order) {

    this._bobname = "MenuItemObject"

	// store parameters passed to the constructor
	this.n_depth  = o_parent.n_depth + 1;
	this.a_menuitems = o_parent.a_menuitems[n_order + (this.n_depth ? 3 : 0)];    //menuitems: (0) Events, (1) Announcements, (2) Users, (3) Change Password. First time throught, selects Events

	// return if required parameters are missing
	if (!this.a_menuitems) return;

	// store info from parent item
	this.o_root    = o_parent.o_root;
	this.o_parent  = o_parent;
	this.n_order   = n_order;

	// register in global and parent's collections
	this.n_id = this.o_root.a_index.length;
	this.o_root.a_index[this.n_id] = this;
	o_parent.a_children[n_order] = this;

	// calculate item's coordinates
	var o_root = this.o_root,
	a_tpl  = this.o_root.a_tpl;

	// assign methods
	this.getdefaulttplprop  = mitem_gettplprop;
	this.getstyle = mitem_getstyle;
	this.upstatus = mitem_upstatus;
	
	// if this item is a top-level menu, n_x is tpldefault('block_left')
	// otherwise, n_x is its parent left plus tpldefault('block_left')
    this.n_x = n_order
	    ? o_parent.a_children[n_order - 1].n_x + this.getdefaulttplprop('left')
	    : o_parent.n_x + this.getdefaulttplprop('block_left');

    this.n_y = n_order
	    ? o_parent.a_children[n_order - 1].n_y + this.getdefaulttplprop('top')
	    : o_parent.n_y + this.getdefaulttplprop('block_top');

//var dog = 	'<a id="e' + o_root.n_id + '_'
//			+ this.n_id +'o" class="' + this.getstyle(0, 0) + '" href="' + this.a_menuitems[1] + '"'
//			+ (this.a_menuitems[2] && this.a_menuitems[2]['tw'] ? ' target="'
//			+ this.a_menuitems[2]['tw'] + '"' : '')
//			+ (this.a_menuitems[2] && this.a_menuitems[2]['tt'] ? ' title="'
//			+ this.a_menuitems[2]['tt'] + '"' : '') + ' style="position: absolute; top: '
//			+ this.n_y + 'px; left: ' + this.n_x + 'px; width: '
//			+ this.getdefaulttplprop('width') + 'px; height: '
//			+ this.getdefaulttplprop('height') + 'px; visibility: hidden;'
//			+' z-index: ' + this.n_depth + ';" '
//			+ 'onclick="return a_MenuObjects[' + o_root.n_id + '].onclick('
//			+ this.n_id + ');" onmouseout="a_MenuObjects[' + o_root.n_id + '].onmouseout('
//			+ this.n_id + ');" onmouseover="a_MenuObjects[' + o_root.n_id + '].onmouseover('
//			+ this.n_id + ');" onmousedown="a_MenuObjects[' + o_root.n_id + '].onmousedown('
//			+ this.n_id + ');"><div  id="e' + o_root.n_id + '_'
//			+ this.n_id +'i" class="' + this.getstyle(1, 0) + '">'
//			+ this.a_menuitems[0] + "</div></a>\n"

//        <a 
//        id=\"e0_0o\" 
//        class=\"m0l0oout\" 
//        href=\"null\" 
//        style=\"position: absolute; top: 130px; left: 300px; width: 130px; height: 24px; visibility: hidden; z-index: 0;\" 
//        onclick=\"return a_MenuObjects[0].onclick(0);\" 
//        onmouseout=\"a_MenuObjects[0].onmouseout(0);\" 
//        onmouseover=\"a_MenuObjects[0].onmouseover(0);\" 
//        onmousedown=\"a_MenuObjects[0].onmousedown(0);\"
//        >
//        <div  
//        id=\"e0_0i\" 
//        class=\"m0l0iout\">Menu Compatibility
//        </div
//        ></a>			



	// generate item's HMTL
	document.write (
		'<a id="e' + o_root.n_id + '_'
			+ this.n_id +'o" class="' + this.getstyle(0, 0) + '" href="' + this.a_menuitems[1] + '"'
			+ (this.a_menuitems[2] && this.a_menuitems[2]['tw'] ? ' target="'
			+ this.a_menuitems[2]['tw'] + '"' : '')
			+ (this.a_menuitems[2] && this.a_menuitems[2]['tt'] ? ' title="'
			+ this.a_menuitems[2]['tt'] + '"' : '') + ' style="position: absolute; top: '
			+ this.n_y + 'px; left: ' + this.n_x + 'px; width: '
			+ this.getdefaulttplprop('width') + 'px; height: '
			+ this.getdefaulttplprop('height') + 'px; visibility: hidden;'
			+' z-index: ' + this.n_depth + ';" '
			+ 'onclick="return a_MenuObjects[' + o_root.n_id + '].onclick('
			+ this.n_id + ');" onmouseout="a_MenuObjects[' + o_root.n_id + '].onmouseout('
			+ this.n_id + ');" onmouseover="a_MenuObjects[' + o_root.n_id + '].onmouseover('
			+ this.n_id + ');" onmousedown="a_MenuObjects[' + o_root.n_id + '].onmousedown('
			+ this.n_id + ');"><div  id="e' + o_root.n_id + '_'
			+ this.n_id +'i" class="' + this.getstyle(1, 0) + '">'
			+ this.a_menuitems[0] + "</div></a>\n"
		);
	this.e_ielement = document.getElementById('e' + o_root.n_id + '_' + this.n_id + 'i');
	this.e_oelement = document.getElementById('e' + o_root.n_id + '_' + this.n_id + 'o');

	this.b_visible = !this.n_depth;

	// no more initialization if leaf
	if (this.a_menuitems.length < 4)
		return;

	// node specific methods and properties
	this.a_children = [];

	// init downline recursively (below top-level)
	for (var n_order = 0; n_order < this.a_menuitems.length - 3; n_order++)
		new menu_item(this, n_order);
}

// --------------------------------------------------------------------------------
// reads property from template file, inherits from parent level if not found
// ------------------------------------------------------------------------------------------
function mitem_gettplprop (s_key) {

	// check if value is defined for current level
	var s_value = null,
		a_level = this.o_root.a_tpl[this.n_depth];

	// return value if explicitly defined
	if (a_level)
		s_value = a_level[s_key];

	// request recursively from parent levels if not defined
	return (s_value == null ? this.o_parent.getdefaulttplprop(s_key) : s_value);
}
// --------------------------------------------------------------------------------
// reads property from template file, inherits from parent level if not found
// ------------------------------------------------------------------------------------------
function mitem_getstyle (n_pos, n_state) {

	var a_css = this.getdefaulttplprop('css');
	var a_oclass = a_css[n_pos ? 'inner' : 'outer'];

	// same class for all states	
	if (typeof(a_oclass) == 'string')
		return a_oclass;

	// inherit class from previous state if not explicitly defined
	for (var n_currst = n_state; n_currst >= 0; n_currst--)
		if (a_oclass[n_currst])
			return a_oclass[n_currst];
}

// ------------------------------------------------------------------------------------------
// updates status bar message of the browser
// ------------------------------------------------------------------------------------------
function mitem_upstatus (b_clear) {
	window.setTimeout("window.status=unescape('" + (b_clear
		? ''
		: (this.a_menuitems[2] && this.a_menuitems[2]['sb']
			? escape(this.a_menuitems[2]['sb'])
			: escape(this.a_menuitems[0]) + (this.a_menuitems[1]
				? ' ('+ escape(this.a_menuitems[1]) + ')'
				: ''))) + "')", 10);
}

// --------------------------------------------------------------------------------
// that's all folks