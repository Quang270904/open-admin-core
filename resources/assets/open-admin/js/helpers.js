/*--------------------------------------------------*/
/* visual */
/*--------------------------------------------------*/

var show = function (list, display) {
	if (typeof (display) === 'undefined') {
		display = "block";
	}
	if (!isNodeList(list)) {
		var list = [list];
	}
	list.forEach(elm => {
		showElm(elm, display);
	});
};
function showElm(elm, display) {
	if (elm.tagName == "TR") {
		elm.style.display = "table-row";
	} else {
		elm.style.display = display;
	}
}

var hide = function (list) {
	if (!isNodeList(list)) {
		var list = [list];
		isNodeList(list)
	}
	list.forEach(elm => {
		elm.style.display = 'none';
	});
};

var toggle = function (list) {
	if (!isNodeList(list)) {
		var list = [list];
	}
	list.forEach(elm => {
		let calculatedStyle = window.getComputedStyle(elm).display;
		if (calculatedStyle === 'block' || calculatedStyle === 'flex' || calculatedStyle === 'table-row') {
			elm.style.display = 'none';
			return;
		}
		showElm(elm);
	});
};

/*--------------------------------------------------*/
/* lang function */
/*--------------------------------------------------*/

var __ = function (trans_string) {
	return admin_lang_arr[trans_string];
}

var trans = __;

/*--------------------------------------------------*/
/* array / object helpers */
/*--------------------------------------------------*/

var merge_default = function (defaults, object, ...rest) {
	return Object.assign({}, defaults, object, ...rest);
}

var arr_remove = function (arr, elem) {
	var indexElement = arr.findIndex(el => el == elem);
	if (indexElement != -1)
		arr.splice(indexElement, 1);
	return arr;
};

var arr_includes = function (arr, elem) {
	var indexElement = arr.findIndex(el => el == elem);
	return (indexElement != -1)
};

/*--------------------------------------------------*/
/* event Handlers  */
/*--------------------------------------------------*/

function delegate(selector, handler) {

	return function (event) {
		var targ = event.target;
		do {
			if (targ.matches(selector)) {
				handler.call(targ, event);
			}
		} while ((targ = targ.parentNode) && targ != event.currentTarget);
	}
}

/*--------------------------------------------------*/
/* html elements */
/*--------------------------------------------------*/

function getOuterHeigt(el) {
	// Get the DOM Node if you pass in a string
	el = (typeof el === 'string') ? document.querySelector(el) : el;

	var styles = window.getComputedStyle(el);
	var margin = parseFloat(styles['marginTop']) +
		parseFloat(styles['marginBottom']);

	return Math.ceil(el.offsetHeight + margin);
}

function isNodeList(nodes) {
	var stringRepr = Object.prototype.toString.call(nodes);

	return typeof nodes === 'object' &&
		/^\[object (HTMLCollection|NodeList|Object)\]$/.test(stringRepr) &&
		(typeof nodes.length === 'number') &&
		(nodes.length === 0 || (typeof nodes[0] === "object" && nodes[0].nodeType > 0));
}

/**
 * @param {String} HTML representing a single element
 * @return {Element}
 */
function htmlToElement(html) {
	var template = document.createElement('template');
	html = html.trim(); // Never return a text node of whitespace as the result
	template.innerHTML = html;
	return template.content.firstChild;
}

/**
 * @param {String} HTML representing any number of sibling elements
 * @return {NodeList}
 */
function htmlToElements(html) {
	var template = document.createElement('template');
	template.innerHTML = html;
	return template.content.childNodes;
}


function bindSubmitButtonWithLoading() {
	$(document).off('click.submit', 'form button[type="submit"]')
		.on('click.submit', 'form button[type="submit"]', function (e) {
			const btn = $(this);
			const form = btn.closest('form');


			if (!form[0].checkValidity()) {
				e.preventDefault();
				const firstInvalid = form.find(':invalid')[0];
				if (firstInvalid) {
					firstInvalid.scrollIntoView({ behavior: 'instant', block: 'center' });
					firstInvalid.reportValidity();
				}
				return;
			}
			;
		});
}

function clickEvent() {
	$(document).pjax('a:not(a[target="_blank"]):not([data-nopjax]):not([href*="export"])', {
		container: '#pjax-container',
		timeout: 2000
	});

	$(document).on('pjax:click', function () {
		document.querySelectorAll('[data-bs-toggle="tooltip"]').forEach(function (el) {
			const instance = bootstrap.Tooltip.getInstance(el);
			if (instance) {
				instance.dispose();
			}
		});
	});
}

// function handleSidebar() {
// 	const path = location.pathname;

// 	//Save the state of currently opened submenus
// 	const openMenus = [];

// 	document.querySelectorAll('.submenu').forEach(submenu => {
// 		if (getComputedStyle(submenu).display !== 'none') {
// 			const parentItem = submenu.closest('.menu-item');
// 			if (parentItem) openMenus.push(parentItem.dataset.uri || submenu.id);
// 		}
// 	});

// 	//Reset all menu items (remove active classes and hide submenus)
// 	document.querySelectorAll('.menu-item').forEach(item => {
// 		item.classList.remove('active');
// 		item.querySelector('.submenu')?.style.setProperty('display', 'none');
// 		item.querySelector('.has-subs')?.classList.remove('active');
// 	});

// 	//Activate menu items based on current URL path
// 	document.querySelectorAll('.menu-item[data-uri]').forEach(item => {
// 		const uri = item.dataset.uri;
// 		if (uri && (path === uri || path.startsWith(uri + '/'))) {
// 			item.classList.add('active');
// 			item.querySelector('.submenu')?.style.setProperty('display', 'block');
// 			item.querySelector('.has-subs')?.classList.add('active');
// 			let parent = item.parentElement;
// 			while (parent && parent.closest('.menu-item')) {
// 				const menuItem = parent.closest('.menu-item');
// 				menuItem.classList.add('active');
// 				menuItem.querySelector('.submenu')?.style.setProperty('display', 'block');
// 				menuItem.querySelector('.has-subs')?.classList.add('active');
// 				parent = menuItem.parentElement;
// 			}
// 		}
// 	});

// 	//Restore previously opened submenus
// 	openMenus.forEach(id => {
// 		let item = document.querySelector(`.menu-item[data-uri="${id}"]`)
// 			|| document.querySelector(`#${id}`)?.closest('.menu-item');
// 		console.log(item);

// 		if (item) {
// 			item.classList.add('active');
// 			item.querySelector('.submenu')?.style.setProperty('display', 'block');
// 			item.querySelector('.has-subs')?.classList.add('active');
// 		}
// 	});

// 	//Reset toggle links to avoid duplicate event listeners
// 	document.querySelectorAll('.has-subs').forEach(toggleLink => {
// 		const newToggle = toggleLink.cloneNode(true);
// 		toggleLink.parentNode.replaceChild(newToggle, toggleLink);
// 	});

// 	// Add click event for submenu toggles
// 	document.querySelectorAll('.has-subs').forEach(toggleLink => {
// 		toggleLink.addEventListener('click', function (e) {
// 			e.preventDefault();
// 			e.stopPropagation();
// 			const targetId = this.getAttribute('data-target');
// 			const submenu = document.querySelector(targetId);
// 			const menuItem = this.closest('.menu-item');
// 			if (!submenu) return;

// 			// Close all other submenus
// 			document.querySelectorAll('.has-subs').forEach(link => {
// 				if (link !== this) {
// 					link.classList.remove('active');
// 					link.closest('.menu-item')?.classList.remove('active');
// 					const otherSubmenu = document.querySelector(link.getAttribute('data-target'));
// 					if (otherSubmenu && otherSubmenu !== submenu) slideUp(otherSubmenu, 500);
// 				}
// 			});

// 			// Toggle the clicked submenu
// 			const isOpening = submenu.style.display === 'none' || getComputedStyle(submenu).display === 'none';
// 			if (isOpening) {
// 				slideDown(submenu, 500);
// 				this.classList.add('active');
// 				menuItem.classList.add('active');
// 			} else {
// 				slideUp(submenu, 500);
// 				this.classList.remove('active');
// 				menuItem.classList.remove('active');
// 			}
// 		});
// 	});

// 	//Reset normal links to avoid duplicate event listeners
// 	document.querySelectorAll('.menu-item > a:not(.has-subs)').forEach(link => {
// 		const newLink = link.cloneNode(true);
// 		link.parentNode.replaceChild(newLink, link);
// 	});

// 	//Add click event for normal menu links
// 	document.querySelectorAll('.menu-item > a:not(.has-subs)').forEach(link => {
// 		link.addEventListener('click', () => {
// 			document.querySelectorAll('.menu-item').forEach(i => i.classList.remove('active'));
// 			link.closest('.menu-item')?.classList.add('active');
// 		});
// 	});

// 	//Prevent submenu container clicks from bubbling up
// 	document.querySelectorAll('.submenu').forEach(submenu => {
// 		submenu.addEventListener('click', e => {
// 			if (!e.target.closest('a')) e.stopPropagation();
// 		});
// 	});
// }

// function changeText() {
// 	$(".wrapper-scroll-top").scroll(function () {
// 		$(".wrapper-scroll-bottom")
// 			.scrollLeft($(".wrapper-scroll-top").scrollLeft());
// 	});
// 	$(".wrapper-scroll-bottom").scroll(function () {
// 		$(".wrapper-scroll-top")
// 			.scrollLeft($(".wrapper-scroll-bottom").scrollLeft());
// 	});
// }

// function slideDown(element, duration = 300) {
// 	element.style.removeProperty('display');
// 	let display = window.getComputedStyle(element).display;

// 	if (display === 'none') display = 'block';
// 	element.style.display = display;

// 	const height = element.scrollHeight;

// 	element.style.overflow = 'hidden';
// 	element.style.height = '0';
// 	element.offsetHeight;

// 	element.style.transition = `height ${duration}ms ease`;
// 	element.style.height = height + 'px';

// 	window.setTimeout(() => {
// 		element.style.removeProperty('height');
// 		element.style.removeProperty('overflow');
// 		element.style.removeProperty('transition');
// 	}, duration);
// }

// function slideUp(element, duration = 300) {
// 	element.style.height = element.scrollHeight + 'px';
// 	element.style.overflow = 'hidden';
// 	element.offsetHeight;

// 	element.style.transition = `height ${duration}ms ease`;
// 	element.style.height = '0';

// 	window.setTimeout(() => {
// 		element.style.display = 'none';
// 		element.style.removeProperty('height');
// 		element.style.removeProperty('overflow');
// 		element.style.removeProperty('transition');
// 	}, duration);
// }

