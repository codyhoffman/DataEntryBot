// Copyright (c) 2012 The Chromium Authors. All rights reserved.
// Use of this source code is governed by a BSD-style license that can be
// found in the LICENSE file.

//*************************************************************************************
//NOTE: in regards to DataEntryBot ... only the JSON option is currently working
//NOTE: the rest of this code is from a demo file from SheetJS -- http://sheetjs.com
//NOTE: a utilities section has also been added at the bottom (not from sheetjs)
//*************************************************************************************


/*jshint browser:true */
/*global XLSX */
var X = XLSX;
var XW = {
	/* worker message */
	msg: 'xlsx',
	/* worker scripts */
	rABS: './xlsxworker2.js',
	norABS: './xlsxworker1.js',
	noxfer: './xlsxworker.js'
};

var rABS = typeof FileReader !== "undefined" && typeof FileReader.prototype !== "undefined" && typeof FileReader.prototype.readAsBinaryString !== "undefined";
if(!rABS) {
	document.getElementsByName("userabs")[0].disabled = true;
	document.getElementsByName("userabs")[0].checked = false;
}

var use_worker = typeof Worker !== 'undefined';
if(!use_worker) {
	document.getElementsByName("useworker")[0].disabled = true;
	document.getElementsByName("useworker")[0].checked = false;
}

var transferable = use_worker;
if(!transferable) {
	document.getElementsByName("xferable")[0].disabled = true;
	document.getElementsByName("xferable")[0].checked = false;
}

var wtf_mode = false;

function fixdata(data) {
	var o = "", l = 0, w = 10240;
	for(; l<data.byteLength/w; ++l) o+=String.fromCharCode.apply(null,new Uint8Array(data.slice(l*w,l*w+w)));
	o+=String.fromCharCode.apply(null, new Uint8Array(data.slice(l*w)));
	return o;
}

function ab2str(data) {
	var o = "", l = 0, w = 10240;
	for(; l<data.byteLength/w; ++l) o+=String.fromCharCode.apply(null,new Uint16Array(data.slice(l*w,l*w+w)));
	o+=String.fromCharCode.apply(null, new Uint16Array(data.slice(l*w)));
	return o;
}

function s2ab(s) {
	var b = new ArrayBuffer(s.length*2), v = new Uint16Array(b);
	for (var i=0; i != s.length; ++i) v[i] = s.charCodeAt(i);
	return [v, b];
}

function xw_noxfer(data, cb) {
	var worker = new Worker(XW.noxfer);
	worker.onmessage = function(e) {
		switch(e.data.t) {
			case 'ready': break;
			case 'e': console.error(e.data.d); break;
			case XW.msg: cb(JSON.parse(e.data.d)); break;
		}
	};
	var arr = rABS ? data : btoa(fixdata(data));
	worker.postMessage({d:arr,b:rABS});
}

function xw_xfer(data, cb) {
	var worker = new Worker(rABS ? XW.rABS : XW.norABS);
	worker.onmessage = function(e) {
		switch(e.data.t) {
			case 'ready': break;
			case 'e': console.error(e.data.d); break;
			default: xx=ab2str(e.data).replace(/\n/g,"\\n").replace(/\r/g,"\\r"); console.log("done"); cb(JSON.parse(xx)); break;
		}
	};
	if(rABS) {
		var val = s2ab(data);
		worker.postMessage(val[1], [val[1]]);
	} else {
		worker.postMessage(data, [data]);
	}
}

function xw(data, cb) {
	transferable = document.getElementsByName("xferable")[0].checked;
	if(transferable) xw_xfer(data, cb);
	else xw_noxfer(data, cb);
}

function get_radio_value( radioName ) {
	var radios = document.getElementsByName( radioName );
	for( var i = 0; i < radios.length; i++ ) {
		if( radios[i].checked || radios.length === 1 ) {
			return radios[i].value;
		}
	}
}



//*************************************************************************************
//NOTE: this JSON section and utilities below are the only portions of this code that have been altered for DataEntryBot 
//*************************************************************************************

function to_json(workbook) {
    var result = {};

    //Simulates a keypress event to pass the validation of ng-keyup
    var triggerKeyDown = function (element, keyCode) {
        var e = $.Event("keydown");
        e.which = keyCode;
        element.trigger(e);
    };

	workbook.SheetNames.forEach(function(sheetName) {
		var iframeName = "";
		var roa = X.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
		if(roa.length > 0){
			result[sheetName] = roa;
			var rows = roa;
			for (var i = 0; i < rows.length; i++) {
				var columns = rows[i];
				console.log(columns);
				for(var columnName in columns){
				   var elementId = "";
				   if(columnName.includes(".")) {
					   var splitString = columnName.split("."); //NOTE:  columnName can designate iframeName by:  iframeName.elementId
					   iframeName = splitString[0];
					   elementId = splitString[1];
					   console.log("iframeName: " + iframeName + " | elementId: " + elementId);
				   } else {
					   elementId = columnName;
					   console.log("elementId: " + elementId);
				   }
				   
					if(columns.hasOwnProperty(columnName)) {					   
					   var value = addslashes(columns[columnName].trim()); //NOTE: this escapes single and double quotes to prevent a string from accidentally ending itself
					   if(value.startsWith("#")) {
						   //NOTE: identified as a Header Row  (for improved human readability)
						   console.log("skipping Header Row from:"  + value);
							break;  //NOTE: skip to next row
					   }
					   if(value.startsWith("!")) {
						   //NOTE: identified as an optional Frame-Specifying Row  
						   iframeName = value.substring(1);  //NOTE: set specific iframe, removing the "!" designator
						   console.log("specifying new iframe:"  + iframeName);
						   break;  //NOTE: skip to next row
					   }
					   
					   if(elementId == "click") {
						   //TODO: fix and use UnableToLocateId() function
						   console.log("clicking: " + value);
						   if(iframeName != "") {
							   chrome.tabs.executeScript(null, {code:"document.querySelector('iframe[name=" + iframeName + "]').contentDocument.getElementById('" + value + "').click();"}); //NOTE: for click columns, the cell value is the element id
						   } else {
							   chrome.tabs.executeScript(null, {code:"document.getElementById('" + value + "').click();"}); //NOTE: for click columns, the Excel cell value is the element id
						   }
						   wait(3000); //TODO: replace this with an intelligent wait/pause that automatically ends once next page is fully loaded or AJAX is fully caught up
					   } else {
						   //TODO: fix and use UnableToLocateId() function
                           if (iframeName != "") {
                               var scriptString1 = "var element = document.querySelector('iframe[name=" + iframeName + "]').contentDocument.getElementById('" + elementId + "')";
                           } else {
                               var scriptString1 = "var element = document.getElementById('" + elementId + "')";
                           }
                           var scriptString2 = "element.setAttribute('novalidate', 'novalidate'); element.value = '" + value + "'; element.blur(); element.click(); element.focus(); element.keydown(); element.keypress(); element.keyup(); ";
                           
                           var keyPress = "element.dispatchEvent(new Event('input', {bubbles: true}));"


                           console.log(scriptString1);
                           chrome.tabs.executeScript(null, { code: scriptString1 });
						   console.log(scriptString2);
                           chrome.tabs.executeScript(null, { code: scriptString2 });
                           console.log(keyPress);
                           chrome.tabs.executeScript(null, { code: keyPress });
                           
					   }
					}
				}
			}
		}
	});
	return result;
}





function to_csv(workbook) {
	var result = [];
	workbook.SheetNames.forEach(function(sheetName) {
		var csv = X.utils.sheet_to_csv(workbook.Sheets[sheetName]);
		if(csv.length > 0){
			result.push("SHEET: " + sheetName);
			result.push("");
			result.push(csv);
		}
	});
	return result.join("\n");
}

function to_formulae(workbook) {
	var result = [];
	workbook.SheetNames.forEach(function(sheetName) {
		var formulae = X.utils.get_formulae(workbook.Sheets[sheetName]);
		if(formulae.length > 0){
			result.push("SHEET: " + sheetName);
			result.push("");
			result.push(formulae.join("\n"));
		}
	});
	return result.join("\n");
}

function to_html(workbook) {
	document.getElementById('htmlout').innerHTML = "";
	var result = [];
	workbook.SheetNames.forEach(function(sheetName) {
		var htmlstr = X.write(workbook, {sheet:sheetName, type:'binary', bookType:'html'});
		document.getElementById('htmlout').innerHTML += htmlstr;
	});
}

var tarea = document.getElementById('b64data');
function b64it() {
	if(typeof console !== 'undefined') console.log("onload", new Date());
	var wb = X.read(tarea.value, {type: 'base64',WTF:wtf_mode});
	process_wb(wb);
}

var global_wb;
function process_wb(wb) {
	global_wb = wb;
	var output = "";
	switch(get_radio_value("format")) {
		case "json":
			output = JSON.stringify(to_json(wb), 2, 2);
			break;
		case "form":
			output = to_formulae(wb);
			break;
		case "html": return to_html(wb);
		default:
			output = to_csv(wb);
	}
	if(out.innerText === undefined) out.textContent = output;
	else out.innerText = output;
	if(typeof console !== 'undefined') console.log("output", new Date());
}
function setfmt() {if(global_wb) process_wb(global_wb); }

var drop = document.getElementById('drop');
function handleDrop(e) {
	e.stopPropagation();
	e.preventDefault();
	rABS = document.getElementsByName("userabs")[0].checked;
	use_worker = document.getElementsByName("useworker")[0].checked;
	var files = e.dataTransfer.files;
	var f = files[0];
	{
		var reader = new FileReader();
		var name = f.name;
		reader.onload = function(e) {
			if(typeof console !== 'undefined') console.log("onload", new Date(), rABS, use_worker);
			var data = e.target.result;
			if(use_worker) {
				xw(data, process_wb);
			} else {
				var wb;
				if(rABS) {
					wb = X.read(data, {type: 'binary'});
				} else {
					var arr = fixdata(data);
					wb = X.read(btoa(arr), {type: 'base64'});
				}
				process_wb(wb);
			}
		};
		if(rABS) reader.readAsBinaryString(f);
		else reader.readAsArrayBuffer(f);
	}
}

function handleDragover(e) {
	e.stopPropagation();
	e.preventDefault();
	e.dataTransfer.dropEffect = 'copy';
}

if(drop.addEventListener) {
	drop.addEventListener('dragenter', handleDragover, false);
	drop.addEventListener('dragover', handleDragover, false);
	drop.addEventListener('drop', handleDrop, false);
}


var xlf = document.getElementById('xlf');
function handleFile(e) {
	rABS = document.getElementsByName("userabs")[0].checked;
	use_worker = document.getElementsByName("useworker")[0].checked;
	var files = e.target.files;
	var f = files[0];
	{
		var reader = new FileReader();
		var name = f.name;
		reader.onload = function(e) {
			if(typeof console !== 'undefined') console.log("onload", new Date(), rABS, use_worker);
			var data = e.target.result;
			if(use_worker) {
				xw(data, process_wb);
			} else {
				var wb;
				if(rABS) {
					wb = X.read(data, {type: 'binary'});
				} else {
					var arr = fixdata(data);
					wb = X.read(btoa(arr), {type: 'base64'});
				}
				process_wb(wb);
			}
		};
		if(rABS) reader.readAsBinaryString(f);
		else reader.readAsArrayBuffer(f);
	}
}

if(xlf.addEventListener) xlf.addEventListener('change', handleFile, false);



// *****************
// *** Utilities ***
// *****************

function wait(ms)
{
	//SOURCE: http://www.endmemo.com/js/pause.php
	var d = new Date();
	var d2 = null;
	do { d2 = new Date(); }
	while(d2-d < ms);
}

function IsEmpty(Value) {
    var isEmpty =
    ((typeof Value === 'undefined')
        || (Value === null)
        || (Value.trim() === '')
        ) ? true : false;
    //console.log("isEmpty('" + Value  + "') ==> " + isEmpty);
    return isEmpty;
}

function UnableToLocateId(id) {
	//TODO: fix
    if (IsEmpty(id))
    {
        return true;
    }
    else
    {
        return (id.indexOf("ERROR") == 0);
    }  
}

function addslashes (str) {
    return (str+'').replace(/[\\"']/g, '\\$&').replace(/\u0000/g, '\\0');
}