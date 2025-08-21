// @link https://documentation.help/MS-Office-JScript/jstutor.htm


function main() {
  console.log('Hello World');
}


// ppt_files = get_ppt_files();

// for (i in ppt_files) {
//   f = ppt_files[i];
//   console_debug(f.Name);
// }

// Get all ppt files in script directory
// @return array

function get_ppt_files() {
  fso = new ActiveXObject("Scripting.FileSystemObject");
  dir_path = fso.GetParentFolderName( WScript.ScriptFullName );
  dir = fso.GetFolder(dir_path);
  dir_files = collection2array(dir.Files);
  // filter .pptx
  files = new Array();
  for (i in dir_files) {
    file = dir_files[i];
    if (file.Name.match(/\.pptx$/)) {
      files.push(file);
    }
  }
  return files;
}

// Get coresponding Excel file from PowerPoint file
// @return File|bool

function get_xls_file(ppt_file) {

}

function collection2array(collection) {
  e = new Enumerator(collection);
  items = new Array();
  while ( ! e.atEnd() ) {
    items.push(e.item());
    e.moveNext();
  }
  return items;
}

// -- Emulate Javacript Standard Library --

function Console() {

  this.debug = function(message) {
    WScript.Echo(date('H:i:s') + ' DEBUG ' + message);
  };

  this.error = function(message) {
    WScript.Echo(date('H:i:s') + ' ERROR ' + message);
    WScript.Quit();
  };

  this.log = function(message) {
    WScript.Echo(date('H:i:s') + " LOG " + message);
  };
}

// -- libphp.js --

var STR_PAD_LEFT = 0;
var STR_PAD_BOTH = 1;
var STR_PAD_RIGHT = 2;

function str_pad(string, length, pad_string, pad_type) {
  if (typeof(string) == 'number') {
    string = string.toString();
  }
  while (string.length < length) {
    if (pad_type == STR_PAD_LEFT) {
      string = pad_string + string;
    } else {
      string = string + pad_string;
    }
  }
  return string;
}

// function php_date(pattern, time) {
function date(pattern) {

  if (arguments.length == 1) {
    var time = new Date();
  }

  items = pattern.split('');
  for (var i = 0; i < items.length; i++) {
    if (items[i] == 'Y') {
      items[i] = time.getYear();
    } else if (items[i] == 'm') {
      items[i] = str_pad(time.getMonth() + 1, 2, '0', STR_PAD_LEFT);
    } else if (items[i] == 'd') {
      items[i] = str_pad(time.getDay(), 2, '0', STR_PAD_LEFT);
    } else if (items[i] == 'H') {
      items[i] = str_pad(time.getHours(), 2, '0', STR_PAD_LEFT);
    } else if (items[i] == 'i') {
      items[i] = str_pad(time.getMinutes(), 2, '0', STR_PAD_LEFT);
    } else if (items[i] == 's') {
      items[i] = str_pad(time.getSeconds(), 2, '0', STR_PAD_LEFT);
    }
  }

  return items.join('');
}

// -- Bootstrap --

var console = new Console;

main();

