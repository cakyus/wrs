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
    WScript.Echo("DEBUG: " + message);
  };

  this.error = function(message) {
    WScript.Echo("ERROR: " + message);
  };

  this.log = function(message) {
    WScript.Echo("LOG: " + message);
  };
}

// -- Bootstrap

var console = new Console;

main();

