// @link https://documentation.help/MS-Office-JScript/jstutor.htm

ppt_files = get_ppt_files();

for (i in ppt_files) {
  f = ppt_files[i];
  console_debug(f.Name);
}

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

function console_debug(message) {
  WScript.Echo("DEBUG: " + message);
}

function console_error(message) {
  WScript.Echo("ERROR : " + message);
  WScript.Quit();
}

