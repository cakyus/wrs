console_debug("Hello World");

function console_debug(message) {
  WScript.Echo("DEBUG: " + message);
}

function console_error(message) {
  WScript.Echo("ERROR : " + message);
  WScript.Quit();
}

