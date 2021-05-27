/* eslint-disable */
"use strict";
function OnFileImport(impType, impPath, impVersion, impGL) {
  try {
    //var app = new ActiveXObject("CaseWare.Application");
    //var sProgramPath = app.ApplicationInfo("ProgramPath");
    include(Client.FilePath + "Script\\CQSStats.js");
    registereventLib("ONFILEIMPORT", impType + ' (' + impVersion + ')' + ' GL: ' + impGL);
    // app = null;
  }
  catch(e) {
  }
}
