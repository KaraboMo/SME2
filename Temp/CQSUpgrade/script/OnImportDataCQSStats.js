/* eslint-disable */
"use strict";
function OnImportData(importType, importFile) {
  try {
    //var app = new ActiveXObject("CaseWare.Application");
    //var sProgramPath = app.ApplicationInfo("ProgramPath");
    include(Client.FilePath + "Script\\CQSStats.js");
    registereventLib("ONIMPORTDATA", importType);
    // app = null;
  }
  catch(e) {
  }
}
