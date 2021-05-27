/* eslint-disable */
"use strict";
function OnCopyMapping (templateFile) {
  try {
    //var app = new ActiveXObject("CaseWare.Application");
    //var sProgramPath = app.ApplicationInfo("ProgramPath");
    include(Client.FilePath + "Script\\CQSStats.js");
    registereventLib("ONCOPYMAPPING", "");
    // app = null;
  }
  catch(e) {
  }
}
