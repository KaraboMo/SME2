/* eslint-disable */
"use strict";
function OnAutoMap(source) {
  try {
    //var app = new ActiveXObject("CaseWare.Application");
    //var sProgramPath = app.ApplicationInfo("ProgramPath");
    include(Client.FilePath + "Script\\CQSStats.js");
    registereventLib("AUTOMAP", source);
    // app = null;
  }
  catch(e) {
  }
}
