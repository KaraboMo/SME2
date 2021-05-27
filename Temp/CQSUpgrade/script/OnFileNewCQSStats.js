/* eslint-disable */
"use strict";
function OnFileNew() {
  //debugger;
  try {
    //var app = new ActiveXObject("CaseWare.Application");
    //var sProgramPath = app.ApplicationInfo("ProgramPath");
    include(Client.FilePath + "Script\\CQSStats.js");
    registereventLib("NEW", "");
    // app = null;
  }
  catch(e) {
  }
}
