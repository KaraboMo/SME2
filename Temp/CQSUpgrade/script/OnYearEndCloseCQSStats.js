/* eslint-disable */
"use strict";
function OnYearEndClose(location, lockdown) {
  lockdownarray = ['FALSE', 'TRUE'];
  try {
    //var app = new ActiveXObject("CaseWare.Application");
    //var sProgramPath = app.ApplicationInfo("ProgramPath");
    include(Client.FilePath + "Script\\CQSStats.js");
    registereventLib("YEC", lockdownarray[lockdown]);
    // app = null;
  }
  catch(e) {
  }
}
