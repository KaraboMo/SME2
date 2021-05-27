/* eslint-disable */
"use strict";
function OnPostAssignMapping(maptype) {
  mapTypeArrray = ['Mapping', 'Group1', 'Group2', 'Group3', 'Group4', 'Group5',
                  'Group6', 'Group7', 'Group8', 'Group9', 'Group10'];
  try {
    //var app = new ActiveXObject("CaseWare.Application");
    //var sProgramPath = app.ApplicationInfo("ProgramPath");
    include(Client.FilePath + "Script\\CQSStats.js");
    registereventLib("ASSIGNMAPPIN", mapTypeArrray[maptype]);
    // app = null;
  }
  catch(e) {
  }
}
